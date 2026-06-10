import {
  PptxRenderer,
  buildZip,
  degreesToOoxml,
  emuToPx,
  extractZip,
  getAllShapes,
  getShapeTransform,
  getSlideScale,
  ooxmlToDegrees,
  pxToEmu
} from 'pptx-svg';
import type { ShapeTransform } from 'pptx-svg';
import wasmBytes from 'pptx-svg/wasm';
import {
  getChartDataDescriptor,
  updateChartTextLabel as patchChartTextLabel,
  updateChartData as patchChartData,
  type ChartDataDescriptor,
  type ChartDataGrid,
  type ChartDataUpdate
} from './ChartData';
import { FontFidelity, type FontSubstitution } from './FontFidelity';
import {
  createSlideObjectClipboard,
  pasteSlideObject,
  type SlideObjectClipboard
} from './ShapeClipboard';

function assertOk(result: string, fallback: string): void {
  if (result.startsWith('ERROR:')) {
    throw new Error(result.slice('ERROR:'.length).trim() || fallback);
  }
}

const SLIDE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const DRAWINGML_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SHAPE_ELEMENT_NAMES = new Set(['cxnSp', 'graphicFrame', 'grpSp', 'pic', 'sp']);

export type GeneratedTextKind = 'chart' | 'table';

export interface GeneratedTextEdit {
  kind: GeneratedTextKind;
  labelIndex: number;
  occurrence: number;
  previousText: string;
}

type ChartAxisOrientation = 'horizontal' | 'vertical';

interface ChartAxisFormat {
  orientation: ChartAxisOrientation;
  formatCode: string;
  min: number;
  max: number;
  majorUnit: number | null;
  date1904: boolean;
}

interface ChartTickRun {
  orientation: ChartAxisOrientation;
  elements: SVGTextElement[];
}

function normalizeLabelText(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

function parseXml(contents: string, partPath: string): XMLDocument {
  const doc = new DOMParser().parseFromString(contents, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`Could not parse PowerPoint XML part: ${partPath}`);
  }
  return doc;
}

function serializeXml(doc: XMLDocument): string {
  return new XMLSerializer().serializeToString(doc);
}

function getDescendants(element: Element | XMLDocument, localName: string): Element[] {
  return Array.from(element.getElementsByTagNameNS('*', localName));
}

function getElementChildren(element: Element | undefined): Element[] {
  return Array.from(element?.childNodes ?? [])
    .filter((node): node is Element => node.nodeType === 1);
}

function getShapeElement(slideDoc: XMLDocument, shapeIndex: number): Element {
  const shapeTree = getDescendants(slideDoc, 'spTree')[0];
  const shape = getElementChildren(shapeTree)
    .filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName))[shapeIndex];
  if (!shape) {
    throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
  }
  return shape;
}

function setDrawingText(container: Element, text: string): void {
  const textElements = getDescendants(container, 't')
    .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  const firstText = textElements[0];
  if (!firstText) {
    throw new Error('This PowerPoint label has no editable text node.');
  }

  firstText.textContent = text;
  for (const element of textElements.slice(1)) {
    element.textContent = '';
  }
}

function getDrawingParagraphs(container: Element): Element[] {
  const textBody = getDescendants(container, 'txBody')
    .find((element) => element.namespaceURI === DRAWINGML_NAMESPACE || element.namespaceURI === container.namespaceURI);
  const scope = textBody ?? container;
  return getElementChildren(scope)
    .filter((element) => element.localName === 'p' && element.namespaceURI === DRAWINGML_NAMESPACE);
}

function setDrawingTextRun(container: Element, paragraphIndex: number, runIndex: number, text: string): void {
  const paragraphs = getDrawingParagraphs(container);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new Error('Could not find the selected text paragraph.');
  }

  const runs = getElementChildren(paragraph)
    .filter((element) => element.localName === 'r' && element.namespaceURI === DRAWINGML_NAMESPACE);
  const run = runs[runIndex];
  if (!run) {
    throw new Error('Could not find the selected text run.');
  }

  const textElement = getElementChildren(run)
    .find((element) => element.localName === 't' && element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? getDescendants(run, 't').find((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  if (!textElement) {
    throw new Error('Could not find the selected text node.');
  }

  textElement.textContent = text;
}

function resolvePartPath(sourcePath: string, target: string): string {
  const parts = sourcePath.split('/');
  parts.pop();

  for (const targetPart of target.replace(/\\/g, '/').split('/')) {
    if (!targetPart || targetPart === '.') continue;
    if (targetPart === '..') {
      parts.pop();
    } else {
      parts.push(targetPart);
    }
  }

  return parts.join('/');
}

function getSlidePath(slideIndex: number): string {
  return `ppt/slides/slide${slideIndex + 1}.xml`;
}

function getSlideRelationshipsPath(slideIndex: number): string {
  return `ppt/slides/_rels/slide${slideIndex + 1}.xml.rels`;
}

function findChartPartPath(
  textFiles: Map<string, string>,
  slideIndex: number,
  shapeIndex: number
): string {
  const slidePath = getSlidePath(slideIndex);
  const slideXml = textFiles.get(slidePath);
  if (!slideXml) {
    throw new Error(`Missing slide XML part: ${slidePath}`);
  }

  const shape = getShapeElement(parseXml(slideXml, slidePath), shapeIndex);
  const chart = getDescendants(shape, 'chart')[0];
  const relationshipId =
    chart?.getAttributeNS(RELATIONSHIP_NAMESPACE, 'id') ||
    chart?.getAttribute('r:id');
  if (!relationshipId) {
    throw new Error('Could not find the embedded chart relationship.');
  }

  const relationshipsPath = getSlideRelationshipsPath(slideIndex);
  const relationshipsXml = textFiles.get(relationshipsPath);
  if (!relationshipsXml) {
    throw new Error(`Missing slide relationship part: ${relationshipsPath}`);
  }

  const relationships = getDescendants(parseXml(relationshipsXml, relationshipsPath), 'Relationship');
  const relationship = relationships.find((element) => element.getAttribute('Id') === relationshipId);
  const target = relationship?.getAttribute('Target');
  if (!target || relationship?.getAttribute('TargetMode') === 'External') {
    throw new Error('Could not resolve the embedded chart XML part.');
  }

  return resolvePartPath(slidePath, target);
}

function hasAncestor(element: Element, localNames: Set<string>): boolean {
  let current = element.parentElement;
  while (current) {
    if (localNames.has(current.localName)) return true;
    current = current.parentElement;
  }
  return false;
}

function getChartTextSources(chartDoc: XMLDocument): Element[] {
  const richText = getDescendants(chartDoc, 't')
    .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  const cachedTextContainers = new Set(['strCache', 'strLit']);
  const cachedText = getDescendants(chartDoc, 'v').filter((element) => {
    return element.parentElement?.localName === 'tx' || hasAncestor(element, cachedTextContainers);
  });

  return [...richText, ...cachedText];
}

function getChartTextValues(textFiles: Map<string, string>, slideIndex: number, shapeIndex: number): string[] {
  const chartPath = findChartPartPath(textFiles, slideIndex, shapeIndex);
  const chartXml = textFiles.get(chartPath);
  if (!chartXml) {
    throw new Error(`Missing chart XML part: ${chartPath}`);
  }

  return getChartTextSources(parseXml(chartXml, chartPath))
    .map((element) => normalizeLabelText(element.textContent || ''))
    .filter(Boolean);
}

function getValAttribute(element: Element, localName: string): string | null {
  return getDescendants(element, localName)[0]?.getAttribute('val') ?? null;
}

function parseFiniteNumber(value: string | null): number | null {
  if (value === null || value.trim() === '') {
    return null;
  }

  const parsedValue = Number(value);
  return Number.isFinite(parsedValue) ? parsedValue : null;
}

function getCachedNumbers(chartDoc: XMLDocument, containerNames: string[]): number[] {
  const values: number[] = [];

  for (const containerName of containerNames) {
    for (const container of getDescendants(chartDoc, containerName)) {
      for (const point of getDescendants(container, 'pt')) {
        const value = parseFiniteNumber(
          getElementChildren(point).find((element) => element.localName === 'v')?.textContent ?? null
        );

        if (value !== null) {
          values.push(value);
        }
      }
    }
  }

  return values;
}

function getNiceStep(range: number): number {
  if (!Number.isFinite(range) || range <= 0) {
    return 1;
  }

  const roughStep = range / 5;
  const magnitude = 10 ** Math.floor(Math.log10(roughStep));
  const normalizedStep = roughStep / magnitude;

  if (normalizedStep <= 1) return magnitude;
  if (normalizedStep <= 2) return 2 * magnitude;
  if (normalizedStep <= 5) return 5 * magnitude;
  return 10 * magnitude;
}

function deriveAxisBounds(
  axisElement: Element,
  values: number[],
  includeZero: boolean
): { min: number; max: number; majorUnit: number | null } {
  const scaling = getDescendants(axisElement, 'scaling')[0] ?? axisElement;
  const explicitMin = parseFiniteNumber(getValAttribute(scaling, 'min'));
  const explicitMax = parseFiniteNumber(getValAttribute(scaling, 'max'));
  const explicitMajorUnit = parseFiniteNumber(getValAttribute(axisElement, 'majorUnit'));
  const fallbackMin = values.length > 0 ? Math.min(...values) : 0;
  const fallbackMax = values.length > 0 ? Math.max(...values) : 1;
  const dataMin = includeZero ? Math.min(0, fallbackMin) : fallbackMin;
  const dataMax = includeZero ? Math.max(0, fallbackMax) : fallbackMax;
  const step = explicitMajorUnit ?? getNiceStep(Math.max(dataMax - dataMin, Number.EPSILON));
  let min = explicitMin ?? (explicitMajorUnit === null ? dataMin : Math.floor(dataMin / step) * step);
  let max = explicitMax ?? (explicitMajorUnit === null ? dataMax : Math.ceil(dataMax / step) * step);

  if (!Number.isFinite(min)) {
    min = 0;
  }

  if (!Number.isFinite(max) || max <= min) {
    max = min + step * 5;
  }

  return { min, max, majorUnit: explicitMajorUnit };
}

function getChartAxisFormats(textFiles: Map<string, string>, slideIndex: number, shapeIndex: number): ChartAxisFormat[] {
  const chartPath = findChartPartPath(textFiles, slideIndex, shapeIndex);
  const chartXml = textFiles.get(chartPath);
  if (!chartXml) {
    throw new Error(`Missing chart XML part: ${chartPath}`);
  }

  const chartDoc = parseXml(chartXml, chartPath);
  const date1904 = getValAttribute(chartDoc.documentElement, 'date1904') === '1';
  const xValues = getCachedNumbers(chartDoc, ['xVal']);
  const categoryValues = getCachedNumbers(chartDoc, ['cat']);
  const yValues = getCachedNumbers(chartDoc, ['yVal']);
  const seriesValues = getCachedNumbers(chartDoc, ['val']);
  const formats: ChartAxisFormat[] = [];

  for (const axisName of ['valAx', 'dateAx']) {
    for (const axisElement of getDescendants(chartDoc, axisName)) {
      const axisPosition = getValAttribute(axisElement, 'axPos');
      const orientation: ChartAxisOrientation =
        axisPosition === 'l' || axisPosition === 'r' ? 'vertical' : 'horizontal';
      let values: number[];

      if (orientation === 'horizontal') {
        values = xValues.length > 0 ? xValues : categoryValues;
        if (values.length === 0 && axisName === 'valAx') {
          values = seriesValues;
        }
      } else {
        values = yValues.length > 0 ? yValues : seriesValues;
        if (values.length === 0 && axisName === 'dateAx') {
          values = categoryValues;
        }
      }

      const bounds = deriveAxisBounds(axisElement, values, axisName === 'valAx');
      formats.push({
        orientation,
        formatCode: getDescendants(axisElement, 'numFmt')[0]?.getAttribute('formatCode') ?? 'General',
        min: bounds.min,
        max: bounds.max,
        majorUnit: bounds.majorUnit,
        date1904
      });
    }
  }

  return formats;
}

function getDecimalPlaces(step: number): number {
  if (!Number.isFinite(step) || step <= 0) {
    return 0;
  }

  return Math.min(10, Math.max(0, Math.ceil(-Math.log10(step) - 1e-10)));
}

function formatFixedNumber(value: number, decimalPlaces: number, useThousandsSeparator: boolean): string {
  const threshold = 0.5 * 10 ** -decimalPlaces;
  const normalizedValue = Math.abs(value) < threshold ? 0 : value;
  const fixedValue = normalizedValue.toFixed(decimalPlaces);
  const [integerPart = '0', decimalPart] = fixedValue.split('.');
  const formattedInteger = useThousandsSeparator
    ? integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
    : integerPart;

  return decimalPart === undefined ? formattedInteger : `${formattedInteger}.${decimalPart}`;
}

function formatGeneralNumber(value: number, step: number): string {
  if (value !== 0 && (Math.abs(value) >= 1e12 || Math.abs(value) < 1e-8)) {
    return value.toExponential(Math.max(0, getDecimalPlaces(step))).replace('e', 'E');
  }

  return formatFixedNumber(value, getDecimalPlaces(step), false)
    .replace(/(\.\d*?)0+$/, '$1')
    .replace(/\.$/, '');
}

function excelSerialToDate(serial: number, date1904: boolean): Date {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return new Date(epoch + Math.round(serial * 24 * 60 * 60 * 1000));
}

function formatExcelDate(serial: number, formatCode: string, date1904: boolean): string {
  const date = excelSerialToDate(serial, date1904);
  const months = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ];
  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();
  const year = date.getUTCFullYear();
  const monthName = months[month - 1] ?? '';

  return (formatCode.split(';')[0] ?? formatCode)
    .replace(/\[[^\]]+\]/g, '')
    .replace(/\\(.)/g, '$1')
    .replace(/"([^"]*)"/g, '$1')
    .replace(/yyyy|yy|mmmm|mmm|mm|m|dd|d/gi, (token) => {
      switch (token.toLowerCase()) {
        case 'yyyy':
          return String(year);
        case 'yy':
          return String(year).slice(-2);
        case 'mmmm':
          return monthName;
        case 'mmm':
          return monthName.slice(0, 3);
        case 'mm':
          return String(month).padStart(2, '0');
        case 'm':
          return String(month);
        case 'dd':
          return String(day).padStart(2, '0');
        default:
          return String(day);
      }
    });
}

export function formatChartAxisValue(
  value: number,
  formatCode: string,
  step: number,
  date1904 = false
): string {
  const primaryFormat = formatCode.split(';')[0] || 'General';
  const normalizedFormat = primaryFormat.replace(/\[[^\]]+\]/g, '').replace(/"[^"]*"/g, '');

  if (/[dmy]/i.test(normalizedFormat) && !/[#0](?:\.[#0]+)?%?/i.test(normalizedFormat)) {
    return formatExcelDate(value, primaryFormat, date1904);
  }

  const isPercentage = normalizedFormat.includes('%');
  const scaledValue = isPercentage ? value * 100 : value;
  const scaledStep = isPercentage ? step * 100 : step;
  const decimalMatch = normalizedFormat.match(/\.([0#]+)/);
  const useThousandsSeparator = /[0#],[0#]/.test(normalizedFormat);
  let formattedValue: string;

  if (/E[+-]?0+/i.test(normalizedFormat)) {
    formattedValue = scaledValue
      .toExponential(decimalMatch?.[1]?.length ?? 0)
      .replace('e', 'E')
      .replace(/E(\d)/, 'E+$1');
  } else if (decimalMatch !== null || /[0#]/.test(normalizedFormat) && normalizedFormat !== 'General') {
    formattedValue = formatFixedNumber(scaledValue, decimalMatch?.[1]?.length ?? 0, useThousandsSeparator);
  } else {
    formattedValue = formatGeneralNumber(scaledValue, scaledStep);
  }

  return isPercentage ? `${formattedValue}%` : formattedValue;
}

function getChartTickRuns(chartGroup: Element): ChartTickRun[] {
  const runs: ChartTickRun[] = [];
  let currentRun: ChartTickRun | null = null;
  let currentKey: string | null = null;
  let previousPosition: number | null = null;

  const finishRun = (): void => {
    if (currentRun !== null && currentRun.elements.length >= 2) {
      runs.push(currentRun);
    }

    currentRun = null;
    currentKey = null;
    previousPosition = null;
  };

  for (const textElement of getDescendants(chartGroup, 'text') as SVGTextElement[]) {
    if (textElement.getAttribute('fill')?.toLowerCase() !== '#666666') {
      finishRun();
      continue;
    }

    const anchor = textElement.getAttribute('text-anchor');
    const x = textElement.getAttribute('x');
    const y = textElement.getAttribute('y');

    if (x === null || y === null || anchor !== 'middle' && anchor !== 'end') {
      finishRun();
      continue;
    }

    const orientation: ChartAxisOrientation = anchor === 'middle' ? 'horizontal' : 'vertical';
    const key = orientation === 'horizontal' ? `h:${y}` : `v:${x}`;
    const position = Number(orientation === 'horizontal' ? x : y);
    const hasReset =
      previousPosition !== null &&
      Number.isFinite(position) &&
      (orientation === 'horizontal' ? position <= previousPosition : position >= previousPosition);

    if (key !== currentKey || hasReset) {
      finishRun();
      currentRun = { orientation, elements: [] };
      currentKey = key;
    }

    currentRun?.elements.push(textElement);
    previousPosition = position;
  }

  finishRun();
  return runs;
}

function removeRedundantTickRuns(runs: ChartTickRun[], axis: ChartAxisFormat): ChartTickRun[] {
  if (runs.length < 2) {
    return runs;
  }

  const keptRuns: ChartTickRun[] = [];

  for (const run of runs) {
    const equivalentIndex = keptRuns.findIndex((keptRun) => {
      if (keptRun.orientation !== run.orientation) {
        return false;
      }

      const keptFirst = keptRun.elements[0];
      const keptLast = keptRun.elements[keptRun.elements.length - 1];
      const first = run.elements[0];
      const last = run.elements[run.elements.length - 1];
      const coordinate = run.orientation === 'horizontal' ? 'x' : 'y';

      if (!keptFirst || !keptLast || !first || !last) {
        return false;
      }

      return (
        keptFirst.getAttribute(coordinate) === first.getAttribute(coordinate) &&
        keptLast.getAttribute(coordinate) === last.getAttribute(coordinate)
      );
    });

    if (equivalentIndex === -1) {
      keptRuns.push(run);
      continue;
    }

    const keptRun = keptRuns[equivalentIndex];
    if (!keptRun) {
      keptRuns.push(run);
      continue;
    }

    const expectedCount =
      axis.majorUnit === null ? null : Math.round((axis.max - axis.min) / axis.majorUnit) + 1;
    const shouldReplace =
      expectedCount !== null &&
      Math.abs(run.elements.length - expectedCount) < Math.abs(keptRun.elements.length - expectedCount);
    const redundantRun = shouldReplace ? keptRun : run;

    for (const element of redundantRun.elements) {
      element.parentNode?.removeChild(element);
    }

    if (shouldReplace) {
      keptRuns[equivalentIndex] = run;
    }
  }

  return keptRuns;
}

async function normalizeSlideManifest(buffer: ArrayBuffer, slideCount: number): Promise<ArrayBuffer> {
  const zip = await extractZip(buffer);
  const presentationPath = 'ppt/presentation.xml';
  const relationshipsPath = 'ppt/_rels/presentation.xml.rels';
  const contentTypesPath = '[Content_Types].xml';
  const presentation = zip.textFiles.get(presentationPath);
  const relationships = zip.textFiles.get(relationshipsPath);
  const contentTypes = zip.textFiles.get(contentTypesPath);
  if (!presentation || !relationships || !contentTypes) {
    throw new Error('Cannot normalize slide metadata because required OOXML parts are missing.');
  }

  let nextRelationshipId = 1;
  for (const match of relationships.matchAll(/\bId="rId(\d+)"/g)) {
    nextRelationshipId = Math.max(nextRelationshipId, Number(match[1]) + 1);
  }

  const slideIds = Array.from(presentation.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"[^>]*\/?>/g))
    .map((match) => Number(match[1]))
    .filter(Number.isFinite);
  let nextSlideId = Math.max(255, ...slideIds) + 1;
  const normalizedSlideEntries: string[] = [];
  const normalizedRelationships: string[] = [];

  for (let index = 0; index < slideCount; index++) {
    const relationshipId = `rId${nextRelationshipId++}`;
    const slideId = slideIds[index] ?? nextSlideId++;
    normalizedSlideEntries.push(`<p:sldId id="${slideId}" r:id="${relationshipId}"/>`);
    normalizedRelationships.push(
      `<Relationship Id="${relationshipId}" Type="${SLIDE_RELATIONSHIP_TYPE}" Target="slides/slide${index + 1}.xml"/>`
    );
  }

  const updatedPresentation = presentation.replace(
    /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    `<p:sldIdLst>${normalizedSlideEntries.join('')}</p:sldIdLst>`
  );
  const updatedRelationships = relationships
    .replace(
      new RegExp(`<Relationship\\b(?=[^>]*\\bType="${SLIDE_RELATIONSHIP_TYPE}")[^>]*/?>`, 'g'),
      ''
    )
    .replace('</Relationships>', `${normalizedRelationships.join('')}</Relationships>`);
  const slideOverrides = Array.from(
    { length: slideCount },
    (_, index) =>
      `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="${SLIDE_CONTENT_TYPE}"/>`
  ).join('');
  const updatedContentTypes = contentTypes
    .replace(/<Override\b(?=[^>]*\bPartName="\/ppt\/slides\/slide\d+\.xml")[^>]*\/>/g, '')
    .replace('</Types>', `${slideOverrides}</Types>`);

  const removals = new Set<string>();
  for (const path of [...zip.textFiles.keys(), ...zip.binaryFiles.keys()]) {
    const match = path.match(/^ppt\/slides\/(?:_rels\/)?slide(\d+)\.xml(?:\.rels)?$/);
    if (match && Number(match[1]) > slideCount) {
      removals.add(path);
    }
  }

  return buildZip(
    buffer,
    new Map([
      [presentationPath, updatedPresentation],
      [relationshipsPath, updatedRelationships],
      [contentTypesPath, updatedContentTypes]
    ]),
    removals
  );
}

async function preserveSlideExtensionLists(previousBuffer: ArrayBuffer, exportedBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const [previousZip, exportedZip] = await Promise.all([
    extractZip(previousBuffer),
    extractZip(exportedBuffer)
  ]);
  const modifications = new Map<string, string>();

  for (const [slidePath, exportedXml] of exportedZip.textFiles) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(slidePath)) continue;

    const previousXml = previousZip.textFiles.get(slidePath);
    if (!previousXml) continue;

    const mergedXml = preserveSlideExtensionList(previousXml, exportedXml, slidePath);
    if (mergedXml !== null && mergedXml !== exportedXml) {
      modifications.set(slidePath, mergedXml);
    }
  }

  return modifications.size > 0
    ? buildZip(exportedBuffer, modifications)
    : exportedBuffer;
}

function preserveSlideExtensionList(previousXml: string, exportedXml: string, slidePath: string): string | null {
  const previousDocument = parseXml(previousXml, slidePath);
  const exportedDocument = parseXml(exportedXml, slidePath);
  const previousCommonSlide = getDescendants(previousDocument, 'cSld')[0];
  const exportedCommonSlide = getDescendants(exportedDocument, 'cSld')[0];
  if (!previousCommonSlide || !exportedCommonSlide) return null;

  const previousExtensionList = getDirectChild(previousCommonSlide, 'extLst');
  if (!previousExtensionList) return null;

  const exportedExtensionList = getDirectChild(exportedCommonSlide, 'extLst');
  const importedExtensionList = exportedDocument.importNode(previousExtensionList, true);

  if (exportedExtensionList) {
    exportedCommonSlide.replaceChild(importedExtensionList, exportedExtensionList);
  } else {
    exportedCommonSlide.appendChild(importedExtensionList);
  }

  return serializeXml(exportedDocument);
}

function getDirectChild(element: Element, localName: string): Element | null {
  return getElementChildren(element)
    .find((child) => child.localName === localName) ?? null;
}

export interface RenderedSlide {
  svg: string;
  slideCount: number;
}

export interface SlideMoveResult {
  slideIndex: number;
  slideCount: number;
}

export class PresentationEngine {
  private renderer: PptxRenderer;
  private fontFidelity: FontFidelity;
  private currentBuffer: ArrayBuffer;
  private slideCountValue = 0;
  private chartTextValues = new Map<string, string[]>();
  private chartAxisFormats = new Map<string, ChartAxisFormat[]>();
  private chartDataDescriptors = new Map<string, ChartDataDescriptor>();

  private constructor(renderer: PptxRenderer, fontFidelity: FontFidelity, slideCount: number, buffer: ArrayBuffer) {
    this.renderer = renderer;
    this.fontFidelity = fontFidelity;
    this.currentBuffer = buffer.slice(0);
    this.slideCountValue = slideCount;
  }

  static async load(buffer: ArrayBuffer): Promise<PresentationEngine> {
    const { renderer, fontFidelity, slideCount } = await PresentationEngine.createRenderer(buffer);
    const engine = new PresentationEngine(renderer, fontFidelity, slideCount, buffer);
    await engine.refreshChartTextValues(buffer);
    return engine;
  }

  private static async createRenderer(buffer: ArrayBuffer): Promise<{
    renderer: PptxRenderer;
    fontFidelity: FontFidelity;
    slideCount: number;
  }> {
    const fontFidelity = new FontFidelity();
    const renderer = new PptxRenderer({
      logLevel: 'error',
      fontFallbacks: fontFidelity.getRendererFallbacks(),
      measureText: (text, fontFace, fontSizePx) => fontFidelity.measureText(text, fontFace, fontSizePx)
    });

    await renderer.init(wasmBytes);
    const { slideCount } = await renderer.loadPptx(buffer);
    return { renderer, fontFidelity, slideCount };
  }

  static async validateRoundTrip(buffer: ArrayBuffer, expectedSlideCount: number): Promise<void> {
    const engine = await PresentationEngine.load(buffer);
    if (engine.slideCount !== expectedSlideCount) {
      throw new Error(`Round-trip slide count mismatch: expected ${expectedSlideCount}, got ${engine.slideCount}.`);
    }

    if (engine.slideCount > 0) {
      engine.renderSlide(0);
    }
  }

  get slideCount(): number {
    return this.slideCountValue;
  }

  renderSlide(slideIndex: number): RenderedSlide {
    const svg = this.renderer.renderSlideSvg(slideIndex);
    assertOk(svg, 'Could not render slide.');
    return { svg, slideCount: this.slideCountValue };
  }

  renderShape(slideIndex: number, shapeIndex: number): string {
    const svg = this.renderer.renderShapeSvg(slideIndex, shapeIndex);
    assertOk(svg, 'Could not render shape.');
    return svg;
  }

  getShapes(svg: SVGSVGElement): SVGGElement[] {
    return getAllShapes(svg);
  }

  applyFontFidelity(svg: SVGSVGElement): FontSubstitution[] {
    return this.fontFidelity.applySvgSubstitutions(svg);
  }

  getChartDataGrid(slideIndex: number, shapeIndex: number): ChartDataGrid | null {
    return this.chartDataDescriptors.get(this.getChartTextKey(slideIndex, shapeIndex))?.grid ?? null;
  }

  async updateChartData(slideIndex: number, shapeIndex: number, update: ChartDataUpdate): Promise<void> {
    const descriptor = this.chartDataDescriptors.get(this.getChartTextKey(slideIndex, shapeIndex));
    if (!descriptor) {
      throw new Error('Could not find chart data for the selected object.');
    }

    const rawExport = await this.exportRendererState();
    const patchedExport = await patchChartData(rawExport, descriptor, update);
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  formatChartAxisLabels(svg: SVGSVGElement, slideIndex: number): void {
    for (const chartGroup of getDescendants(svg, 'g')) {
      if (chartGroup.getAttribute('data-ooxml-shape-type') !== 'chart') {
        continue;
      }

      const shapeIndex = Number(chartGroup.getAttribute('data-ooxml-shape-idx'));
      const formats = this.chartAxisFormats.get(this.getChartTextKey(slideIndex, shapeIndex));

      if (!Number.isInteger(shapeIndex) || formats === undefined) {
        continue;
      }

      for (const orientation of ['horizontal', 'vertical'] as const) {
        const axes = formats.filter((axis) => axis.orientation === orientation);
        const runs = getChartTickRuns(chartGroup).filter((run) => run.orientation === orientation);

        if (axes.length === 0 || runs.length === 0) {
          continue;
        }

        const defaultAxis = axes[0];
        if (!defaultAxis) {
          continue;
        }

        const visibleRuns = axes.length === 1 ? removeRedundantTickRuns(runs, defaultAxis) : runs;

        visibleRuns.forEach((run, index) => {
          const axis = axes[Math.min(index, axes.length - 1)] ?? defaultAxis;
          const step = run.elements.length > 1 ? (axis.max - axis.min) / (run.elements.length - 1) : 0;

          run.elements.forEach((element, tickIndex) => {
            element.textContent = formatChartAxisValue(
              axis.min + step * tickIndex,
              axis.formatCode,
              step,
              axis.date1904
            );
            element.setAttribute('data-native-powerpoint-axis-tick', 'true');
          });
        });
      }
    }
  }

  getShapeTransform(shape: SVGGElement): ShapeTransform {
    return getShapeTransform(shape);
  }

  getSlideScale(svg: SVGSVGElement): number {
    return getSlideScale(svg);
  }

  emuToPx(emu: number): number {
    return emuToPx(emu);
  }

  pxToEmu(px: number): number {
    return pxToEmu(px);
  }

  ooxmlToDegrees(value: number): number {
    return ooxmlToDegrees(value);
  }

  degreesToOoxml(value: number): number {
    return degreesToOoxml(value);
  }

  updateShapeTransform(slideIndex: number, shapeIndex: number, transform: ShapeTransform): void {
    const result = this.renderer.updateShapeTransform(
      slideIndex,
      shapeIndex,
      Math.round(transform.x),
      Math.round(transform.y),
      Math.max(1, Math.round(transform.cx)),
      Math.max(1, Math.round(transform.cy)),
      Math.round(transform.rot)
    );
    assertOk(result, 'Could not update shape transform.');
  }

  async updateShapeText(slideIndex: number, shapeIndex: number, text: string): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shape = getShapeElement(slideDoc, shapeIndex);
    const textElements = getDescendants(shape, 't')
      .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
    if (textElements.length > 0) {
      setDrawingText(shape, text);
      const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
      await this.reloadFromBuffer(patchedExport, this.slideCountValue);
      return;
    }

    const addResult = this.renderer.addShapeText(slideIndex, shapeIndex, text, 1800, 0, 0, 0);
    assertOk(addResult, 'Could not update shape text.');
  }

  async updateTextRun(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    runIndex: number,
    text: string
  ): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shape = getShapeElement(slideDoc, shapeIndex);
    setDrawingTextRun(shape, paragraphIndex, runIndex, text);
    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  canUpdateGeneratedText(slideIndex: number, shapeIndex: number, edit: GeneratedTextEdit): boolean {
    if (edit.kind === 'table') return true;
    const chartValues = this.chartTextValues.get(this.getChartTextKey(slideIndex, shapeIndex)) ?? [];
    return chartValues.includes(normalizeLabelText(edit.previousText));
  }

  async updateGeneratedText(slideIndex: number, shapeIndex: number, edit: GeneratedTextEdit, text: string): Promise<void> {
    const rawExport = await this.exportRendererState();
    const zip = await extractZip(rawExport);
    const modifications = new Map<string, string>();

    if (edit.kind === 'table') {
      const slidePath = getSlidePath(slideIndex);
      const slideXml = zip.textFiles.get(slidePath);
      if (!slideXml) {
        throw new Error(`Missing slide XML part: ${slidePath}`);
      }

      const slideDoc = parseXml(slideXml, slidePath);
      const shape = getShapeElement(slideDoc, shapeIndex);
      const table = getDescendants(shape, 'tbl')[0];
      const cell = table ? getDescendants(table, 'tc')[edit.labelIndex] : null;
      if (!cell) {
        throw new Error('Could not find the selected table cell.');
      }

      setDrawingText(cell, text);
      modifications.set(slidePath, serializeXml(slideDoc));
    } else {
      const descriptor = this.chartDataDescriptors.get(this.getChartTextKey(slideIndex, shapeIndex));
      if (descriptor) {
        const patchedExport = await patchChartTextLabel(rawExport, descriptor, edit.previousText, edit.occurrence, text);
        await this.reloadFromBuffer(patchedExport, this.slideCountValue);
        return;
      }

      const chartPath = findChartPartPath(zip.textFiles, slideIndex, shapeIndex);
      const chartXml = zip.textFiles.get(chartPath);
      if (!chartXml) {
        throw new Error(`Missing chart XML part: ${chartPath}`);
      }

      const chartDoc = parseXml(chartXml, chartPath);
      const previousText = normalizeLabelText(edit.previousText);
      const matches = getChartTextSources(chartDoc)
        .filter((element) => normalizeLabelText(element.textContent || '') === previousText);
      const source = matches[edit.occurrence] ?? matches[0];
      if (!source) {
        throw new Error('This chart label is generated from chart scale or numeric data and cannot be renamed directly.');
      }

      source.textContent = text;
      modifications.set(chartPath, serializeXml(chartDoc));
    }

    const patchedExport = await buildZip(rawExport, modifications);
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  addTextBox(slideIndex: number): number {
    const x = pxToEmu(180);
    const y = pxToEmu(120);
    const cx = pxToEmu(300);
    const cy = pxToEmu(80);
    const result = this.renderer.addShape(slideIndex, 'rect', x, y, cx, cy, -1, -1, -1);
    assertOk(result, 'Could not add text box.');

    const shapeIndex = Number(result.split(':')[1]);
    if (!Number.isFinite(shapeIndex)) {
      throw new Error('The renderer did not return a valid shape index.');
    }

    const textResult = this.renderer.addShapeText(slideIndex, shapeIndex, 'New text', 1800, -1, -1, -1);
    assertOk(textResult, 'Could not add text to the new text box.');
    return shapeIndex;
  }

  deleteShape(slideIndex: number, shapeIndex: number): void {
    const result = this.renderer.deleteShape(slideIndex, shapeIndex);
    assertOk(result, 'Could not delete shape.');
  }

  async copyShape(slideIndex: number, shapeIndex: number): Promise<SlideObjectClipboard> {
    return createSlideObjectClipboard(await this.exportRendererState(), slideIndex, shapeIndex);
  }

  async pasteShape(
    clipboard: SlideObjectClipboard,
    destinationSlideIndex: number
  ): Promise<number> {
    const rawExport = await this.exportRendererState();
    const result = await pasteSlideObject(rawExport, clipboard, destinationSlideIndex);
    await this.reloadFromBuffer(result.buffer, this.slideCountValue);
    return result.shapeIndex;
  }

  async duplicateShape(slideIndex: number, shapeIndex: number): Promise<number> {
    return this.pasteShape(await this.copyShape(slideIndex, shapeIndex), slideIndex);
  }

  async addSlide(afterIndex: number): Promise<SlideMoveResult> {
    const sourceIndex = Math.max(0, Math.min(afterIndex, this.slideCountValue - 1));
    const { slideCount, insertedIdx } = await this.renderer.addSlide(afterIndex, sourceIndex);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: insertedIdx, slideCount };
  }

  async deleteSlide(slideIndex: number): Promise<SlideMoveResult> {
    if (this.slideCountValue <= 1) {
      throw new Error('A presentation must keep at least one slide.');
    }

    const { slideCount } = await this.renderer.deleteSlide(slideIndex);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: Math.min(slideIndex, slideCount - 1), slideCount };
  }

  async moveSlide(slideIndex: number, direction: -1 | 1): Promise<SlideMoveResult> {
    const targetIndex = slideIndex + direction;
    if (targetIndex < 0 || targetIndex >= this.slideCountValue) {
      return { slideIndex, slideCount: this.slideCountValue };
    }

    const order = Array.from({ length: this.slideCountValue }, (_, index) => index);
    const [moved] = order.splice(slideIndex, 1);
    if (moved === undefined) {
      return { slideIndex, slideCount: this.slideCountValue };
    }
    order.splice(targetIndex, 0, moved);
    const { slideCount } = await this.renderer.reorderSlides(order);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: targetIndex, slideCount };
  }

  async export(): Promise<ArrayBuffer> {
    return this.exportRendererState();
  }

  async restoreSnapshot(buffer: ArrayBuffer): Promise<void> {
    const { renderer, fontFidelity, slideCount } = await PresentationEngine.createRenderer(buffer);
    this.renderer = renderer;
    this.fontFidelity = fontFidelity;
    this.currentBuffer = buffer.slice(0);
    this.slideCountValue = slideCount;
    await this.refreshChartTextValues(buffer);
  }

  private async reloadAfterSlideManagement(expectedSlideCount: number): Promise<void> {
    const rawExport = await this.exportRendererState();
    const normalizedExport = await normalizeSlideManifest(rawExport, expectedSlideCount);
    await this.reloadFromBuffer(normalizedExport, expectedSlideCount);
  }

  private async reloadFromBuffer(buffer: ArrayBuffer, expectedSlideCount: number): Promise<void> {
    const { renderer, fontFidelity, slideCount } = await PresentationEngine.createRenderer(buffer);
    if (slideCount !== expectedSlideCount) {
      throw new Error(`Slide management export mismatch: expected ${expectedSlideCount}, got ${slideCount}.`);
    }

    this.renderer = renderer;
    this.fontFidelity = fontFidelity;
    this.currentBuffer = buffer.slice(0);
    this.slideCountValue = slideCount;
    await this.refreshChartTextValues(buffer);
  }

  private async exportRendererState(): Promise<ArrayBuffer> {
    const rawExport = await this.renderer.exportPptx();
    const preservedExport = await preserveSlideExtensionLists(this.currentBuffer, rawExport);
    this.currentBuffer = preservedExport.slice(0);
    return preservedExport;
  }

  private async refreshChartTextValues(buffer: ArrayBuffer): Promise<void> {
    const zip = await extractZip(buffer);
    const chartTextValues = new Map<string, string[]>();
    const chartAxisFormats = new Map<string, ChartAxisFormat[]>();
    const chartDataDescriptors = new Map<string, ChartDataDescriptor>();

    for (let slideIndex = 0; slideIndex < this.slideCountValue; slideIndex++) {
      const slidePath = getSlidePath(slideIndex);
      const slideXml = zip.textFiles.get(slidePath);
      if (!slideXml) continue;

      const shapes = getElementChildren(getDescendants(parseXml(slideXml, slidePath), 'spTree')[0])
        .filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName));
      shapes.forEach((shape, shapeIndex) => {
        if (!getDescendants(shape, 'chart')[0]) return;
        try {
          const chartPath = findChartPartPath(zip.textFiles, slideIndex, shapeIndex);
          chartTextValues.set(
            this.getChartTextKey(slideIndex, shapeIndex),
            getChartTextValues(zip.textFiles, slideIndex, shapeIndex)
          );
          chartAxisFormats.set(
            this.getChartTextKey(slideIndex, shapeIndex),
            getChartAxisFormats(zip.textFiles, slideIndex, shapeIndex)
          );
          chartDataDescriptors.set(
            this.getChartTextKey(slideIndex, shapeIndex),
            getChartDataDescriptor(zip, chartPath)
          );
        } catch {
          // Unsupported chart variants remain viewable and read-only.
        }
      });
    }

    this.chartTextValues = chartTextValues;
    this.chartAxisFormats = chartAxisFormats;
    this.chartDataDescriptors = chartDataDescriptors;
  }

  private getChartTextKey(slideIndex: number, shapeIndex: number): string {
    return `${slideIndex}:${shapeIndex}`;
  }
}
