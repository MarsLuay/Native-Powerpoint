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
import {
  applyParagraphListStyle,
  insertChartIntoPresentation,
  insertTableIntoPresentation,
  mergeMissingPackageParts,
  mergeSlideGraphicFramesFromBuffer,
  type ParagraphListStyle
} from './SlideInsertions';

export type InsertableShapeGeometry =
  | 'rect'
  | 'ellipse'
  | 'roundRect'
  | 'line'
  | 'rightArrow'
  | 'leftArrow'
  | 'upArrow'
  | 'downArrow';

export type SlideLayoutKind = 'blank' | 'title' | 'titleBody';

/** Renderer augmented with the build-time `initJsBackend` patch (see esbuild.config.mjs). */
interface JsBackendCapableRenderer {
  initJsBackend(engine: unknown): void;
}

/**
 * True when a failure from the Wasm renderer indicates the runtime lacks
 * WebAssembly GC (Obsidian installer < 1.5.8 / Chromium < 119). Mirrors the
 * detection in NativePowerPointView so both layers agree.
 */
function isWasmGcUnsupportedError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error || '');
  return /WebAssembly GC support|Wasm init failed/i.test(message);
}

/**
 * Initialize the renderer's backend. Prefers the fast Wasm (wasm-gc) engine and,
 * if the runtime cannot run it, lazily loads the pure-JS engine fallback so PPTX
 * files still open on older Obsidian installers. The fallback module is only
 * fetched/evaluated when actually needed.
 */
async function initRendererBackend(renderer: PptxRenderer): Promise<void> {
  try {
    await renderer.init(wasmBytes);
  } catch (error) {
    if (!isWasmGcUnsupportedError(error)) throw error;
    const { createPptxJsEngine } = await import('./vendor/pptx-js-engine.mjs');
    (renderer as unknown as JsBackendCapableRenderer).initJsBackend(createPptxJsEngine());
  }
}

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
const PACKAGE_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const IMAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const SHAPE_ELEMENT_NAMES = new Set(['cxnSp', 'graphicFrame', 'grpSp', 'pic', 'sp']);

/** Inset crop, expressed as a percentage (0-100) of the source image edge. */
export interface ImageCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** Raw image bytes resolved from an embedded picture, with its MIME type. */
export interface ShapeImageData {
  bytes: Uint8Array;
  mimeType: string;
}

export type GeneratedTextKind = 'chart' | 'table';

export interface GeneratedTextEdit {
  kind: GeneratedTextKind;
  labelIndex: number;
  occurrence: number;
  previousText: string;
}

export type ParagraphAlignment = 'l' | 'ctr' | 'r' | 'just';

/** A specific text run inside a shape, identified by paragraph and run index. */
export interface RunTarget {
  paragraphIndex: number;
  runIndex: number;
}

/** A character range inside one DrawingML paragraph. */
export interface ParagraphTextRange {
  paragraphIndex: number;
  start: number;
  end: number;
}

/**
 * Requested run-level style changes. Omitted fields are left unchanged.
 * `color`/`highlight` use uppercase `RRGGBB` hex (no `#`); `null` clears them.
 */
export interface RunStyleChange {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string | null;
  highlight?: string | null;
}

/** Resolved run-level style read back from a slide, for reflecting toolbar state. */
export interface RunStyleInfo {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  fontFamily: string | null;
  fontSizePt: number | null;
  color: string | null;
  highlight: string | null;
  alignment: ParagraphAlignment | null;
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

interface ShapeBox {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

export type ShapeReorderMode = 'front' | 'back' | 'forward' | 'backward';

function intAttr(value: string | null): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getSpTreeShapes(shapeTree: Element): Element[] {
  return getElementChildren(shapeTree).filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName));
}

function getShapeBox(shape: Element): ShapeBox | null {
  const offset = getDescendants(shape, 'off')[0];
  const extent = getDescendants(shape, 'ext')[0];
  if (!offset || !extent) return null;
  return {
    x: intAttr(offset.getAttribute('x')),
    y: intAttr(offset.getAttribute('y')),
    cx: intAttr(extent.getAttribute('cx')),
    cy: intAttr(extent.getAttribute('cy'))
  };
}

function getShapeTreeElement(slideDoc: XMLDocument): Element {
  const shapeTree = getDescendants(slideDoc, 'spTree')[0];
  if (!shapeTree) {
    throw new Error('Could not find the slide shape tree.');
  }
  return shapeTree;
}

/** Resolve a pptx-svg renderer composite shape index to its OOXML element. */
function getShapeElementByRendererIndex(slideDoc: XMLDocument, shapeIndex: number): Element {
  const shapes = getSpTreeShapes(getShapeTreeElement(slideDoc));
  if (shapeIndex < 1000) {
    const shape = shapes[shapeIndex];
    if (!shape) {
      throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
    }
    return shape;
  }

  const groupIndex = Math.floor(shapeIndex / 1000);
  const childIndex = shapeIndex % 1000;
  const group = shapes[groupIndex];
  if (!group || group.localName !== 'grpSp') {
    throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
  }

  const children = getElementChildren(group).filter((element) =>
    SHAPE_ELEMENT_NAMES.has(element.localName)
  );
  const child = children[childIndex];
  if (!child) {
    throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
  }
  return child;
}

function applyTransformToShape(shape: Element, transform: ShapeTransform): boolean {
  const xfrm = getDescendants(shape, 'xfrm')[0];
  if (!xfrm) return false;

  let offset = getElementChildren(xfrm).find(
    (element) => element.localName === 'off' && element.namespaceURI === DRAWINGML_NAMESPACE
  );
  let extent = getElementChildren(xfrm).find(
    (element) => element.localName === 'ext' && element.namespaceURI === DRAWINGML_NAMESPACE
  );
  if (!offset) {
    offset = shape.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:off');
    xfrm.insertBefore(offset, xfrm.firstChild);
  }
  if (!extent) {
    extent = shape.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:ext');
    xfrm.appendChild(extent);
  }

  offset.setAttribute('x', String(Math.round(transform.x)));
  offset.setAttribute('y', String(Math.round(transform.y)));
  extent.setAttribute('cx', String(Math.max(1, Math.round(transform.cx))));
  extent.setAttribute('cy', String(Math.max(1, Math.round(transform.cy))));
  xfrm.setAttribute('rot', String(Math.round(transform.rot)));
  return true;
}

function nextShapeId(slideDoc: XMLDocument): number {
  let maxId = 1;
  for (const cNvPr of getDescendants(slideDoc, 'cNvPr')) {
    const id = Number(cNvPr.getAttribute('id'));
    if (Number.isFinite(id) && id > maxId) maxId = id;
  }
  return maxId + 1;
}

function qualifyName(reference: Element, localName: string): string {
  return reference.prefix ? `${reference.prefix}:${localName}` : localName;
}

function adjacentUnselectedShape(
  element: Element,
  selected: Set<Element>,
  direction: 1 | -1
): Element | null {
  let current = direction === 1 ? element.nextElementSibling : element.previousElementSibling;
  while (current) {
    if (SHAPE_ELEMENT_NAMES.has(current.localName) && !selected.has(current)) return current;
    current = direction === 1 ? current.nextElementSibling : current.previousElementSibling;
  }
  return null;
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

function clearDrawingParagraphContent(paragraph: Element): Element | null {
  let templateRun: Element | null = null;
  for (const child of [...getElementChildren(paragraph)]) {
    if (child.localName === 'pPr' && child.namespaceURI === DRAWINGML_NAMESPACE) continue;
    if (child.localName === 'r' && child.namespaceURI === DRAWINGML_NAMESPACE && !templateRun) {
      templateRun = child;
    }
    paragraph.removeChild(child);
  }
  return templateRun;
}

function appendDrawingParagraphRun(paragraph: Element, templateRun: Element | null, text: string): void {
  const doc = paragraph.ownerDocument;
  const run = templateRun ? cloneDrawingRun(templateRun, doc) : doc.createElementNS(DRAWINGML_NAMESPACE, 'a:r');
  setDrawingRunText(run, text);
  paragraph.appendChild(run);
}

function setDrawingParagraphText(container: Element, paragraphIndex: number, text: string): void {
  const paragraphs = getDrawingParagraphs(container);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new Error('Could not find the selected text paragraph.');
  }

  if (!text.includes('\n')) {
    const runs = getDrawingRuns(paragraph);
    if (runs.length === 0) {
      throw new Error('Could not find the selected text paragraph runs.');
    }

    runs.forEach((run, runIndex) => {
      setDrawingRunText(run, runIndex === 0 ? text : '');
    });
    return;
  }

  const doc = paragraph.ownerDocument;
  const templateRun = clearDrawingParagraphContent(paragraph);
  const lines = text.split('\n');
  lines.forEach((line, lineIndex) => {
    if (lineIndex > 0) {
      paragraph.appendChild(doc.createElementNS(DRAWINGML_NAMESPACE, 'a:br'));
    }
    appendDrawingParagraphRun(paragraph, templateRun, line);
  });
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

function getDrawingRunText(run: Element): string {
  const textElement = getElementChildren(run)
    .find((element) => element.localName === 't' && element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? getDescendants(run, 't').find((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  return textElement?.textContent ?? '';
}

function setDrawingRunText(run: Element, text: string): void {
  const textElement = getElementChildren(run)
    .find((element) => element.localName === 't' && element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? getDescendants(run, 't').find((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  if (!textElement) {
    throw new Error('Could not find the selected text node.');
  }
  textElement.textContent = text;
}

function cloneDrawingRun(run: Element, doc: XMLDocument): Element {
  return run.cloneNode(true) as Element;
}

interface DrawingRunSegment {
  run: Element;
  runIndex: number;
  start: number;
  end: number;
  text: string;
}

function getDrawingRunSegments(paragraph: Element): DrawingRunSegment[] {
  const segments: DrawingRunSegment[] = [];
  let offset = 0;
  getDrawingRuns(paragraph).forEach((run, runIndex) => {
    const text = getDrawingRunText(run);
    segments.push({ run, runIndex, start: offset, end: offset + text.length, text });
    offset += text.length;
  });
  return segments;
}

function splitDrawingRunAt(
  paragraph: Element,
  runIndex: number,
  localOffset: number,
  doc: XMLDocument
): void {
  const runs = getDrawingRuns(paragraph);
  const run = runs[runIndex];
  if (!run) return;

  const text = getDrawingRunText(run);
  if (localOffset <= 0 || localOffset >= text.length) return;

  setDrawingRunText(run, text.slice(0, localOffset));
  const afterRun = cloneDrawingRun(run, doc);
  setDrawingRunText(afterRun, text.slice(localOffset));

  const next = run.nextSibling;
  if (next) {
    paragraph.insertBefore(afterRun, next);
  } else {
    paragraph.appendChild(afterRun);
  }
}

function splitParagraphAtOffset(paragraph: Element, offset: number, doc: XMLDocument): void {
  if (offset <= 0) return;

  const segments = getDrawingRunSegments(paragraph);
  const total = segments.at(-1)?.end ?? 0;
  if (offset >= total) return;

  for (const segment of segments) {
    if (offset > segment.start && offset < segment.end) {
      splitDrawingRunAt(paragraph, segment.runIndex, offset - segment.start, doc);
      return;
    }
  }
}

function applyRunStyleToParagraphRange(
  paragraph: Element,
  doc: XMLDocument,
  startOffset: number,
  endOffset: number,
  change: RunStyleChange
): boolean {
  let start = Math.max(0, startOffset);
  let end = Math.max(0, endOffset);
  if (start > end) {
    [start, end] = [end, start];
  }

  if (start === end) {
    const segments = getDrawingRunSegments(paragraph);
    const total = segments.at(-1)?.end ?? 0;
    const position = Math.min(start, total);
    const segment = segments.find((candidate) => position >= candidate.start && position <= candidate.end)
      ?? segments.at(-1);
    if (!segment || segment.text.length === 0) return false;
    applyRunPropertyChange(getRunProperties(segment.run, doc), doc, change);
    return true;
  }

  splitParagraphAtOffset(paragraph, start, doc);
  splitParagraphAtOffset(paragraph, end, doc);

  let changed = false;
  for (const segment of getDrawingRunSegments(paragraph)) {
    if (segment.end <= start || segment.start >= end) continue;
    if (segment.text.length === 0) continue;
    applyRunPropertyChange(getRunProperties(segment.run, doc), doc, change);
    changed = true;
  }
  return changed;
}

function isParagraphRangeStyled(
  paragraph: Element,
  startOffset: number,
  endOffset: number,
  flag: 'bold' | 'italic' | 'underline'
): boolean {
  let start = Math.max(0, startOffset);
  let end = Math.max(0, endOffset);
  if (start > end) {
    [start, end] = [end, start];
  }
  if (start === end) return false;

  let matched = false;
  for (const segment of getDrawingRunSegments(paragraph)) {
    if (segment.end <= start || segment.start >= end) continue;
    if (segment.text.length === 0) continue;
    matched = true;

    const runProperties = getElementChildren(segment.run)
      .find((element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
    if (!runProperties) return false;

    if (flag === 'bold') {
      const bold = runProperties.getAttribute('b');
      if (bold !== '1' && bold !== 'true') return false;
    } else if (flag === 'italic') {
      const italic = runProperties.getAttribute('i');
      if (italic !== '1' && italic !== 'true') return false;
    } else {
      const underline = runProperties.getAttribute('u');
      if (!underline || underline === 'none') return false;
    }
  }
  return matched;
}

/**
 * Replace every occurrence of `query` with `replacement` inside a single text
 * string and report how many substitutions were made. Operates on plain run
 * text, so matches that span multiple runs are not handled here.
 */
function replaceTextOccurrences(
  source: string,
  query: string,
  replacement: string,
  matchCase: boolean
): { result: string; count: number } {
  if (!query) {
    return { result: source, count: 0 };
  }

  const haystack = matchCase ? source : source.toLocaleLowerCase();
  const needle = matchCase ? query : query.toLocaleLowerCase();
  let result = '';
  let count = 0;
  let index = 0;

  while (index <= source.length) {
    const found = haystack.indexOf(needle, index);
    if (found === -1) {
      result += source.slice(index);
      break;
    }
    result += source.slice(index, found) + replacement;
    index = found + needle.length;
    count += 1;
  }

  return { result, count };
}

interface DrawingRunPosition {
  paragraphIndex: number;
  runIndex: number;
  run: Element;
}

// Canonical child order of CT_TextCharacterProperties (a:rPr). New children must be
// inserted at the right position or PowerPoint rejects the slide on save.
const RUN_PROPERTY_CHILD_ORDER = [
  'ln',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'effectLst',
  'effectDag',
  'highlight',
  'uLnTx',
  'uLn',
  'uFillTx',
  'uFill',
  'latin',
  'ea',
  'cs',
  'sym',
  'hlinkClick',
  'hlinkMouseOver',
  'rtl',
  'extLst'
];

function getDrawingRuns(paragraph: Element): Element[] {
  return getElementChildren(paragraph)
    .filter((element) => element.localName === 'r' && element.namespaceURI === DRAWINGML_NAMESPACE);
}

function getShapeRunPositions(shape: Element): DrawingRunPosition[] {
  const positions: DrawingRunPosition[] = [];
  getDrawingParagraphs(shape).forEach((paragraph, paragraphIndex) => {
    getDrawingRuns(paragraph).forEach((run, runIndex) => {
      positions.push({ paragraphIndex, runIndex, run });
    });
  });
  return positions;
}

/**
 * Disable shrink-to-fit ("normAutofit") on every text body in the shape. The
 * SVG renderer recomputes normAutofit dynamically and ignores the stored
 * fontScale, so when a user explicitly sets a font size the only way to honor
 * it is to turn the shrinking autofit off. normAutofit is replaced with
 * noAutofit; spAutoFit (resize shape to text) already honors the size and is
 * left intact.
 */
function disableShrinkAutofit(shape: Element, doc: XMLDocument): boolean {
  const bodyProps = getDescendants(shape, 'bodyPr')
    .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);

  let changed = false;
  for (const bodyPr of bodyProps) {
    const normAutofit = getElementChildren(bodyPr)
      .find((element) => element.localName === 'normAutofit' && element.namespaceURI === DRAWINGML_NAMESPACE);
    if (!normAutofit) continue;

    const noAutofit = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:noAutofit');
    bodyPr.replaceChild(noAutofit, normAutofit);
    changed = true;
  }
  return changed;
}

function getRunProperties(run: Element, doc: XMLDocument): Element {
  const existing = getElementChildren(run)
    .find((element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
  if (existing) {
    return existing;
  }

  const rPr = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:rPr');
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}

function insertRunPropertyChild(rPr: Element, child: Element): void {
  const order = RUN_PROPERTY_CHILD_ORDER.indexOf(child.localName);
  const reference = getElementChildren(rPr).find((existing) => {
    const existingOrder = RUN_PROPERTY_CHILD_ORDER.indexOf(existing.localName);
    return existingOrder !== -1 && existingOrder > order;
  }) ?? null;
  rPr.insertBefore(child, reference);
}

function setRunHighlight(rPr: Element, doc: XMLDocument, highlight: string | null): void {
  getElementChildren(rPr)
    .filter((element) => element.localName === 'highlight' && element.namespaceURI === DRAWINGML_NAMESPACE)
    .forEach((element) => rPr.removeChild(element));

  if (highlight === null) {
    return;
  }

  const highlightElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:highlight');
  const colorElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:srgbClr');
  colorElement.setAttribute('val', normalizeHexColor(highlight));
  highlightElement.appendChild(colorElement);
  insertRunPropertyChild(rPr, highlightElement);
}

function normalizeHexColor(hex: string): string {
  return hex.replace(/^#/, '').toUpperCase();
}

function setRunLatinFont(rPr: Element, doc: XMLDocument, fontFamily: string): void {
  getElementChildren(rPr)
    .filter((element) => element.localName === 'latin' && element.namespaceURI === DRAWINGML_NAMESPACE)
    .forEach((element) => rPr.removeChild(element));

  const latin = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:latin');
  latin.setAttribute('typeface', fontFamily);
  insertRunPropertyChild(rPr, latin);
}

function setRunSolidFill(rPr: Element, doc: XMLDocument, color: string | null): void {
  getElementChildren(rPr)
    .filter((element) => element.localName === 'solidFill' && element.namespaceURI === DRAWINGML_NAMESPACE)
    .forEach((element) => rPr.removeChild(element));

  if (color === null) {
    return;
  }

  const solidFill = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:solidFill');
  const colorElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:srgbClr');
  colorElement.setAttribute('val', normalizeHexColor(color));
  solidFill.appendChild(colorElement);
  insertRunPropertyChild(rPr, solidFill);
}

// Apply every requested run-level property directly to an <a:rPr>. The WASM
// renderer does not preserve <a:highlight> when it re-serializes a slide, so
// every run-style edit (not just highlight) is performed via OOXML to keep the
// highlight from being clobbered by a later renderer mutation on the same run.
function applyRunPropertyChange(rPr: Element, doc: XMLDocument, change: RunStyleChange): void {
  if (change.bold !== undefined) {
    rPr.setAttribute('b', change.bold ? '1' : '0');
  }
  if (change.italic !== undefined) {
    rPr.setAttribute('i', change.italic ? '1' : '0');
  }
  if (change.underline !== undefined) {
    rPr.setAttribute('u', change.underline ? 'sng' : 'none');
  }
  if (change.fontSizePt !== undefined) {
    rPr.setAttribute('sz', String(Math.round(change.fontSizePt * 100)));
  }
  if (change.fontFamily !== undefined && change.fontFamily !== '') {
    setRunLatinFont(rPr, doc, change.fontFamily);
  }
  if (change.color !== undefined) {
    setRunSolidFill(rPr, doc, change.color);
  }
  if (change.highlight !== undefined) {
    setRunHighlight(rPr, doc, change.highlight);
  }
}

function getParagraphProperties(paragraph: Element, doc: XMLDocument): Element {
  const existing = getElementChildren(paragraph)
    .find((element) => element.localName === 'pPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
  if (existing) {
    return existing;
  }

  const pPr = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:pPr');
  paragraph.insertBefore(pPr, paragraph.firstChild);
  return pPr;
}

function resolvePptxRunAlignment(value: string | null): 'ctr' | 'just' | 'l' | 'r' | null {
  if (value === 'l' || value === 'ctr' || value === 'r' || value === 'just') {
    return value;
  }
  return null;
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

function getPartExtension(partPath: string): string {
  return partPath.match(/\.([^./]+)$/)?.[1]?.toLowerCase() ?? '';
}

function imageExtensionForMime(mimeType: string): string {
  switch (mimeType.toLowerCase()) {
    case 'image/jpeg':
      return 'jpeg';
    case 'image/gif':
      return 'gif';
    case 'image/webp':
      return 'webp';
    case 'image/bmp':
      return 'bmp';
    case 'image/png':
    default:
      return 'png';
  }
}

function contentTypeForImageExtension(extension: string): string {
  switch (extension.toLowerCase()) {
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'webp':
      return 'image/webp';
    case 'bmp':
      return 'image/bmp';
    case 'png':
    default:
      return 'image/png';
  }
}

function createRelationshipsDocument(): XMLDocument {
  return parseXml(
    `<Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}"/>`,
    '(new relationships part)'
  );
}

function nextRelationshipId(relationships: XMLDocument): string {
  const used = new Set(
    getDescendants(relationships, 'Relationship')
      .map((relationship) => relationship.getAttribute('Id'))
      .filter((id): id is string => Boolean(id))
  );
  let next = 1;
  while (used.has(`rId${next}`)) next++;
  return `rId${next}`;
}

function nextImageMediaPath(
  textFiles: Map<string, string>,
  binaryFiles: Map<string, Uint8Array>,
  extension: string
): string {
  let maxIndex = 0;
  const pattern = /^ppt\/media\/image(\d+)\.[^./]+$/;
  for (const key of [...textFiles.keys(), ...binaryFiles.keys()]) {
    const match = key.match(pattern);
    if (match) maxIndex = Math.max(maxIndex, Number(match[1]));
  }
  return `ppt/media/image${maxIndex + 1}.${extension}`;
}

function ensureDefaultContentType(
  contentTypesDoc: XMLDocument,
  extension: string,
  contentType: string
): void {
  const normalized = extension.toLowerCase();
  const exists = getDescendants(contentTypesDoc, 'Default')
    .some((entry) => entry.getAttribute('Extension')?.toLowerCase() === normalized);
  if (exists) return;

  const namespace = contentTypesDoc.documentElement.namespaceURI;
  const entry = contentTypesDoc.createElementNS(namespace, 'Default');
  entry.setAttribute('Extension', normalized);
  entry.setAttribute('ContentType', contentType);
  contentTypesDoc.documentElement.appendChild(entry);
}

function getBlipEmbedId(blip: Element): string | null {
  return blip.getAttributeNS(RELATIONSHIP_NAMESPACE, 'embed') || blip.getAttribute('r:embed');
}

function setBlipEmbedId(blip: Element, relationshipId: string): void {
  const existing = blip.getAttributeNodeNS(RELATIONSHIP_NAMESPACE, 'embed');
  if (existing) {
    existing.value = relationshipId;
  } else {
    blip.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:embed', relationshipId);
  }
}

// Convert an inset crop percentage (0-100) to the OOXML <a:srcRect> unit of
// 1000ths of a percent (0-100000), clamped to the valid range.
function cropPercentToPermille(percent: number): number {
  if (!Number.isFinite(percent)) return 0;
  return Math.max(0, Math.min(100000, Math.round(percent * 1000)));
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

  let nextRelId = 1;
  for (const match of relationships.matchAll(/\bId="rId(\d+)"/g)) {
    nextRelId = Math.max(nextRelId, Number(match[1]) + 1);
  }

  const slideIds = Array.from(presentation.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"[^>]*\/?>/g))
    .map((match) => Number(match[1]))
    .filter(Number.isFinite);
  let nextSlideId = Math.max(255, ...slideIds) + 1;
  const normalizedSlideEntries: string[] = [];
  const normalizedRelationships: string[] = [];

  for (let index = 0; index < slideCount; index++) {
    const relationshipId = `rId${nextRelId++}`;
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

  let changed = false;

  const previousExtensionList = getDirectChild(previousCommonSlide, 'extLst');
  if (previousExtensionList) {
    const exportedExtensionList = getDirectChild(exportedCommonSlide, 'extLst');
    const importedExtensionList = exportedDocument.importNode(previousExtensionList, true);

    if (exportedExtensionList) {
      exportedCommonSlide.replaceChild(importedExtensionList, exportedExtensionList);
    } else {
      exportedCommonSlide.appendChild(importedExtensionList);
    }
    changed = true;
  }

  if (restoreShapeNonVisualIdentity(previousDocument, exportedDocument)) {
    changed = true;
  }

  return changed ? serializeXml(exportedDocument) : null;
}

interface ShapeIdentity {
  cNvPr: Element;
  id: string;
  name: string;
  extensionList: Element | null;
  fingerprint: string;
  shapeKind: string;
}

// When the renderer re-serializes a slide whose shapes were mutated, it strips each
// shape's non-visual identity: it resets <p:cNvPr> ids to "0", clears names, and drops
// the per-shape extension list (e.g. the <a16:creationId> that Office writes on every
// shape). Restore that identity from the previous slide so edits do not silently lose
// it (which otherwise fails save validation on virtually every real-world deck).
//
// Each renderer mutation touches a single shape, so when the shape count is unchanged
// the previous and exported shape trees line up by index. When a shape was added or
// removed the counts differ, so unchanged shapes are matched by a geometry + text
// fingerprint instead.
function restoreShapeNonVisualIdentity(previousDocument: XMLDocument, exportedDocument: XMLDocument): boolean {
  const previousShapes = collectShapeIdentities(previousDocument);
  const exportedShapes = collectShapeIdentities(exportedDocument);
  if (previousShapes.length === 0 || exportedShapes.length === 0) return false;

  const pairs: Array<[ShapeIdentity, ShapeIdentity]> = [];
  if (previousShapes.length === exportedShapes.length) {
    exportedShapes.forEach((exported, index) => {
      const previous = previousShapes[index];
      if (previous && previous.shapeKind === exported.shapeKind) {
        pairs.push([previous, exported]);
      }
    });
  } else {
    const remaining = [...previousShapes];
    for (const exported of exportedShapes) {
      const matchIndex = remaining.findIndex(
        (candidate) =>
          candidate.shapeKind === exported.shapeKind && candidate.fingerprint === exported.fingerprint
      );
      const previous = matchIndex >= 0 ? remaining[matchIndex] : undefined;
      if (previous) {
        pairs.push([previous, exported]);
        remaining.splice(matchIndex, 1);
      }
    }
  }

  let changed = false;
  for (const [previous, exported] of pairs) {
    if (isAnonymizedShapeId(exported.id) && !isAnonymizedShapeId(previous.id)) {
      exported.cNvPr.setAttribute('id', previous.id);
      if (!exported.name && previous.name) {
        exported.cNvPr.setAttribute('name', previous.name);
      }
      changed = true;
    }
    if (previous.extensionList && !exported.extensionList) {
      exported.cNvPr.appendChild(exportedDocument.importNode(previous.extensionList, true));
      changed = true;
    }
  }

  if (ensureUniqueShapeIds(exportedShapes)) {
    changed = true;
  }

  return changed;
}

function collectShapeIdentities(document: XMLDocument): ShapeIdentity[] {
  return getDescendants(document, 'cNvPr').map((cNvPr) => {
    const shape = cNvPr.parentNode?.parentNode as Element | undefined;
    return {
      cNvPr,
      id: cNvPr.getAttribute('id') ?? '',
      name: cNvPr.getAttribute('name') ?? '',
      extensionList: getDirectChild(cNvPr, 'extLst'),
      fingerprint: getShapeFingerprint(shape),
      shapeKind: shape?.localName ?? ''
    };
  });
}

function getShapeFingerprint(shape: Element | undefined): string {
  if (!shape) return '';
  const transform = getDescendants(shape, 'xfrm')[0];
  const offset = transform ? getDirectChild(transform, 'off') : null;
  const extent = transform ? getDirectChild(transform, 'ext') : null;
  const geometry = [
    offset?.getAttribute('x') ?? '',
    offset?.getAttribute('y') ?? '',
    extent?.getAttribute('cx') ?? '',
    extent?.getAttribute('cy') ?? ''
  ].join(',');
  const text = (shape.textContent ?? '').replace(/\s+/g, ' ').trim().slice(0, 64);
  return `${shape.localName}|${geometry}|${text}`;
}

function isAnonymizedShapeId(id: string): boolean {
  return id === '' || id === '0';
}

function ensureUniqueShapeIds(shapes: ShapeIdentity[]): boolean {
  const reservedIds = new Set<number>();
  for (const shape of shapes) {
    const numericId = Number(shape.cNvPr.getAttribute('id'));
    if (Number.isInteger(numericId) && numericId > 0) {
      reservedIds.add(numericId);
    }
  }

  const consumedIds = new Set<number>();
  let changed = false;
  let nextId = 1;
  for (const shape of shapes) {
    const numericId = Number(shape.cNvPr.getAttribute('id'));
    if (Number.isInteger(numericId) && numericId > 0 && !consumedIds.has(numericId)) {
      consumedIds.add(numericId);
      continue;
    }
    while (consumedIds.has(nextId) || reservedIds.has(nextId)) nextId++;
    shape.cNvPr.setAttribute('id', String(nextId));
    consumedIds.add(nextId);
    changed = true;
  }
  return changed;
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

    await initRendererBackend(renderer);
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

  async updateShapeTransform(
    slideIndex: number,
    shapeIndex: number,
    transform: ShapeTransform
  ): Promise<void> {
    const result = this.renderer.updateShapeTransform(
      slideIndex,
      shapeIndex,
      Math.round(transform.x),
      Math.round(transform.y),
      Math.max(1, Math.round(transform.cx)),
      Math.max(1, Math.round(transform.cy)),
      Math.round(transform.rot)
    );
    if (!result.startsWith('ERROR:')) {
      return;
    }

    const message = result.slice('ERROR:'.length).trim().toLowerCase();
    if (!message.includes('out of range')) {
      assertOk(result, 'Could not update shape transform.');
    }

    await this.updateShapeTransformInOoxml(slideIndex, shapeIndex, transform);
  }

  private async updateShapeTransformInOoxml(
    slideIndex: number,
    shapeIndex: number,
    transform: ShapeTransform
  ): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shape = getShapeElementByRendererIndex(slideDoc, shapeIndex);
    if (!applyTransformToShape(shape, transform)) {
      throw new Error('Could not update shape transform.');
    }

    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
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

  async updateParagraphText(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
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
    setDrawingParagraphText(shape, paragraphIndex, text);
    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
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

  /**
   * Replace text across slides, returning how many occurrences were changed.
   * Replacement happens within individual DrawingML text runs (a:t), so
   * formatting on each run is preserved. Matches that span multiple runs are
   * not substituted. Pass `slideIndex`/`shapeIndex` to limit the replacement to
   * a single shape (used for "Replace" on the current match); omit them to
   * replace everywhere ("Replace all").
   */
  async replaceText(
    query: string,
    replacement: string,
    options: { matchCase?: boolean; slideIndex?: number; shapeIndex?: number } = {}
  ): Promise<number> {
    if (!query) {
      return 0;
    }

    const matchCase = options.matchCase ?? false;
    const scoped = options.slideIndex !== undefined && options.shapeIndex !== undefined;
    const slideStart = scoped ? (options.slideIndex as number) : 0;
    const slideEnd = scoped ? (options.slideIndex as number) + 1 : this.slideCountValue;

    const rawExport = await this.exportRendererState();
    const zip = await extractZip(rawExport);
    const updatedFiles = new Map<string, string>();
    let total = 0;

    for (let slideIndex = slideStart; slideIndex < slideEnd; slideIndex++) {
      const slidePath = getSlidePath(slideIndex);
      const slideXml = zip.textFiles.get(slidePath);
      if (!slideXml) continue;

      const slideDoc = parseXml(slideXml, slidePath);
      let scope: Element | XMLDocument = slideDoc;
      if (scoped) {
        try {
          scope = getShapeElement(slideDoc, options.shapeIndex as number);
        } catch {
          continue;
        }
      }

      const textElements = getDescendants(scope, 't')
        .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
      let slideChanged = false;
      for (const textElement of textElements) {
        const original = textElement.textContent ?? '';
        if (!original) continue;
        const { result, count } = replaceTextOccurrences(original, query, replacement, matchCase);
        if (count > 0) {
          textElement.textContent = result;
          total += count;
          slideChanged = true;
        }
      }

      if (slideChanged) {
        updatedFiles.set(slidePath, serializeXml(slideDoc));
      }
    }

    if (total > 0) {
      const patchedExport = await buildZip(rawExport, updatedFiles);
      await this.reloadFromBuffer(patchedExport, this.slideCountValue);
    }

    return total;
  }

  /**
   * Read the resolved style of a single text run for reflecting toolbar state.
   * Only directly-authored run/paragraph properties are reported; values
   * inherited from a placeholder, layout, or master are not resolved here.
   */
  getRunStyle(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    runIndex: number
  ): RunStyleInfo | null {
    let shape: Element;
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      shape = getShapeElement(slideDoc, shapeIndex);
    } catch {
      return null;
    }

    const paragraph = getDrawingParagraphs(shape)[paragraphIndex];
    if (!paragraph) {
      return null;
    }

    const run = getDrawingRuns(paragraph)[runIndex];
    if (!run) {
      return null;
    }

    const runProperties = getElementChildren(run)
      .find((element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE) ?? null;
    const paragraphProperties = getElementChildren(paragraph)
      .find((element) => element.localName === 'pPr' && element.namespaceURI === DRAWINGML_NAMESPACE) ?? null;

    const bold = runProperties?.getAttribute('b');
    const italic = runProperties?.getAttribute('i');
    const underline = runProperties?.getAttribute('u');
    const fontSize = runProperties?.getAttribute('sz');
    const latin = runProperties
      ? getElementChildren(runProperties).find((element) => element.localName === 'latin')
      : undefined;
    const solidFill = runProperties
      ? getElementChildren(runProperties).find((element) => element.localName === 'solidFill')
      : undefined;
    const fillColor = solidFill
      ? getElementChildren(solidFill).find((element) => element.localName === 'srgbClr')
      : undefined;
    const highlight = runProperties
      ? getElementChildren(runProperties).find((element) => element.localName === 'highlight')
      : undefined;
    const highlightColor = highlight
      ? getElementChildren(highlight).find((element) => element.localName === 'srgbClr')
      : undefined;
    const parsedFontSize = fontSize ? Number(fontSize) : Number.NaN;

    return {
      bold: bold === '1' || bold === 'true',
      italic: italic === '1' || italic === 'true',
      underline: Boolean(underline) && underline !== 'none',
      fontFamily: latin?.getAttribute('typeface') ?? null,
      fontSizePt: Number.isFinite(parsedFontSize) ? parsedFontSize / 100 : null,
      color: this.readColorValue(fillColor),
      highlight: this.readColorValue(highlightColor),
      alignment: resolvePptxRunAlignment(paragraphProperties?.getAttribute('algn') ?? null)
    };
  }

  private readColorValue(colorElement: Element | undefined): string | null {
    const value = colorElement?.getAttribute('val');
    return value ? normalizeHexColor(value) : null;
  }

  /**
   * Apply run-level formatting to a single run, or — when `target` is null — to
   * every run in the shape. All properties (bold/italic/underline/size/color/
   * font/highlight) are applied via direct OOXML editing rather than the WASM
   * renderer's run-style setters: the renderer drops <a:highlight> whenever it
   * re-serializes a slide, so routing every edit through OOXML keeps the
   * highlight intact across subsequent formatting actions.
   */
  async setRunStyle(
    slideIndex: number,
    shapeIndex: number,
    target: RunTarget | null,
    change: RunStyleChange
  ): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const positions = getShapeRunPositions(shape);
      const targets = target
        ? positions.filter(
            (position) => position.paragraphIndex === target.paragraphIndex && position.runIndex === target.runIndex
          )
        : positions;

      let changed = false;
      for (const { run } of targets) {
        applyRunPropertyChange(getRunProperties(run, slideDoc), slideDoc, change);
        changed = true;
      }
      if (change.fontSizePt !== undefined && changed) {
        disableShrinkAutofit(shape, slideDoc);
      }
      return changed;
    });
  }

  /**
   * Apply run-level formatting to the character range `[startOffset, endOffset)`
   * inside a single paragraph. When the range is collapsed, the run containing
   * the caret is styled. Runs are split at the range boundaries as needed.
   */
  async setRunStyleForRange(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    startOffset: number,
    endOffset: number,
    change: RunStyleChange
  ): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const paragraph = getDrawingParagraphs(shape)[paragraphIndex];
      if (!paragraph) {
        throw new Error('Could not find the selected text paragraph.');
      }

      const changed = applyRunStyleToParagraphRange(paragraph, slideDoc, startOffset, endOffset, change);
      if (change.fontSizePt !== undefined && changed) {
        disableShrinkAutofit(shape, slideDoc);
      }
      return changed;
    });
  }

  /**
   * Apply run-level formatting across one or more paragraph ranges in the same
   * shape. This keeps a multi-paragraph toolbar action in one OOXML mutation.
   */
  async setRunStyleForRanges(
    slideIndex: number,
    shapeIndex: number,
    ranges: ParagraphTextRange[],
    change: RunStyleChange
  ): Promise<void> {
    const normalizedRanges = ranges.filter((range) => (
      Number.isFinite(range.paragraphIndex)
      && Number.isFinite(range.start)
      && Number.isFinite(range.end)
    ));
    if (normalizedRanges.length === 0) return;

    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const paragraphs = getDrawingParagraphs(shape);
      let changed = false;
      for (const range of normalizedRanges) {
        const paragraph = paragraphs[range.paragraphIndex];
        if (!paragraph) {
          throw new Error('Could not find the selected text paragraph.');
        }
        changed = applyRunStyleToParagraphRange(paragraph, slideDoc, range.start, range.end, change) || changed;
      }
      if (change.fontSizePt !== undefined && changed) {
        disableShrinkAutofit(shape, slideDoc);
      }
      return changed;
    });
  }

  /** Whether every non-empty run in `[startOffset, endOffset)` has `flag` set. */
  isRangeStyled(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    startOffset: number,
    endOffset: number,
    flag: 'bold' | 'italic' | 'underline'
  ): boolean {
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElement(slideDoc, shapeIndex);
      const paragraph = getDrawingParagraphs(shape)[paragraphIndex];
      if (!paragraph) return false;
      return isParagraphRangeStyled(paragraph, startOffset, endOffset, flag);
    } catch {
      return false;
    }
  }

  /** Whether every non-empty run in every selected range has `flag` set. */
  areRangesStyled(
    slideIndex: number,
    shapeIndex: number,
    ranges: ParagraphTextRange[],
    flag: 'bold' | 'italic' | 'underline'
  ): boolean {
    const normalizedRanges = ranges.filter((range) => (
      Number.isFinite(range.paragraphIndex)
      && Number.isFinite(range.start)
      && Number.isFinite(range.end)
      && range.start !== range.end
    ));
    if (normalizedRanges.length === 0) return false;

    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElement(slideDoc, shapeIndex);
      const paragraphs = getDrawingParagraphs(shape);
      return normalizedRanges.every((range) => {
        const paragraph = paragraphs[range.paragraphIndex];
        if (!paragraph) return false;
        return isParagraphRangeStyled(paragraph, range.start, range.end, flag);
      });
    } catch {
      return false;
    }
  }

  /**
   * Set paragraph alignment on a single paragraph, or — when `paragraphIndex`
   * is null — on every paragraph in the shape. Applied via OOXML for the same
   * highlight-preservation reason as {@link setRunStyle}.
   */
  async setParagraphAlignment(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number | null,
    align: ParagraphAlignment
  ): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const paragraphs = getDrawingParagraphs(shape);
      const targets = paragraphIndex !== null
        ? [paragraphs[paragraphIndex]]
        : paragraphs;

      let changed = false;
      for (const paragraph of targets) {
        if (!paragraph) continue;
        getParagraphProperties(paragraph, slideDoc).setAttribute('algn', align);
        changed = true;
      }
      return changed;
    });
  }

  /** Set paragraph alignment on the paragraphs touched by the selected ranges. */
  async setParagraphAlignmentForRanges(
    slideIndex: number,
    shapeIndex: number,
    ranges: ParagraphTextRange[],
    align: ParagraphAlignment
  ): Promise<void> {
    const paragraphIndices = new Set(
      ranges
        .map((range) => range.paragraphIndex)
        .filter((paragraphIndex) => Number.isFinite(paragraphIndex))
    );
    if (paragraphIndices.size === 0) return;

    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const paragraphs = getDrawingParagraphs(shape);
      let changed = false;
      for (const paragraphIndex of paragraphIndices) {
        const paragraph = paragraphs[paragraphIndex];
        if (!paragraph) {
          throw new Error('Could not find the selected text paragraph.');
        }
        getParagraphProperties(paragraph, slideDoc).setAttribute('algn', align);
        changed = true;
      }
      return changed;
    });
  }

  private async editSlideShape(
    slideIndex: number,
    shapeIndex: number,
    mutate: (shape: Element, slideDoc: XMLDocument) => boolean
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
    if (!mutate(shape, slideDoc)) {
      return;
    }

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

  private parseInsertedShapeIndex(result: string, fallback: string): number {
    assertOk(result, fallback);
    const shapeIndex = Number(result.split(':')[1]);
    if (!Number.isFinite(shapeIndex)) {
      throw new Error('The renderer did not return a valid shape index.');
    }
    return shapeIndex;
  }

  addImage(
    slideIndex: number,
    imageData: Uint8Array,
    mimeType: string,
    widthPx = 320,
    heightPx = 240
  ): number {
    const x = pxToEmu(140);
    const y = pxToEmu(120);
    const cx = pxToEmu(widthPx);
    const cy = pxToEmu(heightPx);
    const result = this.renderer.addImage(slideIndex, imageData, mimeType, x, y, cx, cy);
    return this.parseInsertedShapeIndex(result, 'Could not insert image.');
  }

  addShapeGeometry(slideIndex: number, geometry: InsertableShapeGeometry): number {
    const x = pxToEmu(160);
    const y = pxToEmu(140);
    const cx = pxToEmu(geometry === 'line' ? 220 : 240);
    const cy = pxToEmu(geometry === 'line' ? 0 : 160);
    const result = this.renderer.addShape(slideIndex, geometry, x, y, cx, cy, 66, 133, 244);
    return this.parseInsertedShapeIndex(result, 'Could not insert shape.');
  }

  addTextBox(slideIndex: number): number {
    const x = pxToEmu(180);
    const y = pxToEmu(120);
    const cx = pxToEmu(300);
    const cy = pxToEmu(80);
    const result = this.renderer.addShape(slideIndex, 'rect', x, y, cx, cy, -1, -1, -1);
    const shapeIndex = this.parseInsertedShapeIndex(result, 'Could not add text box.');

    const textResult = this.renderer.addShapeText(slideIndex, shapeIndex, 'New text', 1800, -1, -1, -1);
    assertOk(textResult, 'Could not add text to the new text box.');
    return shapeIndex;
  }

  async addTable(slideIndex: number, rows: number, cols: number): Promise<number> {
    const historyBuffer = await this.exportRendererState();
    const inserted = await insertTableIntoPresentation(historyBuffer, slideIndex, rows, cols);
    await this.reloadFromBuffer(inserted.buffer, this.slideCountValue);
    return inserted.shapeIndex;
  }

  async addChart(slideIndex: number): Promise<number> {
    const historyBuffer = await this.exportRendererState();
    const inserted = await insertChartIntoPresentation(historyBuffer, slideIndex);
    await this.reloadFromBuffer(inserted.buffer, this.slideCountValue);
    await this.refreshChartTextValues(inserted.buffer);
    return inserted.shapeIndex;
  }

  async applyListStyle(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    style: ParagraphListStyle
  ): Promise<void> {
    const rawExport = await this.renderer.exportPptx();
    const mergedSlide = await mergeSlideGraphicFramesFromBuffer(this.currentBuffer, rawExport, slideIndex);
    const mergedPackage = await mergeMissingPackageParts(this.currentBuffer, mergedSlide);
    const patched = await applyParagraphListStyle(mergedPackage, slideIndex, shapeIndex, paragraphIndex, style);
    const preserved = await preserveSlideExtensionLists(this.currentBuffer, patched);
    await this.reloadFromBuffer(preserved, this.slideCountValue);
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

  /**
   * Apply a structural slide-XML mutation to a slide's shape tree, then reload
   * the renderer from the patched buffer. The mutation runs against the live
   * DOM and its return value is forwarded to the caller. Reordering, grouping,
   * and ungrouping all edit OOXML directly (the renderer has no equivalent API)
   * so the existing shape identities are preserved across the round-trip.
   */
  private async mutateSlideTree<T>(
    slideIndex: number,
    mutate: (slideDoc: XMLDocument, shapeTree: Element) => T
  ): Promise<T> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shapeTree = getDescendants(slideDoc, 'spTree')[0];
    if (!shapeTree) {
      throw new Error('Could not find the slide shape tree.');
    }

    const result = mutate(slideDoc, shapeTree);
    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
    return result;
  }

  /**
   * Change the stacking order of one or more top-level shapes on a slide.
   * Selected shapes keep their relative order. Returns their new shape indices.
   */
  async reorderShapes(
    slideIndex: number,
    shapeIndexes: number[],
    mode: ShapeReorderMode
  ): Promise<number[]> {
    return this.mutateSlideTree(slideIndex, (_slideDoc, shapeTree) => {
      const shapes = getSpTreeShapes(shapeTree);
      const selected = new Set(
        shapeIndexes
          .map((index) => shapes[index])
          .filter((element): element is Element => Boolean(element))
      );
      if (selected.size === 0) {
        throw new Error('Select an object to reorder.');
      }

      const ordered = shapes.filter((element) => selected.has(element));
      if (mode === 'front') {
        for (const element of ordered) shapeTree.appendChild(element);
      } else if (mode === 'back') {
        const anchor = shapes.find((element) => !selected.has(element)) ?? null;
        if (anchor) {
          for (const element of ordered) shapeTree.insertBefore(element, anchor);
        }
      } else if (mode === 'forward') {
        for (let index = ordered.length - 1; index >= 0; index--) {
          const element = ordered[index];
          if (!element) continue;
          const next = adjacentUnselectedShape(element, selected, 1);
          if (next) shapeTree.insertBefore(element, next.nextSibling);
        }
      } else {
        for (const element of ordered) {
          const previous = adjacentUnselectedShape(element, selected, -1);
          if (previous) shapeTree.insertBefore(element, previous);
        }
      }

      const finalShapes = getSpTreeShapes(shapeTree);
      return ordered.map((element) => finalShapes.indexOf(element));
    });
  }

  /**
   * Wrap the selected top-level shapes into a new group. The group's bounding
   * box is the union of the children, and chOff/chExt mirror off/ext so each
   * child keeps its slide coordinates. Returns the new group's shape index.
   */
  async groupShapes(slideIndex: number, shapeIndexes: number[]): Promise<number> {
    return this.mutateSlideTree(slideIndex, (slideDoc, shapeTree) => {
      const shapes = getSpTreeShapes(shapeTree);
      const selected = new Set(
        shapeIndexes
          .map((index) => shapes[index])
          .filter((element): element is Element => Boolean(element))
      );
      if (selected.size < 2) {
        throw new Error('Select at least two objects to group.');
      }

      const ordered = shapes.filter((element) => selected.has(element));
      const anchor = ordered[0];
      if (!anchor) {
        throw new Error('Could not resolve the objects to group.');
      }

      let minX = Infinity;
      let minY = Infinity;
      let maxX = -Infinity;
      let maxY = -Infinity;
      for (const element of ordered) {
        const box = getShapeBox(element);
        if (!box) continue;
        minX = Math.min(minX, box.x);
        minY = Math.min(minY, box.y);
        maxX = Math.max(maxX, box.x + box.cx);
        maxY = Math.max(maxY, box.y + box.cy);
      }
      if (!Number.isFinite(minX)) {
        minX = 0;
        minY = 0;
        maxX = 0;
        maxY = 0;
      }

      const offsetX = minX;
      const offsetY = minY;
      const extentCx = Math.max(1, maxX - minX);
      const extentCy = Math.max(1, maxY - minY);
      const presentationNs = shapeTree.namespaceURI;
      const newId = nextShapeId(slideDoc);

      const groupShape = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'grpSp'));
      const nonVisual = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'nvGrpSpPr'));
      const cNvPr = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'cNvPr'));
      cNvPr.setAttribute('id', String(newId));
      cNvPr.setAttribute('name', `Group ${newId}`);
      const cNvGrpSpPr = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'cNvGrpSpPr'));
      const nvPr = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'nvPr'));
      nonVisual.appendChild(cNvPr);
      nonVisual.appendChild(cNvGrpSpPr);
      nonVisual.appendChild(nvPr);

      const groupProps = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'grpSpPr'));
      const xfrm = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:xfrm');
      const off = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:off');
      off.setAttribute('x', String(Math.round(offsetX)));
      off.setAttribute('y', String(Math.round(offsetY)));
      const ext = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:ext');
      ext.setAttribute('cx', String(Math.round(extentCx)));
      ext.setAttribute('cy', String(Math.round(extentCy)));
      const chOff = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:chOff');
      chOff.setAttribute('x', String(Math.round(offsetX)));
      chOff.setAttribute('y', String(Math.round(offsetY)));
      const chExt = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:chExt');
      chExt.setAttribute('cx', String(Math.round(extentCx)));
      chExt.setAttribute('cy', String(Math.round(extentCy)));
      xfrm.appendChild(off);
      xfrm.appendChild(ext);
      xfrm.appendChild(chOff);
      xfrm.appendChild(chExt);
      groupProps.appendChild(xfrm);

      groupShape.appendChild(nonVisual);
      groupShape.appendChild(groupProps);
      shapeTree.insertBefore(groupShape, anchor);
      for (const element of ordered) groupShape.appendChild(element);

      return getSpTreeShapes(shapeTree).indexOf(groupShape);
    });
  }

  /**
   * Unwrap a group: move its child shapes back into the slide shape tree at the
   * group's position and remove the now-empty group. Child coordinates stay
   * valid because grouped shapes are stored with chOff/chExt equal to off/ext.
   * Returns the resulting shape indices of the freed children.
   */
  async ungroupShapes(slideIndex: number, shapeIndex: number): Promise<number[]> {
    return this.mutateSlideTree(slideIndex, (_slideDoc, shapeTree) => {
      const shapes = getSpTreeShapes(shapeTree);
      const group = shapes[shapeIndex];
      if (!group || group.localName !== 'grpSp') {
        throw new Error('Select a group to ungroup.');
      }

      const children = getElementChildren(group).filter((element) =>
        SHAPE_ELEMENT_NAMES.has(element.localName)
      );
      for (const child of children) shapeTree.insertBefore(child, group);
      shapeTree.removeChild(group);

      const finalShapes = getSpTreeShapes(shapeTree);
      return children.map((element) => finalShapes.indexOf(element));
    });
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

  async duplicateSlide(slideIndex: number): Promise<SlideMoveResult> {
    const { slideCount, insertedIdx } = await this.renderer.addSlide(slideIndex, slideIndex);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: insertedIdx, slideCount };
  }

  async reorderSlides(newOrder: number[]): Promise<SlideMoveResult> {
    const { slideCount } = await this.renderer.reorderSlides(newOrder);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: 0, slideCount };
  }

  async addSlideWithLayout(afterIndex: number, layout: SlideLayoutKind): Promise<SlideMoveResult> {
    const result = await this.addSlide(afterIndex);
    if (layout === 'blank') {
      return result;
    }

    const slideIndex = result.slideIndex;
    const { cx, cy } = await this.getSlideSizeEmu();
    const margin = pxToEmu(48);
    const contentWidth = Math.max(pxToEmu(120), cx - margin * 2);

    if (layout === 'title') {
      const titleHeight = pxToEmu(140);
      this.addLayoutPlaceholder(
        slideIndex,
        'Click to add title',
        margin,
        Math.max(margin, Math.round((cy - titleHeight) / 2)),
        contentWidth,
        titleHeight,
        4000,
        true
      );
    } else {
      const titleHeight = pxToEmu(120);
      this.addLayoutPlaceholder(slideIndex, 'Click to add title', margin, margin, contentWidth, titleHeight, 3600, true);
      const bodyTop = margin + titleHeight + pxToEmu(24);
      const bodyHeight = Math.max(pxToEmu(120), cy - bodyTop - margin);
      this.addLayoutPlaceholder(slideIndex, 'Click to add text', margin, bodyTop, contentWidth, bodyHeight, 1800, false);
    }

    return result;
  }

  private addLayoutPlaceholder(
    slideIndex: number,
    text: string,
    x: number,
    y: number,
    cx: number,
    cy: number,
    fontSize: number,
    center: boolean
  ): void {
    const shapeResult = this.renderer.addShape(
      slideIndex,
      'rect',
      Math.round(x),
      Math.round(y),
      Math.max(1, Math.round(cx)),
      Math.max(1, Math.round(cy)),
      -1,
      -1,
      -1
    );
    const shapeIndex = this.parseInsertedShapeIndex(shapeResult, 'Could not add layout placeholder.');
    assertOk(
      this.renderer.addShapeText(slideIndex, shapeIndex, text, fontSize, -1, -1, -1),
      'Could not add layout placeholder text.'
    );
    if (center) {
      this.renderer.updateParagraphAlign(slideIndex, shapeIndex, 0, 'ctr');
    }
  }

  async getSlideSizeEmu(): Promise<{ cx: number; cy: number }> {
    const fallback = { cx: 9144000, cy: 6858000 };
    try {
      const presentationPath = 'ppt/presentation.xml';
      const zip = await extractZip(this.currentBuffer);
      const presentationXml = zip.textFiles.get(presentationPath);
      if (!presentationXml) {
        return fallback;
      }

      const slideSize = getDescendants(parseXml(presentationXml, presentationPath), 'sldSz')[0];
      const cx = Number(slideSize?.getAttribute('cx'));
      const cy = Number(slideSize?.getAttribute('cy'));
      return {
        cx: Number.isFinite(cx) && cx > 0 ? cx : fallback.cx,
        cy: Number.isFinite(cy) && cy > 0 ? cy : fallback.cy
      };
    } catch {
      return fallback;
    }
  }

  async setSlideBackgroundColor(slideIndex: number, hex: string): Promise<void> {
    const normalizedHex = hex.replace(/^#/, '').trim().toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(normalizedHex)) {
      throw new Error('Background color must be a 6-digit RRGGBB hex value.');
    }

    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const commonSlide = getDescendants(slideDoc, 'cSld')[0];
    if (!commonSlide) {
      throw new Error('Slide is missing its common slide data.');
    }

    const presentationNamespace = slideDoc.documentElement.namespaceURI;
    for (const existingBackground of getElementChildren(commonSlide).filter((element) => element.localName === 'bg')) {
      commonSlide.removeChild(existingBackground);
    }

    const background = slideDoc.createElementNS(presentationNamespace, 'p:bg');
    const backgroundProperties = slideDoc.createElementNS(presentationNamespace, 'p:bgPr');
    const solidFill = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:solidFill');
    const color = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:srgbClr');
    color.setAttribute('val', normalizedHex);
    solidFill.appendChild(color);
    backgroundProperties.appendChild(solidFill);
    backgroundProperties.appendChild(slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:effectLst'));
    background.appendChild(backgroundProperties);
    commonSlide.insertBefore(background, commonSlide.firstChild);

    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  getSlideBackgroundColor(slideIndex: number): string | null {
    try {
      const slideXml = this.renderer.getSlideOoxml(slideIndex);
      if (!slideXml || slideXml.startsWith('ERROR:')) {
        return null;
      }

      const commonSlide = getDescendants(parseXml(slideXml, getSlidePath(slideIndex)), 'cSld')[0];
      const background = commonSlide
        ? getElementChildren(commonSlide).find((element) => element.localName === 'bg')
        : undefined;
      if (!background) {
        return null;
      }

      const value = getDescendants(background, 'srgbClr')[0]?.getAttribute('val');
      return value && /^[0-9A-Fa-f]{6}$/.test(value) ? value.toUpperCase() : null;
    } catch {
      return null;
    }
  }

  /**
   * Report whether a top-level shape is an embedded picture (`p:pic` with an
   * `<a:blip>`). Read directly from the slide OOXML so image-only menu items
   * are gated on the authoritative model rather than the rendered SVG.
   */
  isImageShape(slideIndex: number, shapeIndex: number): boolean {
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElement(slideDoc, shapeIndex);
      if (shape.localName !== 'pic') return false;
      return getDescendants(shape, 'blip').some((blip) => Boolean(getBlipEmbedId(blip)));
    } catch {
      return false;
    }
  }

  /**
   * Read the current inset crop (`<a:srcRect>`) of a picture as percentages.
   * Returns zeros when the picture is uncropped, or null when the shape is not
   * a picture.
   */
  getImageCrop(slideIndex: number, shapeIndex: number): ImageCrop | null {
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElement(slideDoc, shapeIndex);
      if (shape.localName !== 'pic') return null;

      const srcRect = getDescendants(shape, 'srcRect')[0];
      const read = (attribute: string): number => {
        const value = Number(srcRect?.getAttribute(attribute));
        return Number.isFinite(value) ? value / 1000 : 0;
      };
      return { left: read('l'), top: read('t'), right: read('r'), bottom: read('b') };
    } catch {
      return null;
    }
  }

  /**
   * Apply an inset crop to a picture via `<a:srcRect>` inside its `<a:blipFill>`.
   * Percentages are stored in OOXML 1000ths-of-a-percent units. Position and
   * size are untouched.
   */
  async setImageCrop(slideIndex: number, shapeIndex: number, crop: ImageCrop): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      if (shape.localName !== 'pic') {
        throw new Error('The selected object is not an image.');
      }

      const blipFill = getDescendants(shape, 'blipFill')[0];
      if (!blipFill) {
        throw new Error('The selected image has no picture fill to crop.');
      }

      let srcRect = getElementChildren(blipFill)
        .find((element) => element.localName === 'srcRect' && element.namespaceURI === DRAWINGML_NAMESPACE);
      if (!srcRect) {
        srcRect = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:srcRect');
        // CT_BlipFillProperties order: blip, srcRect, (tile | stretch).
        const fillMode = getElementChildren(blipFill)
          .find((element) => element.localName === 'stretch' || element.localName === 'tile');
        if (fillMode) {
          blipFill.insertBefore(srcRect, fillMode);
        } else {
          blipFill.appendChild(srcRect);
        }
      }

      srcRect.setAttribute('l', String(cropPercentToPermille(crop.left)));
      srcRect.setAttribute('t', String(cropPercentToPermille(crop.top)));
      srcRect.setAttribute('r', String(cropPercentToPermille(crop.right)));
      srcRect.setAttribute('b', String(cropPercentToPermille(crop.bottom)));
      return true;
    });
  }

  /**
   * Reset a picture to its original appearance: removes any inset crop
   * (`<a:srcRect>`) and common recolor effects (duotone, biLevel, grayscl, lum,
   * clrChange) from the `<a:blip>`. Position, size, and the embedded image are
   * preserved.
   */
  async resetImage(slideIndex: number, shapeIndex: number): Promise<void> {
    const recolorEffects = new Set(['duotone', 'biLevel', 'grayscl', 'lum', 'clrChange']);
    await this.editSlideShape(slideIndex, shapeIndex, (shape) => {
      if (shape.localName !== 'pic') {
        throw new Error('The selected object is not an image.');
      }

      const blipFill = getDescendants(shape, 'blipFill')[0];
      if (!blipFill) return false;

      let changed = false;
      for (const srcRect of getElementChildren(blipFill).filter((element) => element.localName === 'srcRect')) {
        blipFill.removeChild(srcRect);
        changed = true;
      }

      const blip = getElementChildren(blipFill).find((element) => element.localName === 'blip');
      if (blip) {
        for (const effect of getElementChildren(blip).filter((element) => recolorEffects.has(element.localName))) {
          blip.removeChild(effect);
          changed = true;
        }
      }
      return changed;
    });
  }

  /**
   * Toggle a horizontal or vertical flip on a shape by editing the `flipH` /
   * `flipV` attributes of its `<a:xfrm>`. The renderer's transform API does not
   * expose flip, so this is applied directly in OOXML.
   */
  async flipShape(slideIndex: number, shapeIndex: number, axis: 'horizontal' | 'vertical'): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape) => {
      const xfrm = getDescendants(shape, 'xfrm')[0];
      if (!xfrm) {
        throw new Error('The selected object cannot be flipped.');
      }

      const attribute = axis === 'horizontal' ? 'flipH' : 'flipV';
      const current = xfrm.getAttribute(attribute);
      if (current === '1' || current === 'true') {
        xfrm.removeAttribute(attribute);
      } else {
        xfrm.setAttribute(attribute, '1');
      }
      return true;
    });
  }

  /**
   * Swap the picture's embedded image for new bytes while preserving its
   * position, size, and crop. A fresh media part and slide relationship are
   * added, the `<a:blip r:embed>` is repointed, and the content-type default is
   * registered for the new extension. The previous media part is left in place.
   */
  async replaceImage(
    slideIndex: number,
    shapeIndex: number,
    bytes: Uint8Array,
    mimeType: string
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
    if (shape.localName !== 'pic') {
      throw new Error('The selected object is not an image.');
    }

    const blip = getDescendants(shape, 'blip')[0];
    if (!blip) {
      throw new Error('The selected image has no embedded picture data.');
    }

    const extension = imageExtensionForMime(mimeType);
    const mediaPath = nextImageMediaPath(zip.textFiles, zip.binaryFiles, extension);
    const relationship = this.buildSlideImageRelationship(zip, slideIndex, mediaPath);
    setBlipEmbedId(blip, relationship.relationshipId);

    const contentTypesDoc = parseXml(
      zip.textFiles.get('[Content_Types].xml') ?? '<Types/>',
      '[Content_Types].xml'
    );
    ensureDefaultContentType(contentTypesDoc, extension, contentTypeForImageExtension(extension));

    const textModifications = new Map<string, string>([
      [slidePath, serializeXml(slideDoc)],
      [relationship.relationshipsPath, relationship.relationshipsXml],
      ['[Content_Types].xml', serializeXml(contentTypesDoc)]
    ]);
    const binaryModifications = new Map<string, Uint8Array>([[mediaPath, bytes]]);

    const patched = await buildZip(rawExport, textModifications, undefined, binaryModifications);
    await this.reloadFromBuffer(patched, this.slideCountValue);
  }

  /**
   * Read the bytes of a picture's embedded image by resolving its
   * `<a:blip r:embed>` relationship to a media part. Returns null when the
   * shape is not a picture or the media part cannot be located.
   */
  async getShapeImageData(slideIndex: number, shapeIndex: number): Promise<ShapeImageData | null> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) return null;

    const shape = getShapeElement(parseXml(slideXml, slidePath), shapeIndex);
    if (shape.localName !== 'pic') return null;

    const blip = getDescendants(shape, 'blip')[0];
    const relationshipId = blip ? getBlipEmbedId(blip) : null;
    if (!relationshipId) return null;

    const relationshipsPath = getSlideRelationshipsPath(slideIndex);
    const relationshipsXml = zip.textFiles.get(relationshipsPath);
    if (!relationshipsXml) return null;

    const relationship = getDescendants(parseXml(relationshipsXml, relationshipsPath), 'Relationship')
      .find((element) => element.getAttribute('Id') === relationshipId);
    const target = relationship?.getAttribute('Target');
    if (!target || relationship?.getAttribute('TargetMode') === 'External') return null;

    const mediaPath = resolvePartPath(slidePath, target);
    const bytes = zip.binaryFiles.get(mediaPath);
    if (!bytes) return null;

    return {
      bytes: bytes.slice(),
      mimeType: contentTypeForImageExtension(getPartExtension(mediaPath))
    };
  }

  /**
   * Set the slide background to a stretched image, mirroring how
   * {@link setSlideBackgroundColor} rebuilds `<p:bg><p:bgPr>`. A new media part
   * and slide relationship are created and any existing background is replaced.
   */
  async setSlideBackgroundImage(
    slideIndex: number,
    bytes: Uint8Array,
    mimeType: string
  ): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const commonSlide = getDescendants(slideDoc, 'cSld')[0];
    if (!commonSlide) {
      throw new Error('Slide is missing its common slide data.');
    }

    const extension = imageExtensionForMime(mimeType);
    const mediaPath = nextImageMediaPath(zip.textFiles, zip.binaryFiles, extension);
    const relationship = this.buildSlideImageRelationship(zip, slideIndex, mediaPath);

    const presentationNamespace = slideDoc.documentElement.namespaceURI;
    for (const existingBackground of getElementChildren(commonSlide).filter((element) => element.localName === 'bg')) {
      commonSlide.removeChild(existingBackground);
    }

    const background = slideDoc.createElementNS(presentationNamespace, 'p:bg');
    const backgroundProperties = slideDoc.createElementNS(presentationNamespace, 'p:bgPr');
    const blipFill = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:blipFill');
    const blip = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:blip');
    setBlipEmbedId(blip, relationship.relationshipId);
    blipFill.appendChild(blip);
    const stretch = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:stretch');
    stretch.appendChild(slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:fillRect'));
    blipFill.appendChild(stretch);
    backgroundProperties.appendChild(blipFill);
    backgroundProperties.appendChild(slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:effectLst'));
    background.appendChild(backgroundProperties);
    commonSlide.insertBefore(background, commonSlide.firstChild);

    const contentTypesDoc = parseXml(
      zip.textFiles.get('[Content_Types].xml') ?? '<Types/>',
      '[Content_Types].xml'
    );
    ensureDefaultContentType(contentTypesDoc, extension, contentTypeForImageExtension(extension));

    const textModifications = new Map<string, string>([
      [slidePath, serializeXml(slideDoc)],
      [relationship.relationshipsPath, relationship.relationshipsXml],
      ['[Content_Types].xml', serializeXml(contentTypesDoc)]
    ]);
    const binaryModifications = new Map<string, Uint8Array>([[mediaPath, bytes]]);

    const patched = await buildZip(rawExport, textModifications, undefined, binaryModifications);
    await this.reloadFromBuffer(patched, this.slideCountValue);
  }

  /**
   * Compose a new image relationship for a slide's `.rels` part. Returns the
   * new relationship id alongside the serialized relationships XML and its part
   * path, so the caller can include them in a single buildZip pass (the
   * extracted zip is not mutated here).
   */
  private buildSlideImageRelationship(
    zip: { textFiles: Map<string, string>; binaryFiles: Map<string, Uint8Array> },
    slideIndex: number,
    mediaPath: string
  ): { relationshipId: string; relationshipsPath: string; relationshipsXml: string } {
    const relationshipsPath = getSlideRelationshipsPath(slideIndex);
    const relationshipsXml = zip.textFiles.get(relationshipsPath);
    const relationships = relationshipsXml
      ? parseXml(relationshipsXml, relationshipsPath)
      : createRelationshipsDocument();

    const relationshipId = nextRelationshipId(relationships);
    const relationship = relationships.createElementNS(
      relationships.documentElement.namespaceURI,
      'Relationship'
    );
    relationship.setAttribute('Id', relationshipId);
    relationship.setAttribute('Type', IMAGE_RELATIONSHIP_TYPE);
    relationship.setAttribute('Target', `../media/${mediaPath.split('/').pop()}`);
    relationships.documentElement.appendChild(relationship);

    return {
      relationshipId,
      relationshipsPath,
      relationshipsXml: serializeXml(relationships)
    };
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
