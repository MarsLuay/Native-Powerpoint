import { buildZip, extractZip, pxToEmu, type ZipContents } from 'pptx-svg';
import {
  CHART_INSERT_CHART_RELS_XML,
  CHART_INSERT_CHART_XML,
  CHART_INSERT_FRAME_TEMPLATE,
  CHART_INSERT_WORKBOOK_BASE64
} from './chartInsertTemplate';

const DRAWINGML_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PACKAGE_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const CHART_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const SHAPE_ELEMENT_NAMES = new Set(['cxnSp', 'graphicFrame', 'grpSp', 'pic', 'sp']);

export type ParagraphListStyle = 'none' | 'bullet' | 'number';

export interface SlideInsertionResult {
  buffer: ArrayBuffer;
  shapeIndex: number;
}

function parseXml(contents: string, partPath: string): XMLDocument {
  const document = new DOMParser().parseFromString(contents, 'application/xml');
  if (document.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`Could not parse PowerPoint XML part: ${partPath}`);
  }
  return document;
}

function serializeXml(document: XMLDocument): string {
  return new XMLSerializer().serializeToString(document);
}

function getDescendants(element: Element | XMLDocument, localName: string): Element[] {
  return Array.from(element.getElementsByTagNameNS('*', localName));
}

function getElementChildren(element: Element | undefined): Element[] {
  return Array.from(element?.childNodes ?? []).filter((node): node is Element => node.nodeType === 1);
}

function getSlidePath(slideIndex: number): string {
  return `ppt/slides/slide${slideIndex + 1}.xml`;
}

function getRelationshipsPath(partPath: string): string {
  const slashIndex = partPath.lastIndexOf('/');
  const directory = slashIndex === -1 ? '' : `${partPath.slice(0, slashIndex + 1)}`;
  const fileName = slashIndex === -1 ? partPath : partPath.slice(slashIndex + 1);
  return `${directory}_rels/${fileName}.rels`;
}

function getRequiredTextFile(zip: ZipContents, partPath: string): string {
  const contents = zip.textFiles.get(partPath);
  if (!contents) throw new Error(`Missing PowerPoint XML part: ${partPath}`);
  return contents;
}

function getShapeTree(document: XMLDocument): Element {
  const shapeTree = getDescendants(document, 'spTree')[0];
  if (!shapeTree) throw new Error('Could not find the slide shape tree.');
  return shapeTree;
}

function getShapeChildren(shapeTree: Element): Element[] {
  return getElementChildren(shapeTree).filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName));
}

function getNextRelationshipId(document: XMLDocument): string {
  const usedIds = new Set(
    getDescendants(document, 'Relationship')
      .map((relationship) => relationship.getAttribute('Id'))
      .filter((id): id is string => Boolean(id))
  );
  let nextId = 1;
  while (usedIds.has(`rId${nextId}`)) nextId++;
  return `rId${nextId}`;
}

function getNextNumberedPartPath(prefix: string, suffix: string, zip: ZipContents, extra: Iterable<string> = []): string {
  let nextNumber = 1;
  const exists = (path: string) =>
    zip.textFiles.has(path) || zip.binaryFiles.has(path) || [...extra].includes(path);
  while (exists(`${prefix}${nextNumber}${suffix}`)) nextNumber++;
  return `${prefix}${nextNumber}${suffix}`;
}

function ensureContentTypeOverride(
  contentTypesDocument: XMLDocument,
  partPath: string,
  contentType: string
): void {
  const partName = `/${partPath}`;
  const alreadyRegistered = getDescendants(contentTypesDocument, 'Override')
    .some((override) => override.getAttribute('PartName') === partName);
  if (alreadyRegistered) return;

  const override = contentTypesDocument.createElementNS(CONTENT_TYPES_NAMESPACE, 'Override');
  override.setAttribute('PartName', partName);
  override.setAttribute('ContentType', contentType);
  contentTypesDocument.documentElement.appendChild(override);
}

function ensureContentTypeDefault(
  contentTypesDocument: XMLDocument,
  extension: string,
  contentType: string
): void {
  const normalized = extension.toLowerCase();
  const alreadyRegistered = getDescendants(contentTypesDocument, 'Default')
    .some((entry) => entry.getAttribute('Extension')?.toLowerCase() === normalized);
  if (alreadyRegistered) return;

  const entry = contentTypesDocument.createElementNS(CONTENT_TYPES_NAMESPACE, 'Default');
  entry.setAttribute('Extension', normalized);
  entry.setAttribute('ContentType', contentType);
  contentTypesDocument.documentElement.appendChild(entry);
}

function assignUniqueNonVisualIds(destinationSlide: XMLDocument, shape: Element, name: string): void {
  const usedIds = new Set(
    getDescendants(destinationSlide, 'cNvPr')
      .filter((element) => !shape.contains(element))
      .map((element) => Number(element.getAttribute('id')))
      .filter(Number.isFinite)
  );
  let nextId = Math.max(0, ...usedIds) + 1;

  for (const nonVisualProperties of getDescendants(shape, 'cNvPr')) {
    while (usedIds.has(nextId)) nextId++;
    nonVisualProperties.setAttribute('id', String(nextId));
    nonVisualProperties.setAttribute('name', name);
    usedIds.add(nextId++);
  }
}

function escapeXmlText(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildTableCellXml(text: string, fillHex: string): string {
  return `<a:tc>
    <a:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>${escapeXmlText(text)}</a:t></a:r><a:endParaRPr lang="en-US" sz="1800"/></a:p>
    </a:txBody>
    <a:tcPr><a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill></a:tcPr>
  </a:tc>`;
}

function buildTableGraphicFrameXml(rows: number, cols: number): string {
  const safeRows = Math.max(1, Math.min(rows, 20));
  const safeCols = Math.max(1, Math.min(cols, 10));
  const tableWidth = pxToEmu(520);
  const colWidth = Math.floor(tableWidth / safeCols);
  const rowHeight = pxToEmu(36);
  const tableHeight = rowHeight * safeRows;
  const x = pxToEmu(120);
  const y = pxToEmu(120);

  const gridCols = Array.from({ length: safeCols }, () => `<a:gridCol w="${colWidth}"/>`).join('');
  const tableRows: string[] = [];

  for (let rowIndex = 0; rowIndex < safeRows; rowIndex++) {
    const cells: string[] = [];
    for (let colIndex = 0; colIndex < safeCols; colIndex++) {
      const isHeader = rowIndex === 0;
      const text = isHeader ? `Column ${colIndex + 1}` : '';
      const fill = isHeader ? 'E8F0FE' : 'FFFFFF';
      cells.push(buildTableCellXml(text, fill));
    }
    tableRows.push(`<a:tr h="${rowHeight}">${cells.join('')}</a:tr>`);
  }

  return `<p:graphicFrame>
    <p:nvGraphicFramePr>
      <p:cNvPr id="2" name="Table"/>
      <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
      <p:nvPr/>
    </p:nvGraphicFramePr>
    <p:xfrm>
      <a:off x="${x}" y="${y}"/>
      <a:ext cx="${tableWidth}" cy="${tableHeight}"/>
    </p:xfrm>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        <a:tbl>
          <a:tblPr firstRow="1" bandRow="1"/>
          <a:tblGrid>${gridCols}</a:tblGrid>
          ${tableRows.join('')}
        </a:tbl>
      </a:graphicData>
    </a:graphic>
  </p:graphicFrame>`;
}

function appendGraphicFrame(
  slideDocument: XMLDocument,
  frameXml: string,
  displayName: string
): number {
  const wrapper = parseXml(
    `<p:slide xmlns:a="${DRAWINGML_NAMESPACE}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${frameXml}</p:slide>`,
    '(table frame)'
  );
  const imported = slideDocument.importNode(getDescendants(wrapper, 'graphicFrame')[0]!, true);
  const shapeTree = getShapeTree(slideDocument);
  assignUniqueNonVisualIds(slideDocument, imported, displayName);
  shapeTree.appendChild(imported);
  return getShapeChildren(shapeTree).length - 1;
}

function decodeBase64(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function getRelativePartPath(sourcePartPath: string, targetPartPath: string): string {
  const sourceParts = sourcePartPath.split('/');
  sourceParts.pop();
  const targetParts = targetPartPath.split('/');

  while (sourceParts[0] && sourceParts[0] === targetParts[0]) {
    sourceParts.shift();
    targetParts.shift();
  }

  return [...sourceParts.map(() => '..'), ...targetParts].join('/');
}

export async function insertTableIntoPresentation(
  buffer: ArrayBuffer,
  slideIndex: number,
  rows: number,
  cols: number
): Promise<SlideInsertionResult> {
  const zip = await extractZip(buffer);
  const slidePath = getSlidePath(slideIndex);
  const slideDocument = parseXml(getRequiredTextFile(zip, slidePath), slidePath);
  const shapeIndex = appendGraphicFrame(
    slideDocument,
    buildTableGraphicFrameXml(rows, cols),
    'Table'
  );
  const patched = await buildZip(buffer, new Map([[slidePath, serializeXml(slideDocument)]]));
  return { buffer: patched, shapeIndex };
}

export async function insertChartIntoPresentation(
  buffer: ArrayBuffer,
  slideIndex: number
): Promise<SlideInsertionResult> {
  const zip = await extractZip(buffer);
  const slidePath = getSlidePath(slideIndex);
  const slideDocument = parseXml(getRequiredTextFile(zip, slidePath), slidePath);
  const contentTypesDocument = parseXml(getRequiredTextFile(zip, '[Content_Types].xml'), '[Content_Types].xml');
  const relationshipsPath = getRelationshipsPath(slidePath);
  const relationshipsDocument = zip.textFiles.has(relationshipsPath)
    ? parseXml(getRequiredTextFile(zip, relationshipsPath), relationshipsPath)
    : parseXml(`<Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}"/>`, relationshipsPath);

  const chartPath = getNextNumberedPartPath('ppt/charts/chart', '.xml', zip);
  const chartRelsPath = getRelationshipsPath(chartPath);
  const workbookPath = getNextNumberedPartPath(
    'ppt/embeddings/Microsoft_Excel_Worksheet',
    '.xlsx',
    zip,
    [chartPath, chartRelsPath]
  );

  const chartRelationshipId = getNextRelationshipId(relationshipsDocument);
  const chartRelationship = relationshipsDocument.createElementNS(
    PACKAGE_RELATIONSHIP_NAMESPACE,
    'Relationship'
  );
  chartRelationship.setAttribute('Id', chartRelationshipId);
  chartRelationship.setAttribute('Type', CHART_RELATIONSHIP_TYPE);
  chartRelationship.setAttribute('Target', getRelativePartPath(slidePath, chartPath));
  relationshipsDocument.documentElement.appendChild(chartRelationship);

  const frameDocument = parseXml(
    `<p:slide xmlns:a="${DRAWINGML_NAMESPACE}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${CHART_INSERT_FRAME_TEMPLATE}</p:slide>`,
    '(chart frame)'
  );
  const importedFrame = slideDocument.importNode(getDescendants(frameDocument, 'graphicFrame')[0]!, true);
  const chartReference = getDescendants(importedFrame, 'chart')[0];
  chartReference?.setAttributeNS(
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'r:id',
    chartRelationshipId
  );
  assignUniqueNonVisualIds(slideDocument, importedFrame, 'Chart');
  getShapeTree(slideDocument).appendChild(importedFrame);
  const shapeIndex = getShapeChildren(getShapeTree(slideDocument)).length - 1;

  const chartRelsDocument = parseXml(CHART_INSERT_CHART_RELS_XML, chartRelsPath);
  const workbookRelationship = getDescendants(chartRelsDocument, 'Relationship')[0];
  workbookRelationship?.setAttribute(
    'Target',
    getRelativePartPath(chartPath, workbookPath)
  );

  ensureContentTypeOverride(
    contentTypesDocument,
    chartPath,
    'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
  );
  ensureContentTypeDefault(
    contentTypesDocument,
    'xlsx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );

  const textModifications = new Map<string, string>([
    [slidePath, serializeXml(slideDocument)],
    [relationshipsPath, serializeXml(relationshipsDocument)],
    [chartPath, CHART_INSERT_CHART_XML],
    [chartRelsPath, serializeXml(chartRelsDocument)],
    ['[Content_Types].xml', serializeXml(contentTypesDocument)]
  ]);
  const binaryModifications = new Map<string, Uint8Array>([
    [workbookPath, decodeBase64(CHART_INSERT_WORKBOOK_BASE64)]
  ]);

  const patched = await buildZip(buffer, textModifications, undefined, binaryModifications);
  return { buffer: patched, shapeIndex };
}

function getShapeElement(slideDocument: XMLDocument, shapeIndex: number): Element {
  const shape = getShapeChildren(getShapeTree(slideDocument))[shapeIndex];
  if (!shape) throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
  return shape;
}

function resolveTextShapeIndex(slideDocument: XMLDocument, shapeIndex: number): number {
  const shapes = getShapeChildren(getShapeTree(slideDocument));
  if (shapes[shapeIndex] && getDrawingParagraphs(shapes[shapeIndex]).length > 0) {
    return shapeIndex;
  }

  const textShapeIndices = shapes
    .map((shape, index) => ({ index, paragraphCount: getDrawingParagraphs(shape).length }))
    .filter((entry) => entry.paragraphCount > 0)
    .map((entry) => entry.index);

  if (textShapeIndices.length === 0) {
    throw new Error('Could not find a text box to format.');
  }

  if (textShapeIndices.includes(shapeIndex)) {
    return shapeIndex;
  }

  const nextTextShape = textShapeIndices.find((index) => index >= shapeIndex);
  return nextTextShape ?? textShapeIndices[textShapeIndices.length - 1]!;
}

function getGraphicFrameSignature(frame: Element): string {
  const chart = getDescendants(frame, 'chart')[0];
  const table = getDescendants(frame, 'tbl')[0];
  if (chart) {
    return `chart:${chart.getAttributeNS(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'id'
    ) ?? chart.getAttribute('r:id') ?? ''}`;
  }
  if (table) {
    const rows = getDescendants(table, 'tr').length;
    const cols = getDescendants(table, 'gridCol').length;
    return `table:${rows}x${cols}`;
  }
  return `frame:${frame.getAttribute('name') ?? ''}`;
}

function mergePreservedGraphicFrames(previousXml: string, exportedXml: string, slidePath: string): string {
  const previousDocument = parseXml(previousXml, slidePath);
  const exportedDocument = parseXml(exportedXml, slidePath);
  const previousFrames = getShapeChildren(getShapeTree(previousDocument)).filter(
    (shape) => shape.localName === 'graphicFrame'
  );
  const exportedTree = getShapeTree(exportedDocument);
  const exportedFrames = getShapeChildren(exportedTree).filter((shape) => shape.localName === 'graphicFrame');
  const exportedSignatures = new Set(exportedFrames.map(getGraphicFrameSignature));

  let changed = false;
  for (const frame of previousFrames) {
    const signature = getGraphicFrameSignature(frame);
    if (exportedSignatures.has(signature)) continue;
    exportedTree.appendChild(exportedDocument.importNode(frame, true));
    exportedSignatures.add(signature);
    changed = true;
  }

  return changed ? serializeXml(exportedDocument) : exportedXml;
}

export async function mergeSlideGraphicFramesFromBuffer(
  previousBuffer: ArrayBuffer,
  exportedBuffer: ArrayBuffer,
  slideIndex: number
): Promise<ArrayBuffer> {
  const slidePath = getSlidePath(slideIndex);
  const [previousZip, exportedZip] = await Promise.all([
    extractZip(previousBuffer),
    extractZip(exportedBuffer)
  ]);
  const previousXml = previousZip.textFiles.get(slidePath);
  const exportedXml = exportedZip.textFiles.get(slidePath);
  if (!previousXml || !exportedXml) return exportedBuffer;

  const mergedXml = mergePreservedGraphicFrames(previousXml, exportedXml, slidePath);
  return mergedXml === exportedXml
    ? exportedBuffer
    : buildZip(exportedBuffer, new Map([[slidePath, mergedXml]]));
}

const PRESERVED_PACKAGE_PART_PATTERNS = [
  /^\[Content_Types\]\.xml$/,
  /^ppt\/charts\//,
  /^ppt\/embeddings\//,
  /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/
];

function shouldPreservePackagePart(partPath: string): boolean {
  return PRESERVED_PACKAGE_PART_PATTERNS.some((pattern) => pattern.test(partPath));
}

export async function mergeMissingPackageParts(
  previousBuffer: ArrayBuffer,
  exportedBuffer: ArrayBuffer
): Promise<ArrayBuffer> {
  const [previousZip, exportedZip] = await Promise.all([
    extractZip(previousBuffer),
    extractZip(exportedBuffer)
  ]);
  const textModifications = new Map<string, string>();
  const binaryModifications = new Map<string, Uint8Array>();

  for (const [partPath, contents] of previousZip.textFiles) {
    if (!shouldPreservePackagePart(partPath) || exportedZip.textFiles.has(partPath)) continue;
    textModifications.set(partPath, contents);
  }

  for (const [partPath, contents] of previousZip.binaryFiles) {
    if (!shouldPreservePackagePart(partPath) || exportedZip.binaryFiles.has(partPath)) continue;
    binaryModifications.set(partPath, contents);
  }

  if (textModifications.size === 0 && binaryModifications.size === 0) {
    return exportedBuffer;
  }

  return buildZip(exportedBuffer, textModifications, undefined, binaryModifications);
}

function getDrawingParagraphs(container: Element): Element[] {
  const textBody = getDescendants(container, 'txBody')
    .find((element) => element.namespaceURI === DRAWINGML_NAMESPACE || element.namespaceURI === container.namespaceURI);
  const scope = textBody ?? container;
  return getElementChildren(scope).filter(
    (element) => element.localName === 'p' && element.namespaceURI === DRAWINGML_NAMESPACE
  );
}

function ensureParagraphProperties(paragraph: Element): Element {
  let properties = getElementChildren(paragraph).find(
    (element) => element.localName === 'pPr' && element.namespaceURI === DRAWINGML_NAMESPACE
  );
  if (!properties) {
    properties = paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:pPr');
    paragraph.insertBefore(properties, paragraph.firstChild);
  }
  return properties;
}

function clearListMarkers(properties: Element): void {
  for (const child of getElementChildren(properties)) {
    if (child.localName === 'buChar' || child.localName === 'buAutoNum' || child.localName === 'buNone') {
      properties.removeChild(child);
    }
  }
}

function applyListStyleToParagraph(paragraph: Element, style: ParagraphListStyle): void {
  const properties = ensureParagraphProperties(paragraph);
  clearListMarkers(properties);

  if (style === 'none') {
    const marker = paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:buNone');
    properties.appendChild(marker);
    return;
  }

  if (style === 'bullet') {
    const marker = paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:buChar');
    marker.setAttribute('char', '•');
    properties.appendChild(marker);
    return;
  }

  const marker = paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:buAutoNum');
  marker.setAttribute('type', 'arabicPeriod');
  properties.appendChild(marker);
}

export async function applyParagraphListStyle(
  buffer: ArrayBuffer,
  slideIndex: number,
  shapeIndex: number,
  paragraphIndex: number,
  style: ParagraphListStyle
): Promise<ArrayBuffer> {
  const slidePath = getSlidePath(slideIndex);
  const zip = await extractZip(buffer);
  const slideDocument = parseXml(getRequiredTextFile(zip, slidePath), slidePath);
  const shape = getShapeElement(slideDocument, resolveTextShapeIndex(slideDocument, shapeIndex));
  const paragraphs = getDrawingParagraphs(shape);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new Error('Could not find the selected text paragraph.');
  }

  applyListStyleToParagraph(paragraph, style);
  return buildZip(buffer, new Map([[slidePath, serializeXml(slideDocument)]]));
}
