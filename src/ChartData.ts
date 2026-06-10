import { buildZip, extractZip, type ZipContents } from 'pptx-svg';

const CHART_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SPREADSHEET_NAMESPACE =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

type ChartDataValueKind = 'number' | 'string';

interface ChartDataSource {
  formula: string;
  kind: ChartDataValueKind;
  values: string[];
}

interface ChartPointLabelBinding {
  indexes: number[];
  source: ChartDataSource;
}

interface ChartDataSeriesDescriptor {
  pointLabelBindings: ChartPointLabelBinding[];
  values: ChartDataSource;
}

interface CellRange {
  cellReferences: string[];
  sheetName: string;
}

interface WorkbookCellEdit {
  kind: ChartDataValueKind;
  values: string[];
}

export interface ChartDataSeries {
  name: string;
  pointLabels: string[] | null;
  values: string[];
}

export interface ChartDataGrid {
  categoryLabel: string;
  categories: string[];
  editable: boolean;
  reason: string;
  series: ChartDataSeries[];
}

export interface ChartDataUpdateSeries {
  pointLabels: string[] | null;
  values: string[];
}

export interface ChartDataUpdate {
  categories: string[];
  series: ChartDataUpdateSeries[];
}

export interface ChartDataDescriptor {
  categories: ChartDataSource | null;
  chartPath: string;
  grid: ChartDataGrid;
  series: ChartDataSeriesDescriptor[];
  workbookPath: string | null;
}

function parseXml(contents: string, partPath: string): XMLDocument {
  const doc = new DOMParser().parseFromString(contents, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`Could not parse Open XML part: ${partPath}`);
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

function getDirectChild(element: Element | undefined, localName: string): Element | undefined {
  return getElementChildren(element).find((child) => child.localName === localName);
}

function getValAttribute(element: Element | undefined, localName: string): string | null {
  return element ? getDescendants(element, localName)[0]?.getAttribute('val') ?? null : null;
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

function getRelationshipsPath(partPath: string): string {
  const parts = partPath.split('/');
  const fileName = parts.pop();
  return [...parts, '_rels', `${fileName}.rels`].join('/');
}

function getFormula(element: Element | undefined): string | null {
  const formula = element ? getDescendants(element, 'f')[0]?.textContent?.trim() : null;
  return formula || null;
}

function getCache(element: Element | undefined): Element | undefined {
  return getElementChildren(element)
    .find((child) => /(?:Cache|Lit)$/.test(child.localName));
}

function getCachedValues(cache: Element | undefined, expectedCount = 0): string[] {
  const values: string[] = Array.from({ length: expectedCount }, () => '');
  if (!cache) return values;

  for (const [fallbackIndex, point] of getDescendants(cache, 'pt').entries()) {
    const index = Number(point.getAttribute('idx') ?? fallbackIndex);
    if (!Number.isInteger(index) || index < 0) continue;
    while (values.length <= index) values.push('');
    values[index] = getDirectChild(point, 'v')?.textContent ?? '';
  }

  return values;
}

function parseColumnNumber(columnName: string): number {
  let value = 0;
  for (const character of columnName.toUpperCase()) {
    value = value * 26 + character.charCodeAt(0) - 64;
  }
  return value;
}

function getColumnName(columnNumber: number): string {
  let value = columnNumber;
  let name = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function parseCellRange(formula: string): CellRange | null {
  const match = formula.trim().match(
    /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i
  );
  if (!match) return null;

  const sheetName = (match[1] ?? match[2] ?? '').replace(/''/g, "'");
  const startColumn = parseColumnNumber(match[3] ?? '');
  const startRow = Number(match[4]);
  const endColumn = parseColumnNumber(match[5] ?? match[3] ?? '');
  const endRow = Number(match[6] ?? match[4]);
  if (!sheetName || startColumn < 1 || startRow < 1 || endColumn < startColumn || endRow < startRow) {
    return null;
  }

  const cellReferences: string[] = [];
  for (let row = startRow; row <= endRow; row++) {
    for (let column = startColumn; column <= endColumn; column++) {
      cellReferences.push(`${getColumnName(column)}${row}`);
    }
  }

  return { cellReferences, sheetName };
}

function getReferenceSource(element: Element | undefined): ChartDataSource | null {
  const reference = getElementChildren(element)
    .find((child) => child.localName === 'numRef' || child.localName === 'strRef')
    ?? element;
  const formula = getFormula(reference);
  if (!reference || !formula) return null;

  const range = parseCellRange(formula);
  if (!range) return null;

  const kind: ChartDataValueKind =
    reference.localName === 'numRef' || getCache(reference)?.localName === 'numCache'
      ? 'number'
      : 'string';
  return {
    formula,
    kind,
    values: getCachedValues(getCache(reference), range.cellReferences.length)
  };
}

function getSeriesName(series: Element, seriesIndex: number): string {
  const text = getDirectChild(series, 'tx');
  const source = getReferenceSource(text);
  return source?.values[0]
    || getDescendants(text ?? series, 'v')[0]?.textContent
    || `Series ${seriesIndex + 1}`;
}

function getPointLabelBindings(series: Element, rowCount: number): ChartPointLabelBinding[] {
  const rangeElement = getDescendants(series, 'datalabelsRange')[0];
  const rangeSource = getReferenceSource(rangeElement);
  if (rangeSource) {
    return [{
      indexes: Array.from({ length: rowCount }, (_, index) => index),
      source: rangeSource
    }];
  }

  const bindings: ChartPointLabelBinding[] = [];
  for (const label of getDescendants(series, 'dLbl')) {
    const index = Number(getValAttribute(label, 'idx'));
    const source = getReferenceSource(getDirectChild(label, 'tx'));
    if (Number.isInteger(index) && index >= 0 && index < rowCount && source) {
      bindings.push({ indexes: [index], source });
    }
  }
  return bindings;
}

function readPointLabels(bindings: ChartPointLabelBinding[], rowCount: number): string[] | null {
  if (bindings.length === 0) return null;

  const labels = Array.from({ length: rowCount }, () => '');
  for (const binding of bindings) {
    binding.indexes.forEach((index, sourceIndex) => {
      labels[index] = binding.source.values[sourceIndex] ?? '';
    });
  }
  return labels;
}

function getEmbeddedWorkbookPath(
  zip: ZipContents,
  chartDoc: XMLDocument,
  chartPath: string
): { path: string | null; reason: string } {
  const externalData = getDescendants(chartDoc, 'externalData')[0];
  const relationshipId =
    externalData?.getAttributeNS(RELATIONSHIP_NAMESPACE, 'id')
    || externalData?.getAttribute('r:id');
  if (!relationshipId) {
    return { path: null, reason: 'This chart has no embedded Excel workbook.' };
  }

  const relationshipsPath = getRelationshipsPath(chartPath);
  const relationshipsXml = zip.textFiles.get(relationshipsPath);
  if (!relationshipsXml) {
    return { path: null, reason: 'This chart workbook relationship is missing.' };
  }

  const relationships = getDescendants(parseXml(relationshipsXml, relationshipsPath), 'Relationship');
  const relationship = relationships.find((element) => element.getAttribute('Id') === relationshipId);
  if (!relationship) {
    return { path: null, reason: 'This chart workbook relationship could not be resolved.' };
  }

  if (relationship.getAttribute('TargetMode') === 'External') {
    return { path: null, reason: 'This chart uses an external workbook link. Edit it in the linked workbook.' };
  }

  const target = relationship.getAttribute('Target');
  if (!target || !relationship.getAttribute('Type')?.endsWith('/package')) {
    return { path: null, reason: 'This chart workbook type is not supported for in-place editing.' };
  }

  const workbookPath = resolvePartPath(chartPath, target);
  if (!zip.binaryFiles.has(workbookPath)) {
    return { path: null, reason: 'The embedded Excel workbook bytes are missing.' };
  }

  return { path: workbookPath, reason: '' };
}

function getUnsupportedReason(
  categories: ChartDataSource | null,
  seriesSources: ChartDataSeriesDescriptor[],
  categoryFormulas: string[],
  workbookReason: string
): string {
  if (workbookReason) return workbookReason;
  if (!categories) return 'This chart does not expose a source-backed category or X-value range.';
  if (seriesSources.length === 0) return 'This chart does not expose source-backed numeric series ranges.';
  if (categoryFormulas.some((formula) => formula !== categories.formula)) {
    return 'This chart uses different category or X-value ranges per series.';
  }
  return '';
}

export function getChartDataDescriptor(zip: ZipContents, chartPath: string): ChartDataDescriptor {
  const chartXml = zip.textFiles.get(chartPath);
  if (!chartXml) {
    throw new Error(`Missing chart XML part: ${chartPath}`);
  }

  const chartDoc = parseXml(chartXml, chartPath);
  const workbook = getEmbeddedWorkbookPath(zip, chartDoc, chartPath);
  const chartSeries = getDescendants(chartDoc, 'ser');
  const categories = chartSeries
    .map((series) => getReferenceSource(getDirectChild(series, 'cat') ?? getDirectChild(series, 'xVal')))
    .find((source): source is ChartDataSource => source !== null) ?? null;
  const categoryFormulas = chartSeries
    .map((series) => getReferenceSource(getDirectChild(series, 'cat') ?? getDirectChild(series, 'xVal'))?.formula)
    .filter((formula): formula is string => Boolean(formula));
  const rowCount = categories?.values.length ?? 0;
  const seriesDescriptors: ChartDataSeriesDescriptor[] = [];
  const gridSeries: ChartDataSeries[] = [];

  chartSeries.forEach((series, seriesIndex) => {
    const values = getReferenceSource(getDirectChild(series, 'val') ?? getDirectChild(series, 'yVal'));
    if (!values) return;

    const pointLabelBindings = getPointLabelBindings(series, rowCount);
    seriesDescriptors.push({ pointLabelBindings, values });
    gridSeries.push({
      name: getSeriesName(series, seriesIndex),
      pointLabels: readPointLabels(pointLabelBindings, rowCount),
      values: values.values
    });
  });

  const reason = getUnsupportedReason(categories, seriesDescriptors, categoryFormulas, workbook.reason);
  return {
    categories,
    chartPath,
    grid: {
      categoryLabel: chartSeries.some((series) => getDirectChild(series, 'xVal')) ? 'X value' : 'Category',
      categories: categories?.values ?? [],
      editable: reason === '',
      reason,
      series: gridSeries
    },
    series: seriesDescriptors,
    workbookPath: workbook.path
  };
}

function normalizeNumber(value: string, label: string): string {
  const trimmedValue = value.trim();
  if (trimmedValue === '') return '';

  const numericValue = Number(trimmedValue);
  if (!Number.isFinite(numericValue)) {
    throw new Error(`${label} must be a finite number.`);
  }
  return trimmedValue;
}

function normalizeValues(values: string[], kind: ChartDataValueKind, label: string): string[] {
  return kind === 'number'
    ? values.map((value, index) => normalizeNumber(value, `${label} row ${index + 1}`))
    : values;
}

function createChildElement(parent: Element, localName: string, template?: Element): Element {
  const namespace = template?.namespaceURI || parent.namespaceURI || CHART_NAMESPACE;
  const prefix = template?.prefix || parent.prefix;
  return parent.ownerDocument.createElementNS(namespace, prefix ? `${prefix}:${localName}` : localName);
}

function updateCache(cache: Element, values: string[]): void {
  const existingPoints = getDescendants(cache, 'pt');
  const pointTemplate = existingPoints[0];
  const valueTemplate = pointTemplate ? getDirectChild(pointTemplate, 'v') : undefined;
  for (const point of existingPoints) {
    point.parentNode?.removeChild(point);
  }

  let pointCount = getDirectChild(cache, 'ptCount');
  if (!pointCount) {
    pointCount = createChildElement(cache, 'ptCount');
    cache.insertBefore(pointCount, cache.firstChild);
  }
  pointCount.setAttribute('val', String(values.length));

  values.forEach((value, index) => {
    const point = createChildElement(cache, 'pt', pointTemplate);
    point.setAttribute('idx', String(index));
    const valueElement = createChildElement(point, 'v', valueTemplate);
    valueElement.textContent = value;
    point.appendChild(valueElement);
    cache.appendChild(point);
  });
}

function updateChartCaches(chartDoc: XMLDocument, formula: string, values: string[]): void {
  for (const formulaElement of getDescendants(chartDoc, 'f')) {
    if (formulaElement.textContent?.trim() !== formula) continue;
    const cache = getCache(formulaElement.parentElement ?? undefined);
    if (cache) updateCache(cache, values);
  }
}

function addWorkbookEdit(
  edits: Map<string, WorkbookCellEdit>,
  source: ChartDataSource,
  values: string[],
  label: string
): void {
  const normalizedValues = normalizeValues(values, source.kind, label);
  const range = parseCellRange(source.formula);
  if (!range || range.cellReferences.length !== normalizedValues.length) {
    throw new Error(`${label} does not match its embedded workbook range.`);
  }

  const existing = edits.get(source.formula);
  if (existing && existing.values.join('\u0000') !== normalizedValues.join('\u0000')) {
    throw new Error(`Conflicting edits target the same workbook range: ${source.formula}`);
  }
  edits.set(source.formula, { kind: source.kind, values: normalizedValues });
}

function getSheetPaths(workbookFiles: ZipContents): Map<string, string> {
  const workbookPath = 'xl/workbook.xml';
  const relationshipsPath = 'xl/_rels/workbook.xml.rels';
  const workbookXml = workbookFiles.textFiles.get(workbookPath);
  const relationshipsXml = workbookFiles.textFiles.get(relationshipsPath);
  if (!workbookXml || !relationshipsXml) {
    throw new Error('The embedded workbook is missing worksheet metadata.');
  }

  const workbookDoc = parseXml(workbookXml, workbookPath);
  const relationships = getDescendants(parseXml(relationshipsXml, relationshipsPath), 'Relationship');
  const sheetPaths = new Map<string, string>();
  for (const sheet of getDescendants(workbookDoc, 'sheet')) {
    const relationshipId =
      sheet.getAttributeNS(RELATIONSHIP_NAMESPACE, 'id')
      || sheet.getAttribute('r:id');
    const name = sheet.getAttribute('name');
    const relationship = relationships.find((element) => element.getAttribute('Id') === relationshipId);
    const target = relationship?.getAttribute('Target');
    if (name && target && relationship?.getAttribute('Type')?.endsWith('/worksheet')) {
      sheetPaths.set(name, resolvePartPath(workbookPath, target));
    }
  }
  return sheetPaths;
}

function parseCellReference(reference: string): { column: number; row: number } {
  const match = reference.match(/^([A-Z]{1,3})(\d+)$/i);
  if (!match) throw new Error(`Unsupported worksheet cell reference: ${reference}`);
  return { column: parseColumnNumber(match[1] ?? ''), row: Number(match[2]) };
}

function getOrCreateRow(sheetDoc: XMLDocument, rowNumber: number): Element {
  const sheetData = getDescendants(sheetDoc, 'sheetData')[0];
  if (!sheetData) throw new Error('The embedded worksheet has no sheetData element.');

  const existingRow = getElementChildren(sheetData)
    .find((row) => row.localName === 'row' && Number(row.getAttribute('r')) === rowNumber);
  if (existingRow) return existingRow;

  const row = sheetDoc.createElementNS(SPREADSHEET_NAMESPACE, 'row');
  row.setAttribute('r', String(rowNumber));
  const nextRow = getElementChildren(sheetData)
    .find((candidate) => candidate.localName === 'row' && Number(candidate.getAttribute('r')) > rowNumber);
  sheetData.insertBefore(row, nextRow ?? null);
  return row;
}

function getOrCreateCell(sheetDoc: XMLDocument, reference: string): Element {
  const location = parseCellReference(reference);
  const row = getOrCreateRow(sheetDoc, location.row);
  const existingCell = getElementChildren(row)
    .find((cell) => cell.localName === 'c' && cell.getAttribute('r') === reference);
  if (existingCell) return existingCell;

  const cell = sheetDoc.createElementNS(SPREADSHEET_NAMESPACE, 'c');
  cell.setAttribute('r', reference);
  const nextCell = getElementChildren(row).find((candidate) => {
    const candidateReference = candidate.getAttribute('r');
    return candidate.localName === 'c'
      && candidateReference !== null
      && parseCellReference(candidateReference).column > location.column;
  });
  row.insertBefore(cell, nextCell ?? null);
  return cell;
}

function clearCellContents(cell: Element): void {
  for (const child of getElementChildren(cell)) {
    if (child.localName === 'f' || child.localName === 'is' || child.localName === 'v') {
      child.parentNode?.removeChild(child);
    }
  }
}

function setCellValue(sheetDoc: XMLDocument, reference: string, value: string, kind: ChartDataValueKind): void {
  const cell = getOrCreateCell(sheetDoc, reference);
  clearCellContents(cell);

  if (value === '') {
    cell.removeAttribute('t');
    return;
  }

  if (kind === 'number') {
    cell.removeAttribute('t');
    const valueElement = sheetDoc.createElementNS(SPREADSHEET_NAMESPACE, 'v');
    valueElement.textContent = value;
    cell.appendChild(valueElement);
    return;
  }

  cell.setAttribute('t', 'inlineStr');
  const inlineString = sheetDoc.createElementNS(SPREADSHEET_NAMESPACE, 'is');
  const text = sheetDoc.createElementNS(SPREADSHEET_NAMESPACE, 't');
  if (value.trim() !== value) {
    text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  }
  text.textContent = value;
  inlineString.appendChild(text);
  cell.appendChild(inlineString);
}

function markWorkbookForRecalculation(workbookDoc: XMLDocument): void {
  let calculation = getDescendants(workbookDoc, 'calcPr')[0];
  if (!calculation) {
    calculation = workbookDoc.createElementNS(SPREADSHEET_NAMESPACE, 'calcPr');
    workbookDoc.documentElement.appendChild(calculation);
  }
  calculation.setAttribute('fullCalcOnLoad', '1');
  calculation.setAttribute('forceFullCalc', '1');
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

async function updateEmbeddedWorkbook(
  workbookBytes: Uint8Array,
  edits: Map<string, WorkbookCellEdit>
): Promise<Uint8Array> {
  const workbookBuffer = toArrayBuffer(workbookBytes);
  const workbookFiles = await extractZip(workbookBuffer);
  const sheetPaths = getSheetPaths(workbookFiles);
  const sheetDocs = new Map<string, XMLDocument>();
  const modifications = new Map<string, string>();

  for (const [formula, edit] of edits) {
    const range = parseCellRange(formula);
    if (!range) throw new Error(`Unsupported embedded workbook range: ${formula}`);
    const sheetPath = sheetPaths.get(range.sheetName);
    if (!sheetPath) throw new Error(`Could not find worksheet "${range.sheetName}".`);
    const sheetXml = workbookFiles.textFiles.get(sheetPath);
    if (!sheetXml) throw new Error(`Missing embedded worksheet XML part: ${sheetPath}`);
    const sheetDoc = sheetDocs.get(sheetPath) ?? parseXml(sheetXml, sheetPath);
    sheetDocs.set(sheetPath, sheetDoc);

    range.cellReferences.forEach((reference, index) => {
      setCellValue(sheetDoc, reference, edit.values[index] ?? '', edit.kind);
    });
  }

  for (const [sheetPath, sheetDoc] of sheetDocs) {
    modifications.set(sheetPath, serializeXml(sheetDoc));
  }

  const workbookXml = workbookFiles.textFiles.get('xl/workbook.xml');
  if (workbookXml) {
    const workbookDoc = parseXml(workbookXml, 'xl/workbook.xml');
    markWorkbookForRecalculation(workbookDoc);
    modifications.set('xl/workbook.xml', serializeXml(workbookDoc));
  }

  return new Uint8Array(await buildZip(workbookBuffer, modifications));
}

export async function updateChartData(
  buffer: ArrayBuffer,
  descriptor: ChartDataDescriptor,
  update: ChartDataUpdate
): Promise<ArrayBuffer> {
  if (!descriptor.grid.editable || !descriptor.categories || !descriptor.workbookPath) {
    throw new Error(descriptor.grid.reason || 'This chart data grid is read-only.');
  }
  if (update.categories.length !== descriptor.categories.values.length) {
    throw new Error('The category row count changed. Adding or deleting chart rows is not supported yet.');
  }
  if (update.series.length !== descriptor.series.length) {
    throw new Error('The series count changed. Adding or deleting chart series is not supported yet.');
  }

  const zip = await extractZip(buffer);
  const chartXml = zip.textFiles.get(descriptor.chartPath);
  const workbookBytes = zip.binaryFiles.get(descriptor.workbookPath);
  if (!chartXml || !workbookBytes) {
    throw new Error('The chart source parts are missing from the exported presentation.');
  }

  const chartDoc = parseXml(chartXml, descriptor.chartPath);
  const workbookEdits = new Map<string, WorkbookCellEdit>();
  addWorkbookEdit(workbookEdits, descriptor.categories, update.categories, descriptor.grid.categoryLabel);
  updateChartCaches(chartDoc, descriptor.categories.formula, update.categories);

  descriptor.series.forEach((series, seriesIndex) => {
    const seriesUpdate = update.series[seriesIndex];
    if (!seriesUpdate || seriesUpdate.values.length !== series.values.values.length) {
      throw new Error(`Series ${seriesIndex + 1} row count changed.`);
    }

    addWorkbookEdit(workbookEdits, series.values, seriesUpdate.values, `Series ${seriesIndex + 1}`);
    updateChartCaches(chartDoc, series.values.formula, seriesUpdate.values);

    const pointLabels = seriesUpdate.pointLabels;
    if (pointLabels === null && series.pointLabelBindings.length === 0) return;
    if (!pointLabels || pointLabels.length !== descriptor.categories?.values.length) {
      throw new Error(`Series ${seriesIndex + 1} point label row count changed.`);
    }

    for (const binding of series.pointLabelBindings) {
      const values = binding.indexes.map((index) => pointLabels[index] ?? '');
      addWorkbookEdit(workbookEdits, binding.source, values, `Series ${seriesIndex + 1} point label`);
      updateChartCaches(chartDoc, binding.source.formula, values);
    }
  });

  const workbook = await updateEmbeddedWorkbook(workbookBytes, workbookEdits);
  return buildZip(
    buffer,
    new Map([[descriptor.chartPath, serializeXml(chartDoc)]]),
    undefined,
    new Map([[descriptor.workbookPath, workbook]])
  );
}
