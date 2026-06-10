import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { build } from 'esbuild';
import { prepareNativePowerPointSmokeFixtures } from './fixtures/native-powerpoint-smoke-fixtures.mjs';

const require = createRequire(import.meta.url);
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const configuredTableSampleArgument = process.argv[2] || process.env.NATIVE_POWERPOINT_TABLE_SAMPLE;
const configuredChartSampleArgument = process.argv[3] || process.env.NATIVE_POWERPOINT_CHART_SAMPLE;
const generatedFixtures = !configuredTableSampleArgument || !configuredChartSampleArgument
  ? await prepareNativePowerPointSmokeFixtures()
  : null;
const tableSampleArgument = configuredTableSampleArgument || generatedFixtures.tableSample;
const chartSampleArgument = configuredChartSampleArgument || generatedFixtures.chartSample;
const tempDir = await mkdtemp(path.join(os.tmpdir(), 'native-powerpoint-object-clipboard-'));
const bundlePath = path.join(tempDir, 'PresentationEngine.cjs');
const packageBundlePath = path.join(tempDir, 'PowerPointPackage.cjs');
const tableSample = path.resolve(tableSampleArgument);
const chartSample = path.resolve(chartSampleArgument);

if (generatedFixtures) {
  console.log(`Object clipboard smoke using generated fixtures: ${path.relative(process.cwd(), tableSample)}`);
}

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

const inlinePptxSvgWasmPlugin = {
  name: 'inline-pptx-svg-wasm',
  setup(buildContext) {
    buildContext.onLoad({ filter: /pptx-renderer\.js$/ }, async (args) => {
      const source = await readFile(args.path, 'utf8');
      return {
        contents: source.replace(
          "const DEFAULT_WASM_URL = new URL('./main.wasm', import.meta.url).href;",
          'const DEFAULT_WASM_URL = undefined;'
        ),
        loader: 'js'
      };
    });
  }
};

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function parseSvg(svg) {
  const normalizedSvg = svg.replace(/(\sfont-size="[^"]*")(?=[^<>]*\sfont-size=")/g, '');
  return new DOMParser().parseFromString(normalizedSvg, 'image/svg+xml').documentElement;
}

function findShape(engine, type) {
  for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
    const svg = parseSvg(engine.renderSlide(slideIndex).svg);
    const shape = Array.from(svg.getElementsByTagName('g'))
      .find((group) => group.getAttribute('data-ooxml-shape-type') === type);
    if (shape) {
      return {
        slideIndex,
        shapeIndex: Number(shape.getAttribute('data-ooxml-shape-idx'))
      };
    }
  }
  throw new Error(`Could not find a ${type} shape in the sample.`);
}

function findEditableChart(engine) {
  for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
    for (let shapeIndex = 0; shapeIndex < 80; shapeIndex++) {
      const grid = engine.getChartDataGrid(slideIndex, shapeIndex);
      if (grid?.editable) return { grid, shapeIndex, slideIndex };
    }
  }
  throw new Error('Could not find an editable chart in the sample.');
}

function cloneGridUpdate(grid) {
  return {
    categories: [...grid.categories],
    series: grid.series.map((series) => ({
      values: [...series.values],
      pointLabels: series.pointLabels === null ? null : [...series.pointLabels]
    }))
  };
}

try {
  await Promise.all([
    build({
      entryPoints: [path.resolve('src/PresentationEngine.ts')],
      bundle: true,
      format: 'cjs',
      loader: { '.wasm': 'binary' },
      outfile: bundlePath,
      platform: 'node',
      plugins: [inlinePptxSvgWasmPlugin]
    }),
    build({
      entryPoints: [path.resolve('src/PowerPointPackage.ts')],
      bundle: true,
      format: 'cjs',
      outfile: packageBundlePath,
      platform: 'node'
    })
  ]);

  const { PresentationEngine } = require(bundlePath);
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents
  } = require(packageBundlePath);

  const tableBuffer = toArrayBuffer(await readFile(tableSample));
  const tableEngine = await PresentationEngine.load(tableBuffer);
  const table = findShape(tableEngine, 'table');
  const tableDestinationSlide = (table.slideIndex + 1) % tableEngine.slideCount;
  const originalTableZip = await JSZip.loadAsync(Buffer.from(tableBuffer));
  const originalTableSlideXml = await originalTableZip.file(`ppt/slides/slide${tableDestinationSlide + 1}.xml`).async('string');
  const originalTableCount = (originalTableSlideXml.match(/<a:tbl\b/g) ?? []).length;
  const pastedTableIndex = await tableEngine.pasteShape(
    await tableEngine.copyShape(table.slideIndex, table.shapeIndex),
    tableDestinationSlide
  );
  assert.ok(Number.isFinite(pastedTableIndex));
  const pastedTableBuffer = await tableEngine.export();
  await PresentationEngine.validateRoundTrip(pastedTableBuffer, tableEngine.slideCount);
  assert.equal((await validatePowerPointExportContents(tableBuffer, pastedTableBuffer)).ok, true);
  const tableZip = await JSZip.loadAsync(Buffer.from(pastedTableBuffer));
  const tableSlideXml = await tableZip.file(`ppt/slides/slide${tableDestinationSlide + 1}.xml`).async('string');
  assert.equal((tableSlideXml.match(/<a:tbl\b/g) ?? []).length, originalTableCount + 1);

  const chartBuffer = toArrayBuffer(await readFile(chartSample));
  const chartEngine = await PresentationEngine.load(chartBuffer);
  const chart = findEditableChart(chartEngine);
  const chartDestinationSlide = (chart.slideIndex + 1) % chartEngine.slideCount;
  const originalCategory = chart.grid.categories[0];
  const pastedChartIndex = await chartEngine.pasteShape(
    await chartEngine.copyShape(chart.slideIndex, chart.shapeIndex),
    chartDestinationSlide
  );
  const pastedChart = chartEngine.getChartDataGrid(chartDestinationSlide, pastedChartIndex);
  assert.equal(pastedChart?.editable, true);

  const update = cloneGridUpdate(pastedChart);
  update.categories[0] = `${originalCategory} copy`;
  await chartEngine.updateChartData(chartDestinationSlide, pastedChartIndex, update);
  assert.equal(chartEngine.getChartDataGrid(chart.slideIndex, chart.shapeIndex)?.categories[0], originalCategory);
  assert.equal(chartEngine.getChartDataGrid(chartDestinationSlide, pastedChartIndex)?.categories[0], `${originalCategory} copy`);

  const pastedChartBuffer = await chartEngine.export();
  await PresentationEngine.validateRoundTrip(pastedChartBuffer, chartEngine.slideCount);
  assert.equal(
    validatePowerPointExport(
      inspectPowerPointPackage(chartBuffer),
      inspectPowerPointPackage(pastedChartBuffer),
      chartEngine.slideCount
    ).ok,
    true
  );
  assert.equal((await validatePowerPointExportContents(chartBuffer, pastedChartBuffer)).ok, true);
  const chartZip = await JSZip.loadAsync(Buffer.from(pastedChartBuffer));
  assert.ok(Object.keys(chartZip.files).filter((file) => /^ppt\/charts\/chart\d+\.xml$/.test(file)).length >= 2);
  assert.ok(Object.keys(chartZip.files).filter((file) => /^ppt\/embeddings\/.*\.xlsx$/i.test(file)).length >= 2);

  console.log('Object clipboard smoke passed: tables duplicate intact and charts receive independent caches and workbooks.');
} finally {
  await rm(tempDir, { recursive: true, force: true });
}
