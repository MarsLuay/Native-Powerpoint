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
const configuredSampleArgument = process.argv[2] || process.env.NATIVE_POWERPOINT_CHART_DATA_SAMPLE;
const configuredExternalSampleArgument = process.argv[3] || process.env.NATIVE_POWERPOINT_EXTERNAL_CHART_SAMPLE;
const generatedFixtures = !configuredSampleArgument || !configuredExternalSampleArgument
  ? await prepareNativePowerPointSmokeFixtures()
  : null;
const sampleArgument = configuredSampleArgument || generatedFixtures.chartDataSample;
const externalSampleArgument = configuredExternalSampleArgument || generatedFixtures.externalChartSample;
const tempDir = await mkdtemp(path.join(os.tmpdir(), 'native-powerpoint-chart-data-'));
const bundlePath = path.join(tempDir, 'PresentationEngine.cjs');
const packageBundlePath = path.join(tempDir, 'PowerPointPackage.cjs');

if (generatedFixtures) {
  console.log(`Chart data edit smoke using generated fixtures: ${path.relative(process.cwd(), sampleArgument)}`);
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

function cloneGridUpdate(grid) {
  return {
    categories: [...grid.categories],
    series: grid.series.map((series) => ({
      values: [...series.values],
      pointLabels: series.pointLabels === null ? null : [...series.pointLabels]
    }))
  };
}

function findEditableChart(engine) {
  for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
    for (let shapeIndex = 0; shapeIndex < 80; shapeIndex++) {
      const grid = engine.getChartDataGrid(slideIndex, shapeIndex);
      if (grid?.editable) return { grid, shapeIndex, slideIndex };
    }
  }
  throw new Error('Could not find an editable embedded-workbook chart in the sample.');
}

try {
  await build({
    entryPoints: [path.resolve('src/PresentationEngine.ts')],
    bundle: true,
    format: 'cjs',
    loader: { '.wasm': 'binary' },
    outfile: bundlePath,
    platform: 'node',
    plugins: [inlinePptxSvgWasmPlugin]
  });
  await build({
    entryPoints: [path.resolve('src/PowerPointPackage.ts')],
    bundle: true,
    format: 'cjs',
    outfile: packageBundlePath,
    platform: 'node'
  });

  const { PresentationEngine } = require(bundlePath);
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents
  } = require(packageBundlePath);
  const samplePath = path.resolve(sampleArgument);
  const sampleBuffer = toArrayBuffer(await readFile(samplePath));
  const engine = await PresentationEngine.load(sampleBuffer);
  const chart = findEditableChart(engine);
  const update = cloneGridUpdate(chart.grid);

  assert.equal(chart.grid.categories[0], '1g');
  assert.equal(chart.grid.series[0]?.values[0], '1.04');

  update.categories[0] = '1.5g';
  update.series[0].values[0] = '1.11';
  await engine.updateChartData(chart.slideIndex, chart.shapeIndex, update);

  const reloadedGrid = engine.getChartDataGrid(chart.slideIndex, chart.shapeIndex);
  assert.equal(reloadedGrid?.categories[0], '1.5g');
  assert.equal(reloadedGrid?.series[0]?.values[0], '1.11');

  const output = await engine.export();
  await PresentationEngine.validateRoundTrip(output, engine.slideCount);
  assert.equal(
    validatePowerPointExport(
      inspectPowerPointPackage(sampleBuffer),
      inspectPowerPointPackage(output),
      engine.slideCount
    ).ok,
    true
  );
  assert.equal((await validatePowerPointExportContents(sampleBuffer, output)).ok, true);

  const pptx = await JSZip.loadAsync(Buffer.from(output));
  const chartXml = await pptx.file('ppt/charts/chart1.xml').async('string');
  assert.match(chartXml, /<c:v>1\.5g<\/c:v>/);
  assert.match(chartXml, /<c:v>1\.11<\/c:v>/);

  const workbook = await JSZip.loadAsync(
    await pptx.file('ppt/embeddings/Microsoft_Excel_Worksheet.xlsx').async('nodebuffer')
  );
  const worksheetXml = await workbook.file('xl/worksheets/sheet1.xml').async('string');
  const workbookXml = await workbook.file('xl/workbook.xml').async('string');
  assert.match(worksheetXml, /<c r="A2" t="inlineStr"><is><t>1\.5g<\/t><\/is><\/c>/);
  assert.match(worksheetXml, /<c r="B2"><v>1\.11<\/v><\/c>/);
  assert.match(workbookXml, /fullCalcOnLoad="1"/);
  assert.match(workbookXml, /forceFullCalc="1"/);

  const externalPath = path.resolve(externalSampleArgument);
  const externalEngine = await PresentationEngine.load(toArrayBuffer(await readFile(externalPath)));
  let externalReason = '';
  for (let slideIndex = 0; slideIndex < externalEngine.slideCount && !externalReason; slideIndex++) {
    for (let shapeIndex = 0; shapeIndex < 80; shapeIndex++) {
      const grid = externalEngine.getChartDataGrid(slideIndex, shapeIndex);
      if (grid && !grid.editable && grid.reason.includes('external workbook link')) {
        externalReason = grid.reason;
        break;
      }
    }
  }
  assert.match(externalReason, /external workbook link/i);

  console.log('Chart data edit smoke passed: chart caches and embedded workbook cells changed together.');
} finally {
  await rm(tempDir, { recursive: true, force: true });
}
