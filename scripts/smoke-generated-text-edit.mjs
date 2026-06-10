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
const configuredSampleArgument = process.argv[2] || process.env.NATIVE_POWERPOINT_GENERATED_TEXT_SAMPLE;
const generatedFixtures = configuredSampleArgument ? null : await prepareNativePowerPointSmokeFixtures();
const sampleArgument = configuredSampleArgument || generatedFixtures.tableSample;
const tempDir = await mkdtemp(path.join(os.tmpdir(), 'native-powerpoint-generated-text-'));
const engineBundlePath = path.join(tempDir, 'PresentationEngine.cjs');
const packageBundlePath = path.join(tempDir, 'PowerPointPackage.cjs');

if (generatedFixtures) {
  console.log(`Generated text edit smoke using generated fixtures: ${path.relative(process.cwd(), sampleArgument)}`);
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

function getShapeGroup(svg, kind) {
  return Array.from(svg.getElementsByTagName('g'))
    .find((group) => group.getAttribute('data-ooxml-shape-type') === kind);
}

function getShapeIndex(group) {
  const shapeIndex = Number(group?.getAttribute('data-ooxml-shape-idx'));
  assert.ok(Number.isFinite(shapeIndex), `Could not find ${group?.getAttribute('data-ooxml-shape-type') ?? 'shape'} index.`);
  return shapeIndex;
}

function getTexts(group) {
  assert.ok(group, 'Expected a generated table or chart group.');
  return Array.from(group.getElementsByTagName('text')).map((text) => text.textContent.trim());
}

try {
  await Promise.all([
    build({
      entryPoints: [path.resolve('src/PresentationEngine.ts')],
      bundle: true,
      format: 'cjs',
      loader: { '.wasm': 'binary' },
      outfile: engineBundlePath,
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

  const { PresentationEngine } = require(engineBundlePath);
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents
  } = require(packageBundlePath);
  const samplePath = path.resolve(sampleArgument);
  const sampleBuffer = toArrayBuffer(await readFile(samplePath));
  const engine = await PresentationEngine.load(sampleBuffer);

  const tableGroup = getShapeGroup(parseSvg(engine.renderSlide(3).svg), 'table');
  const tableShapeIndex = getShapeIndex(tableGroup);
  assert.deepEqual(getTexts(tableGroup), ['Samples', 'Value', '1g', '1.04']);
  await engine.updateGeneratedText(
    3,
    tableShapeIndex,
    { kind: 'table', labelIndex: 0, occurrence: 0, previousText: 'Samples' },
    'Specimens'
  );
  assert.deepEqual(
    getTexts(getShapeGroup(parseSvg(engine.renderSlide(3).svg), 'table')),
    ['Specimens', 'Value', '1g', '1.04']
  );

  const chartGroup = getShapeGroup(parseSvg(engine.renderSlide(4).svg), 'chart');
  const chartShapeIndex = getShapeIndex(chartGroup);
  assert.ok(getTexts(chartGroup).includes('1g'));
  assert.ok(getTexts(chartGroup).includes('Samples'));
  assert.equal(
    engine.canUpdateGeneratedText(
      4,
      chartShapeIndex,
      { kind: 'chart', labelIndex: 2, occurrence: 0, previousText: '1' }
    ),
    false
  );
  assert.equal(
    engine.canUpdateGeneratedText(
      4,
      chartShapeIndex,
      { kind: 'chart', labelIndex: 0, occurrence: 0, previousText: '1g' }
    ),
    true
  );

  await engine.updateGeneratedText(
    4,
    chartShapeIndex,
    { kind: 'chart', labelIndex: 0, occurrence: 0, previousText: '1g' },
    '1.5g'
  );
  await engine.updateGeneratedText(
    4,
    chartShapeIndex,
    { kind: 'chart', labelIndex: 8, occurrence: 0, previousText: 'Samples' },
    'Specimens'
  );
  const updatedChartTexts = getTexts(getShapeGroup(parseSvg(engine.renderSlide(4).svg), 'chart'));
  assert.ok(updatedChartTexts.includes('1.5g'));
  assert.ok(updatedChartTexts.includes('Specimens'));
  assert.ok(!updatedChartTexts.includes('1g'));
  assert.ok(!updatedChartTexts.includes('Samples'));

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
  const tableSlideXml = await pptx.file('ppt/slides/slide4.xml').async('string');
  const chartXml = await pptx.file('ppt/charts/chart1.xml').async('string');
  assert.match(tableSlideXml, /<a:t>Specimens<\/a:t>/);
  assert.match(chartXml, /<c:v>1\.5g<\/c:v>/);
  assert.match(chartXml, /<c:v>Specimens<\/c:v>/);

  const workbook = await JSZip.loadAsync(
    await pptx.file('ppt/embeddings/Microsoft_Excel_Worksheet.xlsx').async('nodebuffer')
  );
  const worksheetXml = await workbook.file('xl/worksheets/sheet1.xml').async('string');
  const workbookXml = await workbook.file('xl/workbook.xml').async('string');
  assert.match(worksheetXml, /<c r="A2" t="inlineStr"><is><t>1\.5g<\/t><\/is><\/c>/);
  assert.match(worksheetXml, /<c r="B1" t="inlineStr"><is><t>Specimens<\/t><\/is><\/c>/);
  assert.match(workbookXml, /fullCalcOnLoad="1"/);
  assert.match(workbookXml, /forceFullCalc="1"/);

  console.log('Generated text edit smoke passed: table cells, chart labels, and embedded workbook text stayed in sync.');
} finally {
  await rm(tempDir, { recursive: true, force: true });
}
