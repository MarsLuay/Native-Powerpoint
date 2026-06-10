import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { build } from 'esbuild';
import { prepareNativePowerPointSmokeFixtures } from './fixtures/native-powerpoint-smoke-fixtures.mjs';

const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const configuredTableSampleArgument = process.argv[2] || process.env.NATIVE_POWERPOINT_TABLE_SAMPLE;
const configuredChartSampleArgument = process.argv[3] || process.env.NATIVE_POWERPOINT_CHART_SAMPLE;
const generatedFixtures = !configuredTableSampleArgument || !configuredChartSampleArgument
  ? await prepareNativePowerPointSmokeFixtures()
  : null;
const tableSampleArgument = configuredTableSampleArgument || generatedFixtures.tableSample;
const chartSampleArgument = configuredChartSampleArgument || generatedFixtures.chartSample;
const tempDir = await mkdtemp(path.join(os.tmpdir(), 'native-powerpoint-text-halos-'));
const engineBundlePath = path.join(tempDir, 'PresentationEngine.cjs');
const haloBundlePath = path.join(tempDir, 'TextHalo.cjs');
const tableSample = path.resolve(tableSampleArgument);
const chartSample = path.resolve(chartSampleArgument);

if (generatedFixtures) {
  console.log(`Text halo smoke using generated fixtures: ${path.relative(process.cwd(), tableSample)}`);
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

function getShapeGroups(svg, kind) {
  return Array.from(svg.getElementsByTagName('g'))
    .filter((group) => group.getAttribute('data-ooxml-shape-type') === kind);
}

function getTextElements(element) {
  return Array.from(element.getElementsByTagName('text'));
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
      entryPoints: [path.resolve('src/TextHalo.ts')],
      bundle: true,
      format: 'cjs',
      outfile: haloBundlePath,
      platform: 'node'
    })
  ]);

  const { PresentationEngine } = require(engineBundlePath);
  const { applyBackgroundAwareTextHalos, DEFAULT_TEXT_HALO_COLOR } = require(haloBundlePath);

  const tableEngine = await PresentationEngine.load(toArrayBuffer(await readFile(tableSample)));
  const tableSvg = parseSvg(tableEngine.renderSlide(3).svg);
  const tableGroups = getShapeGroups(tableSvg, 'table');
  tableGroups.forEach((group) => applyBackgroundAwareTextHalos(group, 'table'));
  const sampleLabels = tableGroups.flatMap(getTextElements)
    .filter((text) => text.textContent.includes('Samples'));
  assert.ok(sampleLabels.some((text) => text.getAttribute('data-native-powerpoint-halo-color') === 'rgb(197, 254, 220)'));

  const chartEngine = await PresentationEngine.load(toArrayBuffer(await readFile(chartSample)));
  const chartSvg = parseSvg(chartEngine.renderSlide(4).svg);
  const chartGroup = getShapeGroups(chartSvg, 'chart')[0];
  assert.ok(chartGroup);
  applyBackgroundAwareTextHalos(chartGroup, 'chart');
  assert.ok(
    getTextElements(chartGroup)
      .every((text) => text.getAttribute('data-native-powerpoint-halo-color') === 'rgb(255, 255, 255)')
  );

  const fallbackSvg = parseSvg('<svg xmlns="http://www.w3.org/2000/svg"><g><rect fill="none"/><text>Fallback</text></g></svg>');
  const fallbackGroup = fallbackSvg.getElementsByTagName('g')[0];
  assert.ok(fallbackGroup);
  applyBackgroundAwareTextHalos(fallbackGroup, 'table');
  assert.equal(
    fallbackGroup.getElementsByTagName('text')[0]?.getAttribute('data-native-powerpoint-halo-color'),
    DEFAULT_TEXT_HALO_COLOR
  );

  console.log('Text halo smoke passed: colored table, chart background, and fallback halos verified.');
} finally {
  await rm(tempDir, { recursive: true, force: true });
}
