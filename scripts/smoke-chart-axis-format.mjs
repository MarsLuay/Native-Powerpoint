import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { build } from 'esbuild';

const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const tempDir = await mkdtemp(path.join(os.tmpdir(), 'native-powerpoint-chart-axis-'));
const bundlePath = path.join(tempDir, 'PresentationEngine.cjs');

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

function getAxisLabels(svg) {
  const normalizedSvg = svg.replace(/(\sfont-size="[^"]*")(?=[^<>]*\sfont-size=")/g, '');
  const document = new DOMParser().parseFromString(normalizedSvg, 'image/svg+xml');
  return { document, root: document.documentElement };
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

  const { PresentationEngine, formatChartAxisValue } = require(bundlePath);

  assert.equal(formatChartAxisValue(0.08, 'General', 0.04), '0.08');
  assert.equal(formatChartAxisValue(0.08, '0.00', 0.04), '0.08');
  assert.equal(formatChartAxisValue(0.125, '0.0%', 0.025), '12.5%');
  assert.equal(formatChartAxisValue(45292, 'm/d/yyyy', 1), '1/1/2024');
  assert.equal(formatChartAxisValue(45292, 'mmm-yy', 1), 'Jan-24');
  assert.equal(formatChartAxisValue(0, 'yyyy', 1, true), '1904');

  console.log('Synthetic chart-axis formats passed: decimal, percentage, date, and 1904 date system.');

  for (const filePath of process.argv.slice(2)) {
    const absolutePath = path.resolve(filePath);
    const engine = await PresentationEngine.load(toArrayBuffer(await readFile(absolutePath)));

    for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
      let root;
      try {
        ({ root } = getAxisLabels(engine.renderSlide(slideIndex).svg));
      } catch (error) {
        console.warn(`${path.relative(process.cwd(), absolutePath)} slide ${slideIndex + 1}: skipped strict XML probe (${error.message})`);
        continue;
      }

      engine.formatChartAxisLabels(root, slideIndex);

      const labels = Array.from(root.getElementsByTagName('text'))
        .filter((element) => element.getAttribute('data-native-powerpoint-axis-tick') === 'true')
        .map((element) => element.textContent);

      if (labels.length > 0) {
        console.log(`${path.relative(process.cwd(), absolutePath)} slide ${slideIndex + 1}: ${labels.join(', ')}`);
      }
    }
  }
} finally {
  await rm(tempDir, { recursive: true, force: true });
}
