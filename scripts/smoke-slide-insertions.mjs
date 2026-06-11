import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { build } from 'esbuild';
import { extractZip } from 'pptx-svg';

const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

const inlineWasm = {
  name: 'inline-pptx-svg-wasm',
  setup(ctx) {
    ctx.onLoad({ filter: /pptx-renderer\.js$/ }, async (a) => ({
      contents: (await readFile(a.path, 'utf8')).replace(
        "const DEFAULT_WASM_URL = new URL('./main.wasm', import.meta.url).href;",
        'const DEFAULT_WASM_URL = undefined;'
      ),
      loader: 'js'
    }));
  }
};

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

const tmp = await mkdtemp(path.join(os.tmpdir(), 'native-powerpoint-insert-'));
try {
  const engineOut = path.join(tmp, 'engine.cjs');
  await build({
    entryPoints: [path.resolve('src/PresentationEngine.ts')],
    bundle: true,
    format: 'cjs',
    loader: { '.wasm': 'binary' },
    outfile: engineOut,
    platform: 'node',
    plugins: [inlineWasm],
    logLevel: 'silent'
  });
  const { PresentationEngine } = require(engineOut);

  const blank = toArrayBuffer(await readFile('test-results/native-powerpoint-fixtures/table-and-editable-chart.pptx'));
  const engine = await PresentationEngine.load(blank);
  const { slideIndex } = await engine.addSlide(engine.slideCount - 1);

  const rectIndex = engine.addShapeGeometry(slideIndex, 'rect');
  assert.equal(typeof rectIndex, 'number');

  const lineIndex = engine.addShapeGeometry(slideIndex, 'line');
  assert.notEqual(lineIndex, rectIndex);

  const png = Uint8Array.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d,
    0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
    0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49,
    0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82
  ]);
  const imageIndex = engine.addImage(slideIndex, png, 'image/png', 64, 48);
  assert.equal(typeof imageIndex, 'number');

  const tableIndex = await engine.addTable(slideIndex, 3, 2);
  assert.equal(typeof tableIndex, 'number');

  const chartIndex = await engine.addChart(slideIndex);
  assert.equal(typeof chartIndex, 'number');

  const textBoxIndex = engine.addTextBox(slideIndex);
  await engine.applyListStyle(slideIndex, textBoxIndex, 0, 'bullet');

  const exported = await engine.export();
  const zip = await extractZip(exported);
  const slidePath = `ppt/slides/slide${slideIndex + 1}.xml`;
  const slideXml = zip.textFiles.get(slidePath) ?? '';
  assert.match(slideXml, /graphicFrame/);
  assert.match(slideXml, /<a:buChar/);

  const rendered = engine.renderSlide(slideIndex).svg;
  assert.match(rendered, /data-ooxml-shape-idx=/);

  console.log('Slide insertion smoke passed: image, shapes, table, chart, and list formatting.');
} finally {
  await rm(tmp, { recursive: true, force: true });
}
