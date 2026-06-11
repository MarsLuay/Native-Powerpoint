import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { build } from 'esbuild';

const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

const samplePath = path.resolve(process.argv[2] || 'test_files/10MB-Sample-PPT-File.pptx');
const outDir = path.resolve('scripts/visual-output/real-svg');
const tempDir = await mkdtemp(path.join(os.tmpdir(), 'np-dump-'));
const engineBundlePath = path.join(tempDir, 'PresentationEngine.cjs');

const inlineWasm = {
  name: 'inline-pptx-svg-wasm',
  setup(ctx) {
    ctx.onLoad({ filter: /pptx-renderer\.js$/ }, async (args) => {
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

await build({
  entryPoints: [path.resolve('src/PresentationEngine.ts')],
  bundle: true,
  format: 'cjs',
  loader: { '.wasm': 'binary' },
  outfile: engineBundlePath,
  platform: 'node',
  plugins: [inlineWasm]
});

const { PresentationEngine } = require(engineBundlePath);
const buffer = toArrayBuffer(await readFile(samplePath));
const engine = await PresentationEngine.load(buffer);
await mkdir(outDir, { recursive: true });
const serializer = new XMLSerializer();

console.log(`Deck: ${samplePath}`);
console.log(`Slides: ${engine.slideCount}`);

for (let i = 0; i < engine.slideCount; i++) {
  const rendered = engine.renderSlide(i);
  const file = path.join(outDir, `slide-${String(i + 1).padStart(2, '0')}.svg`);
  await writeFile(file, rendered.svg, 'utf8');
}
console.log(`Wrote SVGs to ${outDir}`);
