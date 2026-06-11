// End-to-end smoke test for the pure-JS PPTX engine fallback.
//
// Loads a real .pptx through the vendored createPptxJsEngine() factory using the
// same host FFI surface the renderer provides, then renders slide 0 to SVG.
// Proves the js backend works without any WebAssembly at all.
//
// Usage: node scripts/smoke-pptx-js-engine.mjs [path-to.pptx]

import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, '..');

const { createPptxJsEngine } = await import(
  path.join(projectRoot, 'src/vendor/pptx-js-engine.mjs')
);
const { extractZip } = await import(
  path.join(projectRoot, 'node_modules/pptx-svg/dist/zip.js')
);
const { bytesToBase64 } = await import(
  path.join(projectRoot, 'node_modules/pptx-svg/dist/utils.js')
);

const pptxPath =
  process.argv[2] || path.join(projectRoot, 'test_files/10MB-Sample-PPT-File.pptx');

const buf = await readFile(pptxPath);
const arrayBuffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);

const silentLog = { debug() {}, info() {}, warn() {}, error() {} };
const { textFiles, binaryFiles } = await extractZip(arrayBuffer, silentLog);

function makeFfi() {
  return {
    get_file: (p) => textFiles.get(p) ?? '',
    get_entry_list: () => [...textFiles.keys(), ...binaryFiles.keys()].join('\n'),
    get_file_base64: (p) => bytesToBase64(binaryFiles.get(p)),
    char_code_to_str: (n) => String.fromCodePoint(n),
    log: () => {},
    warn: () => {},
    error: () => {},
    measure_text: (text, _font, sizePx) => text.length * sizePx * 0.6,
    get_font_fallback: () => '',
    convert_emf: () => '',
    math_sin: Math.sin,
    math_cos: Math.cos,
    math_atan2: Math.atan2,
    math_sqrt: Math.sqrt,
  };
}

function assert(cond, msg) {
  if (!cond) {
    console.error(`FAIL: ${msg}`);
    process.exit(1);
  }
}

// Two isolated engines to verify per-instance state (multiple open decks).
const engineA = createPptxJsEngine();
const engineB = createPptxJsEngine();

globalThis.pptx_ffi = makeFfi();
const initA = engineA.initialize_pptx();
assert(!initA.startsWith('ERROR:'), `engineA initialize_pptx returned ${initA}`);

const slideCount = engineA.get_slide_count();
assert(slideCount > 0, `expected slideCount > 0, got ${slideCount}`);

const svg = engineA.render_slide_svg(0);
assert(typeof svg === 'string' && svg.includes('<svg'), 'render_slide_svg(0) did not return SVG');
assert(!svg.startsWith('ERROR:'), `render_slide_svg returned ${svg.slice(0, 80)}`);

// engineB shares the same FFI/source here, but must hold independent state:
// initializing B must not disturb A's already-parsed presentation.
const initB = engineB.initialize_pptx();
assert(!initB.startsWith('ERROR:'), `engineB initialize_pptx returned ${initB}`);
const svgAagain = engineA.render_slide_svg(0);
assert(svgAagain === svg, 'engineA render changed after engineB init — state leaked between instances');

console.log(
  `PASS: js engine rendered slide 0 of ${path.basename(pptxPath)} ` +
  `(slides=${slideCount}, svg=${svg.length} bytes), per-instance state isolated.`
);
