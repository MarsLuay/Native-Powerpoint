// Regenerate the pure-JavaScript PPTX engine fallback (src/vendor/pptx-js-engine.mjs).
//
// Background
// ----------
// The default PPTX renderer (`pptx-svg`) ships a MoonBit module compiled to the
// `wasm-gc` target, which requires WebAssembly GC (Chromium 119+ / Electron 28+).
// Obsidian installs older than 1.5.8 bundle a Chromium that cannot run it. MoonBit
// can also compile the *same source* to a pure-JS backend that runs everywhere, so
// we vendor that build as a lazy-loaded fallback.
//
// The raw MoonBit JS output is an ES module with module-scoped state and a single
// trailing `export { internal as public, ... }`. Module state is a singleton, so
// two open presentations would clobber each other. This script rewraps the output
// into a `createPptxJsEngine()` factory that returns a fresh, isolated set of the
// exported functions per call.
//
// Prerequisites (one-time, only when regenerating):
//   1. Install MoonBit:  curl -fsSL https://cli.moonbitlang.com/install/unix.sh | bash
//   2. Clone pptx-svg at the version in node_modules/pptx-svg/package.json.
//   3. Add a `"js"` block (mirroring the `"wasm-gc"` exports) to src/main/moon.pkg.
//   4. Build:  moon build --target js --release
//
// Usage:
//   node scripts/build-pptx-js-engine.mjs <path-to-moon-js-build/main.js>
//   PPTX_SVG_JS_BUILD=<path> node scripts/build-pptx-js-engine.mjs

import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, '..');

const inputPath = process.argv[2] || process.env.PPTX_SVG_JS_BUILD;
if (!inputPath) {
  console.error(
    'Missing input. Pass the MoonBit JS build path, e.g.\n' +
    '  node scripts/build-pptx-js-engine.mjs /path/to/pptx-svg/_build/js/release/build/main/main.js'
  );
  process.exit(1);
}

const outputPath = path.join(projectRoot, 'src/vendor/pptx-js-engine.mjs');

const raw = await readFile(inputPath, 'utf8');

// MoonBit emits exactly one trailing `export { a as b, ... }` statement. Pull it
// out and turn it into a `return { b: a, ... }` so each factory call yields its
// own bound copy of the exported functions.
const exportMatch = raw.match(/export\s*\{([\s\S]*?)\}\s*;?\s*$/);
if (!exportMatch) {
  console.error('Could not find a trailing `export { ... }` statement in the MoonBit output.');
  process.exit(1);
}

const body = raw.slice(0, exportMatch.index);
const pairs = exportMatch[1]
  .split(',')
  .map((entry) => entry.trim())
  .filter(Boolean)
  .map((entry) => {
    const m = entry.match(/^(\S+)\s+as\s+(\S+)$/);
    if (!m) {
      console.error(`Unexpected export entry: "${entry}"`);
      process.exit(1);
    }
    return { internal: m[1], external: m[2] };
  });

const returnObject = pairs
  .map(({ internal, external }) => `    ${external}: ${internal},`)
  .join('\n');

const banner =
`/*
 * GENERATED FILE — DO NOT EDIT BY HAND.
 *
 * Pure-JavaScript PPTX engine, produced from the MoonBit \`js\` backend of
 * pptx-svg and rewrapped into a per-instance factory by
 * scripts/build-pptx-js-engine.mjs. Used as a lazy-loaded fallback when the
 * runtime lacks WebAssembly GC (Obsidian installer < 1.5.8 / Chromium < 119).
 *
 * The engine resolves its host calls through the global \`pptx_ffi\`, which the
 * renderer sets immediately before each (synchronous) call. Each
 * createPptxJsEngine() call returns an isolated set of exports with its own
 * module state, so multiple presentations can be open at once.
 */
`;

const output =
`${banner}
export function createPptxJsEngine() {
${indent(body.trimEnd(), 2)}

  return {
${returnObject}
  };
}
`;

await writeFile(outputPath, output, 'utf8');
console.log(`Wrote ${path.relative(projectRoot, outputPath)} (${output.length} bytes, ${pairs.length} exports).`);

function indent(text, spaces) {
  const pad = ' '.repeat(spaces);
  return text
    .split('\n')
    .map((line) => (line.length ? pad + line : line))
    .join('\n');
}
