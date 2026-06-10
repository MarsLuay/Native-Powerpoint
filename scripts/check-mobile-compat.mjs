// Guards Obsidian Mobile (iOS/Android) compatibility by scanning the built
// bundle for hard CommonJS dependencies that do not exist in the mobile WebView.
//
// Obsidian provides require("obsidian") on every platform, but Node builtins
// (fs, path, stream, ...) and require("electron") only exist on desktop. Any
// static require of those in main.js would crash the plugin on load on mobile.
import assert from 'node:assert/strict';
import { builtinModules } from 'node:module';
import { readFile } from 'node:fs/promises';
import path from 'node:path';

const bundlePath = path.resolve('main.js');
const source = await readFile(bundlePath, 'utf8');

const allowed = new Set(['obsidian']);
const nodeBuiltins = new Set([
  ...builtinModules,
  ...builtinModules.map((name) => `node:${name}`),
]);

const found = new Set();
for (const match of source.matchAll(/require\(\s*["']([^"']+)["']\s*\)/g)) {
  found.add(match[1]);
}

const offenders = [...found].filter(
  (name) => !allowed.has(name) && (name === 'electron' || nodeBuiltins.has(name)),
);

assert.deepEqual(
  offenders,
  [],
  `main.js statically requires platform-only modules that break Obsidian Mobile: ${offenders.join(', ')}. ` +
    'Alias them to a browser shim (see esbuild.config.mjs) or guard them behind window.require.',
);

console.log(`Mobile compat check passed: ${bundlePath} requires only [${[...found].join(', ')}].`);
