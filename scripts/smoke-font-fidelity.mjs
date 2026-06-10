import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { build } from 'esbuild';

const require = createRequire(import.meta.url);
const { DOMParser } = require('@xmldom/xmldom');
const tempDir = await mkdtemp(path.join(os.tmpdir(), 'native-powerpoint-font-fidelity-'));
const bundlePath = path.join(tempDir, 'FontFidelity.cjs');

function parseSvg(svg) {
  return new DOMParser().parseFromString(svg, 'image/svg+xml').documentElement;
}

try {
  await build({
    entryPoints: [path.resolve('src/FontFidelity.ts')],
    bundle: true,
    format: 'cjs',
    outfile: bundlePath,
    platform: 'node'
  });

  const { FontFidelity, splitCssFontFamilies } = require(bundlePath);
  const measuredFonts = [];
  const availableFonts = new Set(['Arial', 'Georgia', 'Courier New', 'Installed Display']);
  const fidelity = new FontFidelity({
    isFontAvailable: (fontFamily) => availableFonts.has(fontFamily),
    measureText: (text, fontFamily, fontSizePx) => {
      measuredFonts.push(fontFamily);
      return text.length * fontSizePx * (fontFamily === 'Arial' ? 0.55 : 0.6);
    }
  });

  assert.deepEqual(
    splitCssFontFamilies('"Installed Display", \'Fallback Font\', sans-serif'),
    ['Installed Display', 'Fallback Font', 'sans-serif']
  );
  assert.ok(Math.abs(fidelity.measureText('PowerPoint', 'Missing Sans', 20) - 110) < 0.0001);
  assert.equal(measuredFonts.at(-1), 'Arial');

  const svg = parseSvg(`
    <svg xmlns="http://www.w3.org/2000/svg">
      <text font-family="Missing Sans">Fallback</text>
      <text font-family="Installed Display">Installed</text>
      <text style="font-family: Missing Serif; font-size: 18px">Serif</text>
      <text font-family="Calibri">Office</text>
    </svg>
  `);
  const substitutions = fidelity.applySvgSubstitutions(svg);
  const text = Array.from(svg.getElementsByTagName('text'));

  assert.deepEqual(substitutions, [
    { requested: 'Calibri', substitute: 'Arial' },
    { requested: 'Missing Sans', substitute: 'Arial' },
    { requested: 'Missing Serif', substitute: 'Georgia' }
  ]);
  assert.match(text[0].getAttribute('font-family'), /^"Missing Sans", "Arial"/);
  assert.equal(text[0].getAttribute('data-native-powerpoint-font-substitution'), 'Arial');
  assert.equal(text[1].getAttribute('font-family'), 'Installed Display');
  assert.match(text[2].getAttribute('style'), /font-family: "Missing Serif", "Georgia"/);
  assert.match(text[3].getAttribute('font-family'), /^"Calibri", "Arial"/);

  console.log('Font fidelity smoke passed: detection, fallback measurement, SVG stacks, and substitution reporting verified.');
} finally {
  await rm(tempDir, { recursive: true, force: true });
}
