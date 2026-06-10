import assert from 'node:assert/strict';
import { existsSync, statSync } from 'node:fs';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';

const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

// Mirror the plugin's zoom envelope (MIN_ZOOM..MAX_ZOOM) plus the default.
const ZOOM_LEVELS = [0.25, 0.5, 1, 2, 3];
// A representative canvas pane viewport; padding matches .native-powerpoint-canvas-pane (28px).
const PANE_WIDTH = 1000;
const PANE_HEIGHT = 620;
const PANE_PADDING = 28;

const sampleArgument = process.argv[2] || process.env.NATIVE_POWERPOINT_VISUAL_SAMPLE;
if (!sampleArgument) {
  throw new Error('Usage: node scripts/visual-powerpoint-zoom.mjs <file.pptx>');
}
const samplePath = path.resolve(sampleArgument);
const outputDir = path.resolve('scripts/visual-output');
const htmlPath = path.join(outputDir, 'powerpoint-zoom.html');
const zoomShotPath = path.join(outputDir, 'powerpoint-zoom.png');
const slidesShotPath = path.join(outputDir, 'powerpoint-slides.png');

function findChrome() {
  const candidates = [
    process.env.CHROME_PATH,
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    '/usr/bin/google-chrome',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser'
  ].filter(Boolean);
  const chrome = candidates.find((candidate) => existsSync(candidate));
  if (!chrome) throw new Error('Chrome/Edge not found. Set CHROME_PATH.');
  return chrome;
}

function fileHasBytes(filePath) {
  try { return statSync(filePath).size > 0; } catch { return false; }
}

function runChrome(chromePath, args, options = {}) {
  const timeoutMs = options.timeoutMs ?? 20000;
  return new Promise((resolve, reject) => {
    const child = spawn(chromePath, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill('SIGTERM');
      setTimeout(() => { if (child.exitCode === null) child.kill('SIGKILL'); }, 2000).unref();
    }, timeoutMs);
    child.stdout.on('data', (c) => { stdout += c; });
    child.stderr.on('data', (c) => { stderr += c; });
    child.on('error', (error) => { clearTimeout(timer); reject(error); });
    child.on('close', (code) => {
      clearTimeout(timer);
      if (code === 0) return resolve({ stdout, stderr });
      if (timedOut && options.resolveOnTimeout?.({ stdout, stderr }) === true) return resolve({ stdout, stderr });
      reject(new Error(timedOut ? `Chrome timed out: ${stderr || stdout}` : `Chrome exited ${code}: ${stderr || stdout}`));
    });
  });
}

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

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

function getIntrinsicSize(svg) {
  const document = new DOMParser().parseFromString(svg, 'image/svg+xml');
  const root = document.documentElement;
  const viewBox = (root.getAttribute('viewBox') || '').split(/[\s,]+/).map(Number);
  if (viewBox.length === 4 && viewBox[2] > 0 && viewBox[3] > 0) {
    return { width: viewBox[2], height: viewBox[3] };
  }
  const width = Number.parseFloat(root.getAttribute('width') || '');
  const height = Number.parseFloat(root.getAttribute('height') || '');
  if (width > 0 && height > 0) return { width, height };
  throw new Error('Slide SVG has no intrinsic size.');
}

// Reproduces NativePowerPointView.updateSlideScale()
function computeScaledSize(intrinsic, zoom) {
  const availableWidth = Math.max(1, PANE_WIDTH - PANE_PADDING * 2);
  const availableHeight = Math.max(1, PANE_HEIGHT - PANE_PADDING * 2);
  const fitScale = Math.min(1, availableWidth / intrinsic.width, availableHeight / intrinsic.height);
  const scale = Math.max(0.05, fitScale * zoom);
  return {
    fitScale,
    width: Math.max(1, Math.floor(intrinsic.width * scale)),
    height: Math.max(1, Math.floor(intrinsic.height * scale))
  };
}

function sizeSvg(svg, width, height) {
  return svg
    .replace(/(<svg\b[^>]*?)\swidth="[^"]*"/, '$1')
    .replace(/(<svg\b[^>]*?)\sheight="[^"]*"/, '$1')
    .replace(/<svg\b/, `<svg width="${width}" height="${height}"`);
}

const tmp = await mkdtemp(path.join(os.tmpdir(), 'native-powerpoint-visual-'));
try {
  const engineOut = path.join(tmp, 'engine.cjs');
  await build({
    entryPoints: [path.resolve('src/PresentationEngine.ts')],
    bundle: true, format: 'cjs', loader: { '.wasm': 'binary' },
    outfile: engineOut, platform: 'node', plugins: [inlineWasm], logLevel: 'silent'
  });
  const { PresentationEngine } = require(engineOut);

  const buffer = toArrayBuffer(await readFile(samplePath));
  const engine = await PresentationEngine.load(buffer);
  const slideCount = engine.slideCount;
  assert.ok(slideCount > 0, 'Sample has no slides.');

  const slide0 = engine.renderSlide(0).svg;
  const intrinsic = getIntrinsicSize(slide0);

  // Zoom frames: same slide at each zoom level.
  const zoomMetrics = [];
  const zoomFrames = ZOOM_LEVELS.map((zoom) => {
    const scaled = computeScaledSize(intrinsic, zoom);
    const shapeCount = (slide0.match(/data-ooxml-shape-idx=/g) || []).length;
    zoomMetrics.push({ zoom, width: scaled.width, height: scaled.height, shapes: shapeCount });
    return `<section class="frame">
      <div class="label"><strong>Zoom ${zoom}x</strong>
        <span class="metric">${scaled.width}x${scaled.height}px</span>
        <span class="metric">${shapeCount} shapes</span></div>
      <div class="pane" style="width:${PANE_WIDTH}px;height:${Math.min(PANE_HEIGHT, scaled.height + PANE_PADDING * 2)}px">
        <div class="surface">${sizeSvg(slide0, scaled.width, scaled.height)}</div>
      </div>
    </section>`;
  });

  // Slide strip: every slide at zoom 1 (render fidelity).
  const stripCells = [];
  for (let index = 0; index < slideCount; index++) {
    const svg = engine.renderSlide(index).svg;
    const scaled = computeScaledSize(getIntrinsicSize(svg), 0.45);
    stripCells.push(`<figure class="cell">
      <div class="surface">${sizeSvg(svg, scaled.width, scaled.height)}</div>
      <figcaption>Slide ${index + 1}</figcaption>
    </figure>`);
  }

  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8">
<title>Native PowerPoint Zoom Visual</title><style>
  * { box-sizing: border-box; }
  body { background:#eef1f6; color:#1f2937; font-family:Arial, sans-serif; margin:0; padding:24px; }
  h1 { font-size:22px; margin:0 0 4px; } .hint{ color:#4b5563; margin:0 0 18px; }
  .frame { margin:0 0 20px; } .label{ display:flex; gap:10px; align-items:center; margin:0 0 6px; }
  .metric{ background:#e0f2fe; border:1px solid #7dd3fc; border-radius:999px; color:#0c4a6e; font-size:12px; padding:2px 8px; }
  .pane{ background:#d9dee9; border:1px solid #aab3c5; border-radius:8px; overflow:auto; padding:${PANE_PADDING}px; }
  .surface{ background:#fff; box-shadow:0 8px 24px rgba(15,23,42,.18); display:inline-block; }
  .surface svg{ display:block; }
  .strip{ display:flex; flex-wrap:wrap; gap:14px; } .cell{ margin:0; } .cell figcaption{ color:#4b5563; font-size:12px; margin-top:4px; text-align:center; }
</style></head><body>
  <h1>Native PowerPoint — zoom + render visual</h1>
  <p class="hint">Real slide rendered at the plugin's zoom envelope (0.25x–3x) using its exact sizing math, plus every slide at a fit-to-strip scale.</p>
  <div id="zoom">${zoomFrames.join('')}</div>
  <h1>All slides</h1>
  <div class="strip">${stripCells.join('')}</div>
  <script>window.addEventListener('load',()=>{document.body.dataset.metrics=encodeURIComponent(JSON.stringify(${JSON.stringify(zoomMetrics)}))});</script>
</body></html>`;

  await mkdir(outputDir, { recursive: true });
  await writeFile(htmlPath, html, 'utf8');

  // Assertions on the zoom sizing math (zoomed out shrinks, zoomed in grows, clamped).
  for (let i = 1; i < zoomMetrics.length; i++) {
    assert.ok(zoomMetrics[i].width >= zoomMetrics[i - 1].width, `width should not shrink from ${ZOOM_LEVELS[i - 1]}x to ${ZOOM_LEVELS[i]}x`);
  }
  assert.ok(zoomMetrics[0].width < zoomMetrics[2].width, 'zoomed-out (0.25x) must be smaller than default (1x)');
  assert.ok(zoomMetrics.at(-1).width > zoomMetrics[2].width, 'zoomed-in (3x) must be larger than default (1x)');
  assert.ok(zoomMetrics.every((m) => m.shapes > 0), 'every zoom frame must contain rendered shapes (non-blank)');

  const chromePath = findChrome();
  const url = pathToFileURL(htmlPath).href;
  const userDataDir = path.join(os.tmpdir(), `native-powerpoint-visual-chrome-${process.pid}`);
  const baseArgs = ['--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
    `--user-data-dir=${userDataDir}`, '--force-device-scale-factor=1', '--virtual-time-budget=2500'];

  await runChrome(chromePath, [...baseArgs, '--window-size=1120,1700', '--hide-scrollbars',
    `--screenshot=${zoomShotPath}`, url], { resolveOnTimeout: () => fileHasBytes(zoomShotPath) });
  await runChrome(chromePath, [...baseArgs, '--window-size=1120,900', '--hide-scrollbars',
    `--screenshot=${slidesShotPath}`, url], { resolveOnTimeout: () => fileHasBytes(slidesShotPath) });

  assert.ok(fileHasBytes(zoomShotPath), 'zoom screenshot was not produced');
  assert.ok(fileHasBytes(slidesShotPath), 'slides screenshot was not produced');

  const dump = await runChrome(chromePath, [...baseArgs, '--dump-dom', url],
    { resolveOnTimeout: ({ stdout }) => stdout.includes('data-metrics=') });
  const match = dump.stdout.match(/data-metrics="([^"]+)"/);
  assert.ok(match, 'zoom metrics not emitted by page');
  JSON.parse(decodeURIComponent(match[1]));

  console.log(`PowerPoint zoom visual passed: ${path.relative(process.cwd(), samplePath)} (${slideCount} slides)`);
  console.log(`Zoom screenshot:   ${zoomShotPath}`);
  console.log(`Slides screenshot: ${slidesShotPath}`);
  console.table(zoomMetrics);
} finally {
  await rm(tmp, { recursive: true, force: true });
}
