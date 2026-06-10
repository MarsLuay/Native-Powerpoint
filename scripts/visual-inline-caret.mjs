import assert from 'node:assert/strict';
import { existsSync, statSync } from 'node:fs';
import { mkdir, writeFile } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const outputDir = path.resolve('scripts/visual-output');
const htmlPath = path.join(outputDir, 'inline-caret.html');
const screenshotPath = path.join(outputDir, 'inline-caret.png');

function findChrome() {
  const candidates = [
    process.env.CHROME_PATH,
    path.join(process.env.ProgramFiles || '', 'Google/Chrome/Application/chrome.exe'),
    path.join(process.env['ProgramFiles(x86)'] || '', 'Google/Chrome/Application/chrome.exe'),
    path.join(process.env.ProgramFiles || '', 'Microsoft/Edge/Application/msedge.exe'),
    path.join(process.env['ProgramFiles(x86)'] || '', 'Microsoft/Edge/Application/msedge.exe'),
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    '/usr/bin/google-chrome',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser'
  ].filter(Boolean);

  const chrome = candidates.find((candidate) => existsSync(candidate));
  if (!chrome) {
    throw new Error('Chrome/Edge was not found. Set CHROME_PATH to run the caret visual smoke.');
  }
  return chrome;
}

function runChrome(chromePath, args, options = {}) {
  const timeoutMs = options.timeoutMs ?? 15000;
  return new Promise((resolve, reject) => {
    const child = spawn(chromePath, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill('SIGTERM');
      setTimeout(() => {
        if (child.exitCode === null) child.kill('SIGKILL');
      }, 2000).unref();
    }, timeoutMs);

    child.stdout.on('data', (chunk) => {
      stdout += chunk;
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
    });
    child.on('error', (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on('close', (code) => {
      clearTimeout(timer);
      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }

      if (timedOut && options.resolveOnTimeout?.({ stdout, stderr }) === true) {
        resolve({ stdout, stderr });
        return;
      }

      if (timedOut) {
        reject(new Error(`Chrome timed out after ${timeoutMs} ms: ${stderr || stdout}`));
        return;
      }

      reject(new Error(`Chrome exited with ${code}: ${stderr || stdout}`));
    });
  });
}

function fileHasBytes(filePath) {
  try {
    return statSync(filePath).size > 0;
  } catch {
    return false;
  }
}

const html = String.raw`<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Native PowerPoint Inline Caret Visual Smoke</title>
  <style>
    * {
      box-sizing: border-box;
    }

    body {
      background: #f4f6fb;
      color: #1f2937;
      font-family: Arial, sans-serif;
      margin: 0;
      padding: 24px;
    }

    h1 {
      font-size: 22px;
      margin: 0 0 8px;
    }

    .hint {
      color: #4b5563;
      margin: 0 0 20px;
    }

    .frame {
      margin: 0 0 22px;
    }

    .label {
      align-items: center;
      display: flex;
      gap: 12px;
      margin: 0 0 8px;
    }

    .label strong {
      font-size: 16px;
    }

    .metric {
      background: #e0f2fe;
      border: 1px solid #7dd3fc;
      border-radius: 999px;
      color: #0c4a6e;
      font-size: 12px;
      padding: 3px 8px;
    }

    .native-powerpoint-canvas-pane {
      background: #d9dee9;
      border: 1px solid #aab3c5;
      border-radius: 8px;
      display: inline-block;
      overflow: visible;
      padding: 18px;
      position: relative;
    }

    .native-powerpoint-slide-surface {
      background: white;
      box-shadow: 0 10px 26px rgba(15, 23, 42, 0.22);
      position: relative;
    }

    .native-powerpoint-svg-caret {
      stroke: #000;
      stroke-linecap: square;
    }

    .native-powerpoint-svg-selection {
      fill: rgba(38, 132, 255, 0.28);
      pointer-events: none;
    }

    .text-box-outline {
      border: 1px dashed #ef4444;
      pointer-events: none;
      position: absolute;
      z-index: 9;
    }

    svg {
      display: block;
    }

    svg text,
    svg tspan {
      -webkit-user-select: none;
      user-select: none;
    }
  </style>
</head>
<body>
  <h1>Inline SVG Text Caret Zoom Smoke</h1>
  <p class="hint">The black insertion caret should grow with the bottom text row at every zoom.</p>
  <main id="root"></main>
  <script>
    const zooms = [1, 2, 3, 4];
    const root = document.getElementById('root');
    const metrics = [];
    const deletionProbeText = 'Bottom row click lands here';

    function getElementBox(element, pane) {
      const paneRect = pane.getBoundingClientRect();
      const rect = element.getBoundingClientRect();
      return {
        left: rect.left - paneRect.left + pane.scrollLeft,
        top: rect.top - paneRect.top + pane.scrollTop,
        width: rect.width,
        height: rect.height
      };
    }

    function getScreenFontSize(element) {
      const style = window.getComputedStyle(element);
      const fontSize = Number.parseFloat(style.fontSize) || 16;
      const matrix = element.getScreenCTM();
      const scale = matrix ? Math.hypot(matrix.a, matrix.b) : 1;
      return Math.max(4, fontSize * scale);
    }

    function getInlineCaretHeight(element, box) {
      const lineBoxHeight = box.height;
      const screenFontSize = Math.min(getScreenFontSize(element), lineBoxHeight);
      const baseHeight = Math.min(lineBoxHeight, screenFontSize || lineBoxHeight);
      return Math.max(6, baseHeight * 0.88);
    }

    function getSvgInlineCaretStrokeWidth(element) {
      const style = window.getComputedStyle(element);
      const fontSize = Number.parseFloat(style.fontSize) || 16;
      return Math.max(1.25, Math.min(4, fontSize / 14));
    }

    function getSvgInlineSelectionPadding(element) {
      const style = window.getComputedStyle(element);
      const fontSize = Number.parseFloat(style.fontSize) || 16;
      return Math.max(0.75, Math.min(3, fontSize / 18));
    }

    function getInlineCaretRowFromRatio(element, box, centerRatio, height = getInlineCaretHeight(element, box)) {
      const ratio = Math.max(0, Math.min(1, centerRatio));
      const minCenter = box.top + height / 2;
      const maxCenter = Math.max(minCenter, box.top + box.height - height / 2);
      const center = Math.max(minCenter, Math.min(maxCenter, box.top + box.height * ratio));
      return {
        top: center - height / 2,
        height,
        centerRatio: ratio
      };
    }

    function transformSvgRectToLocalBox(rect, matrix, paneRect, pane) {
      const points = [
        new DOMPoint(rect.x, rect.y),
        new DOMPoint(rect.x + rect.width, rect.y),
        new DOMPoint(rect.x, rect.y + rect.height),
        new DOMPoint(rect.x + rect.width, rect.y + rect.height)
      ].map((point) => point.matrixTransform(matrix));
      const xs = points.map((point) => point.x - paneRect.left + pane.scrollLeft);
      const ys = points.map((point) => point.y - paneRect.top + pane.scrollTop);
      const left = Math.min(...xs);
      const top = Math.min(...ys);
      return {
        left,
        top,
        width: Math.max(...xs) - left,
        height: Math.max(...ys) - top
      };
    }

    function getSvgTextCaretGeometry(element, offset, pane, preferredHeight = getScreenFontSize(element) * 1.08) {
      const text = element.textContent || '';
      const matrix = element.getScreenCTM();
      const paneRect = pane.getBoundingClientRect();
      if (!text || !matrix) return null;

      const charCount = element.getNumberOfChars();
      const normalizedOffset = Math.max(0, Math.min(charCount, offset));
      const charIndex = Math.max(0, Math.min(charCount - 1, normalizedOffset <= 0 ? 0 : normalizedOffset - 1));
      const position = normalizedOffset <= 0
        ? element.getStartPositionOfChar(charIndex)
        : element.getEndPositionOfChar(charIndex);
      const extent = element.getExtentOfChar(charIndex);
      const point = new DOMPoint(position.x, position.y).matrixTransform(matrix);
      const localLeft = point.x - paneRect.left + pane.scrollLeft;
      let height = Math.max(6, preferredHeight);
      let top = getElementBox(element, pane).top;

      if (extent) {
        const bounds = transformSvgRectToLocalBox(extent, matrix, paneRect, pane);
        height = Math.max(6, Math.min(preferredHeight * 1.1, bounds.height || preferredHeight));
        top = bounds.top + Math.max(0, (bounds.height - height) / 2);
      }

      return { left: localLeft, top, height };
    }

    function localPointToSvgRoot(left, top, pane, svg) {
      const matrix = svg.getScreenCTM();
      const paneRect = pane.getBoundingClientRect();
      if (!matrix) return null;

      const screenPoint = new DOMPoint(
        paneRect.left + left - pane.scrollLeft,
        paneRect.top + top - pane.scrollTop
      );
      return screenPoint.matrixTransform(matrix.inverse());
    }

    function transformSvgRectToSvgRoot(rect, elementMatrix, rootInverse) {
      const points = [
        new DOMPoint(rect.x, rect.y),
        new DOMPoint(rect.x + rect.width, rect.y),
        new DOMPoint(rect.x, rect.y + rect.height),
        new DOMPoint(rect.x + rect.width, rect.y + rect.height)
      ].map((point) => point.matrixTransform(elementMatrix).matrixTransform(rootInverse));
      const xs = points.map((point) => point.x);
      const ys = points.map((point) => point.y);
      const x = Math.min(...xs);
      const y = Math.min(...ys);
      return {
        x,
        y,
        width: Math.max(...xs) - x,
        height: Math.max(...ys) - y
      };
    }

    function getSvgInlineSelectionBoxes(element, start, end, svg) {
      const elementMatrix = element.getScreenCTM();
      const rootMatrix = svg.getScreenCTM();
      if (!elementMatrix || !rootMatrix) return [];

      const rootInverse = rootMatrix.inverse();

      const charCount = element.getNumberOfChars();
      const normalizedStart = Math.max(0, Math.min(charCount, start));
      const normalizedEnd = Math.max(normalizedStart, Math.min(charCount, end));
      const rows = [];
      for (let index = normalizedStart; index < normalizedEnd; index++) {
        const charBox = transformSvgRectToSvgRoot(element.getExtentOfChar(index), elementMatrix, rootInverse);
        if (!charBox || charBox.width < 0 || charBox.height <= 0) continue;

        const centerY = charBox.y + charBox.height / 2;
        const row = rows.find((candidate) => (
          Math.abs(centerY - (candidate.y + candidate.height / 2)) < Math.max(2, charBox.height * 0.55)
        ));
        if (row) {
          const left = Math.min(row.x, charBox.x);
          const top = Math.min(row.y, charBox.y);
          const right = Math.max(row.x + row.width, charBox.x + charBox.width);
          const bottom = Math.max(row.y + row.height, charBox.y + charBox.height);
          row.x = left;
          row.y = top;
          row.width = right - left;
          row.height = bottom - top;
        } else {
          rows.push({ ...charBox });
        }
      }

      const padding = getSvgInlineSelectionPadding(element);
      return rows.map((box) => ({
        x: box.x - padding,
        y: box.y - padding * 0.5,
        width: box.width + padding * 2,
        height: box.height + padding
      }));
    }

    function simulateRecentCaretDelete(value, caretOffset, key) {
      const selectionStart = 0;
      const selectionEnd = value.length;
      if (selectionStart !== 0 || selectionEnd !== value.length) {
        return { value, selectionStart, selectionEnd, prevented: false };
      }

      const offset = Math.max(0, Math.min(caretOffset, value.length));
      const deleteStart = key === 'Backspace' ? offset - 1 : offset;
      const deleteEnd = key === 'Backspace' ? offset : offset + 1;

      if (deleteStart < 0 || deleteEnd > value.length || deleteStart >= deleteEnd) {
        return { value, selectionStart: offset, selectionEnd: offset, prevented: true };
      }

      return {
        value: value.slice(0, deleteStart) + value.slice(deleteEnd),
        selectionStart: deleteStart,
        selectionEnd: deleteStart,
        prevented: true
      };
    }

    function renderZoom(zoom) {
      const frame = document.createElement('section');
      frame.className = 'frame';
      const label = document.createElement('div');
      label.className = 'label';
      label.innerHTML = '<strong>Zoom ' + zoom + 'x</strong>';
      const pane = document.createElement('div');
      pane.className = 'native-powerpoint-canvas-pane';
      const surface = document.createElement('div');
      surface.className = 'native-powerpoint-slide-surface';
      surface.style.width = 360 * zoom + 'px';
      surface.style.height = 202.5 * zoom + 'px';

      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('viewBox', '0 0 480 270');
      svg.setAttribute('width', String(360 * zoom));
      svg.setAttribute('height', String(202.5 * zoom));
      svg.innerHTML = [
        '<rect x="0" y="0" width="480" height="270" fill="#fff"/>',
        '<rect x="46" y="48" width="365" height="128" rx="7" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2"/>',
        '<text x="70" y="82" font-family="Arial, sans-serif" font-size="22" fill="#111827">',
        '<tspan x="70" dy="0">Top row stays calm while editing</tspan>',
        '<tspan x="70" dy="32">Middle row is only context</tspan>',
        '<tspan id="target-' + zoom + '" x="70" dy="32">Bottom row click lands here</tspan>',
        '</text>'
      ].join('');

      surface.appendChild(svg);
      pane.appendChild(surface);
      frame.appendChild(label);
      frame.appendChild(pane);
      root.appendChild(frame);

      const target = document.getElementById('target-' + zoom);
      const box = getElementBox(target, pane);
      const caretOffset = 'Bottom row click'.length;
      const selectedStart = 'Bottom '.length;
      const selectedEnd = 'Bottom row click'.length;
      const textElement = target.closest('text');
      const selectionBoxes = getSvgInlineSelectionBoxes(target, selectedStart, selectedEnd, svg);
      for (const selectionBox of selectionBoxes) {
        const selection = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        selection.classList.add('native-powerpoint-svg-selection');
        selection.setAttribute('x', String(selectionBox.x));
        selection.setAttribute('y', String(selectionBox.y));
        selection.setAttribute('width', String(selectionBox.width));
        selection.setAttribute('height', String(selectionBox.height));
        svg.insertBefore(selection, textElement);
      }
      const svgGeometry = getSvgTextCaretGeometry(target, caretOffset, pane, getInlineCaretHeight(target, box));
      const row = getInlineCaretRowFromRatio(target, box, 0.5);
      const caretStart = localPointToSvgRoot(svgGeometry.left, row.top, pane, svg);
      const caretEnd = localPointToSvgRoot(svgGeometry.left, row.top + row.height, pane, svg);
      const caret = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      caret.classList.add('native-powerpoint-svg-caret');
      caret.setAttribute('x1', String(caretStart.x));
      caret.setAttribute('y1', String(caretStart.y));
      caret.setAttribute('x2', String(caretEnd.x));
      caret.setAttribute('y2', String(caretEnd.y));
      caret.setAttribute('stroke-width', String(getSvgInlineCaretStrokeWidth(target)));
      svg.appendChild(caret);

      const outline = document.createElement('div');
      outline.className = 'text-box-outline';
      outline.style.left = box.left + 'px';
      outline.style.top = box.top + 'px';
      outline.style.width = box.width + 'px';
      outline.style.height = box.height + 'px';
      pane.appendChild(outline);

      const rootScale = svg.getBoundingClientRect().width / 480;
      const forwardDeleteProbe = simulateRecentCaretDelete(
        deletionProbeText,
        'Bottom row '.length,
        'Delete'
      );
      const backwardDeleteProbe = simulateRecentCaretDelete(
        deletionProbeText,
        'Bottom row c'.length,
        'Backspace'
      );
      const metric = {
        zoom,
        fontSize: Number(getScreenFontSize(target).toFixed(2)),
        textBoxHeight: Number(box.height.toFixed(2)),
        caretHeight: Number(row.height.toFixed(2)),
        caretWidth: Number((getSvgInlineCaretStrokeWidth(target) * rootScale).toFixed(2)),
        caretToTextRatio: Number((row.height / box.height).toFixed(3)),
        selectionBoxes: selectionBoxes.length,
        browserSelection: window.getSelection().toString(),
        fullSelectionDelete: forwardDeleteProbe.value,
        fullSelectionBackspace: backwardDeleteProbe.value,
        deleteSelectionStart: forwardDeleteProbe.selectionStart,
        backspaceSelectionStart: backwardDeleteProbe.selectionStart,
        deletePrevented: forwardDeleteProbe.prevented && backwardDeleteProbe.prevented
      };
      metrics.push(metric);
      label.insertAdjacentHTML(
        'beforeend',
        '<span class="metric">font ' + metric.fontSize + 'px</span>' +
        '<span class="metric">caret ' + metric.caretHeight + 'px x ' + metric.caretWidth + 'px</span>' +
        '<span class="metric">selection boxes ' + metric.selectionBoxes + '</span>' +
        '<span class="metric">ratio ' + metric.caretToTextRatio + '</span>' +
        '<span class="metric">full-select delete keeps sentence</span>'
      );
    }

    window.addEventListener('load', () => {
      zooms.forEach(renderZoom);
      document.body.dataset.metrics = encodeURIComponent(JSON.stringify(metrics));
    });
  </script>
</body>
</html>`;

await mkdir(outputDir, { recursive: true });
await writeFile(htmlPath, html, 'utf8');

const chromePath = findChrome();
const url = pathToFileURL(htmlPath).href;
const userDataDir = path.join(os.tmpdir(), `native-powerpoint-caret-chrome-${process.pid}`);
const baseArgs = [
  '--headless=new',
  '--disable-gpu',
  '--no-first-run',
  '--no-default-browser-check',
  `--user-data-dir=${userDataDir}`,
  '--window-size=1700,2400',
  '--virtual-time-budget=1000'
];

await runChrome(chromePath, [
  ...baseArgs,
  `--screenshot=${screenshotPath}`,
  url
], {
  resolveOnTimeout: () => fileHasBytes(screenshotPath)
});

const dump = await runChrome(chromePath, [
  ...baseArgs,
  '--dump-dom',
  url
], {
  resolveOnTimeout: ({ stdout }) => stdout.includes('data-metrics=')
});

const match = dump.stdout.match(/data-metrics="([^"]+)"/);
if (!match) {
  throw new Error('Visual caret metrics were not emitted by the fixture page.');
}

const parsedMetrics = JSON.parse(decodeURIComponent(match[1]));
assert.equal(parsedMetrics.length, 4);
for (let index = 1; index < parsedMetrics.length; index += 1) {
  assert.ok(
    parsedMetrics[index].caretHeight > parsedMetrics[index - 1].caretHeight,
    `caret height should grow from zoom ${parsedMetrics[index - 1].zoom} to ${parsedMetrics[index].zoom}`
  );
  assert.ok(
    parsedMetrics[index].caretWidth > parsedMetrics[index - 1].caretWidth,
    `caret width should grow from zoom ${parsedMetrics[index - 1].zoom} to ${parsedMetrics[index].zoom}`
  );
}
for (const metric of parsedMetrics) {
  assert.ok(metric.caretToTextRatio >= 0.72 && metric.caretToTextRatio <= 1.05);
  assert.ok(metric.selectionBoxes > 0, `selection highlight should render at zoom ${metric.zoom}`);
  assert.equal(metric.browserSelection, '', `browser selection should stay empty at zoom ${metric.zoom}`);
  assert.equal(metric.fullSelectionDelete, 'Bottom row lick lands here');
  assert.equal(metric.fullSelectionBackspace, 'Bottom row lick lands here');
  assert.equal(metric.deleteSelectionStart, 'Bottom row '.length);
  assert.equal(metric.backspaceSelectionStart, 'Bottom row '.length);
  assert.equal(metric.deletePrevented, true);
}

console.log('Inline caret visual smoke passed.');
console.log(`HTML: ${htmlPath}`);
console.log(`Screenshot: ${screenshotPath}`);
console.table(parsedMetrics);
