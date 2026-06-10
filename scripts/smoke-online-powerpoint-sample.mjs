import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { build } from 'esbuild';

const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const sampleArgument = process.argv[2] || process.env.NATIVE_POWERPOINT_ONLINE_SAMPLE || 'test-results/online-powerpoint-samples/suu-example.pptx';
const tempDir = await mkdtemp(path.join(os.tmpdir(), 'native-powerpoint-online-sample-'));
const engineBundlePath = path.join(tempDir, 'PresentationEngine.cjs');
const packageBundlePath = path.join(tempDir, 'PowerPointPackage.cjs');
const samplePath = path.resolve(sampleArgument);
const auditDir = path.resolve(
  'test-results',
  'online-powerpoint-audit',
  path.basename(samplePath, path.extname(samplePath))
);

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
  const document = new DOMParser().parseFromString(normalizedSvg, 'image/svg+xml');
  if (document.getElementsByTagName('parsererror').length > 0) {
    throw new Error('Rendered slide SVG is not valid XML.');
  }
  assert.equal(document.documentElement.localName, 'svg');
  return { document, root: document.documentElement };
}

function getShapeGroups(root) {
  return Array.from(root.getElementsByTagName('g'))
    .filter((group) => group.getAttribute('data-ooxml-shape-idx') !== null);
}

function getShapeByIndex(root, shapeIndex) {
  return getShapeGroups(root)
    .find((group) => Number(group.getAttribute('data-ooxml-shape-idx')) === shapeIndex);
}

function assertValidation(label, result) {
  assert.equal(result.ok, true, `${label}: ${result.errors.join('; ')}`);
}

async function auditRenderedSlides(engine, outputDir) {
  await mkdir(outputDir, { recursive: true });
  const serializer = new XMLSerializer();
  const stats = [];

  for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
    const rendered = engine.renderSlide(slideIndex);
    assert.equal(rendered.slideCount, engine.slideCount);
    assert.match(rendered.svg, /<svg\b/);

    const { root } = parseSvg(rendered.svg);
    engine.applyFontFidelity(root);
    engine.formatChartAxisLabels(root, slideIndex);

    const shapeGroups = getShapeGroups(root);
    const textElements = Array.from(root.getElementsByTagName('text'));
    const chartGroups = shapeGroups.filter((group) => group.getAttribute('data-ooxml-shape-type') === 'chart');
    const tableGroups = shapeGroups.filter((group) => group.getAttribute('data-ooxml-shape-type') === 'table');
    assert.ok(root.getAttribute('viewBox') || root.getAttribute('width'), `Slide ${slideIndex + 1} has no viewport metadata.`);

    const serializedSvg = serializer.serializeToString(root);
    await writeFile(path.join(outputDir, `slide-${String(slideIndex + 1).padStart(2, '0')}.svg`), serializedSvg, 'utf8');
    stats.push({
      slide: slideIndex + 1,
      shapes: shapeGroups.length,
      text: textElements.length,
      charts: chartGroups.length,
      tables: tableGroups.length,
      characters: root.textContent.trim().length
    });
  }

  assert.ok(stats.some((slide) => slide.shapes > 0), 'The sample rendered without any detectable slide objects.');
  return stats;
}

async function validateExport(label, PresentationEngine, packageApi, originalBuffer, exportedBuffer, expectedSlideCount) {
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents,
    validatePowerPointPackageStructure
  } = packageApi;

  assertValidation(label, validatePowerPointPackageStructure(inspectPowerPointPackage(exportedBuffer), expectedSlideCount));
  assertValidation(
    label,
    validatePowerPointExport(
      inspectPowerPointPackage(originalBuffer),
      inspectPowerPointPackage(exportedBuffer),
      expectedSlideCount
    )
  );
  assertValidation(label, await validatePowerPointExportContents(originalBuffer, exportedBuffer));
  await PresentationEngine.validateRoundTrip(exportedBuffer, expectedSlideCount);
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
      entryPoints: [path.resolve('src/PowerPointPackage.ts')],
      bundle: true,
      format: 'cjs',
      outfile: packageBundlePath,
      platform: 'node'
    })
  ]);

  const { PresentationEngine } = require(engineBundlePath);
  const packageApi = require(packageBundlePath);
  const {
    inspectPowerPointPackage,
    validatePowerPointPackageStructure
  } = packageApi;

  const originalBuffer = toArrayBuffer(await readFile(samplePath));
  const originalInspection = inspectPowerPointPackage(originalBuffer);
  assertValidation('source package', validatePowerPointPackageStructure(originalInspection));

  const engine = await PresentationEngine.load(originalBuffer);
  assert.equal(engine.slideCount, originalInspection.slidePaths.length);
  assert.ok(engine.slideCount > 0);

  const stats = await auditRenderedSlides(engine, auditDir);
  const unmodifiedExport = await engine.export();
  await validateExport('unmodified export', PresentationEngine, packageApi, originalBuffer, unmodifiedExport, engine.slideCount);

  const editEngine = await PresentationEngine.load(originalBuffer);
  const originalSlideCount = editEngine.slideCount;
  const slideIndex = 0;
  const insertedShapeIndex = editEngine.addTextBox(slideIndex);
  await editEngine.updateShapeText(slideIndex, insertedShapeIndex, 'Codex online sample text box');
  await editEngine.updateTextRun(slideIndex, insertedShapeIndex, 0, 0, 'Codex inline run edit');

  const { root: editedRoot } = parseSvg(editEngine.renderSlide(slideIndex).svg);
  const insertedShape = getShapeByIndex(editedRoot, insertedShapeIndex);
  assert.ok(insertedShape, 'Inserted text box did not render as a selectable shape.');

  const transform = editEngine.getShapeTransform(insertedShape);
  editEngine.updateShapeTransform(slideIndex, insertedShapeIndex, {
    ...transform,
    x: transform.x + editEngine.pxToEmu(12),
    y: transform.y + editEngine.pxToEmu(10),
    cx: transform.cx + editEngine.pxToEmu(18),
    cy: transform.cy + editEngine.pxToEmu(8),
    rot: transform.rot + editEngine.degreesToOoxml(2)
  });

  const duplicatedShapeIndex = await editEngine.duplicateShape(slideIndex, insertedShapeIndex);
  assert.notEqual(duplicatedShapeIndex, insertedShapeIndex);
  editEngine.deleteShape(slideIndex, duplicatedShapeIndex);

  const clipboard = await editEngine.copyShape(slideIndex, insertedShapeIndex);
  const destinationSlideIndex = originalSlideCount > 1 ? originalSlideCount - 1 : slideIndex;
  const pastedShapeIndex = await editEngine.pasteShape(clipboard, destinationSlideIndex);
  assert.ok(Number.isInteger(pastedShapeIndex));

  const objectEditedExport = await editEngine.export();
  await validateExport('object edits', PresentationEngine, packageApi, originalBuffer, objectEditedExport, originalSlideCount);

  const addedSlide = await editEngine.addSlide(slideIndex);
  assert.equal(addedSlide.slideCount, originalSlideCount + 1);
  editEngine.renderSlide(addedSlide.slideIndex);

  const movedLeft = await editEngine.moveSlide(addedSlide.slideIndex, -1);
  assert.equal(movedLeft.slideCount, originalSlideCount + 1);
  const movedRight = await editEngine.moveSlide(movedLeft.slideIndex, 1);
  assert.equal(movedRight.slideCount, originalSlideCount + 1);

  const deletedSlide = await editEngine.deleteSlide(movedRight.slideIndex);
  assert.equal(deletedSlide.slideCount, originalSlideCount);

  const slideEditedExport = await editEngine.export();
  await validateExport('slide edits', PresentationEngine, packageApi, originalBuffer, slideEditedExport, originalSlideCount);

  console.log(`Online PowerPoint sample smoke passed: ${path.relative(process.cwd(), samplePath)}`);
  console.table(stats);
  console.log(`Rendered audit SVGs: ${auditDir}`);
} finally {
  await rm(tempDir, { recursive: true, force: true });
}
