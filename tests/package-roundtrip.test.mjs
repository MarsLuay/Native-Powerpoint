import assert from "node:assert/strict";
import { test } from "node:test";
import { extractZip } from "pptx-svg";
import {
  loadPowerPointPackageModule,
  loadPresentationEngineModule,
} from "./helpers/load-plugin-modules.mjs";
import { createRenderer, readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const editableFixtures = ["features.pptx", "features.ppsx", "features.potx"];
const macroFixtures = ["macro-view-only.pptm", "macro-view-only.ppsm", "macro-view-only.potm"];

test("pptx, ppsx, and potx fixtures load, render, export, and validate", async (t) => {
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents,
    validatePowerPointPackageStructure,
  } = await loadPowerPointPackageModule();

  for (const name of editableFixtures) {
    await t.test(name, async () => {
      const input = await readDeck(name);
      const original = inspectPowerPointPackage(toArrayBuffer(input));
      assert.equal(validatePowerPointPackageStructure(original, 1).ok, true);

      const renderer = await createRenderer(input);
      assert.equal(renderer.getSlideCount(), 1);
      assert.match(renderer.renderSlideSvg(0), /^<svg\b/);

      const output = await renderer.exportPptx();
      const exported = inspectPowerPointPackage(output);
      assert.deepEqual(validatePowerPointExport(original, exported, 1).errors, []);
      assert.deepEqual((await validatePowerPointExportContents(toArrayBuffer(input), output)).errors, []);
    });
  }
});

test("feature fixture preserves rich OOXML parts during an untouched round trip", async () => {
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents,
  } = await loadPowerPointPackageModule();
  const input = await readDeck("features.pptx");
  const original = inspectPowerPointPackage(toArrayBuffer(input));
  const renderer = await createRenderer(input);

  assert.deepEqual(renderer.getSlideNotes(0), ["Fixture speaker notes survive round trip."]);
  const initialSlide = renderer.getSlideOoxml(0);
  assert.match(initialSlide, /<a:hlinkClick\b/);
  assert.match(initialSlide, /<c:chart\b/);
  assert.match(initialSlide, /<a:tbl>/);
  assert.match(initialSlide, /<p:grpSp>/);
  assert.match(initialSlide, /<p:timing>/);
  assert.match(initialSlide, /preserve="unknown-ooxml"/);

  const output = await renderer.exportPptx();
  const exported = inspectPowerPointPackage(output);
  assert.deepEqual(validatePowerPointExport(original, exported, 1).errors, []);
  assert.deepEqual((await validatePowerPointExportContents(toArrayBuffer(input), output)).errors, []);

  const originalZip = await extractZip(toArrayBuffer(input));
  const exportedZip = await extractZip(output);
  for (const path of [
    "ppt/theme/theme1.xml",
    "ppt/notesSlides/notesSlide1.xml",
    "ppt/charts/chart1.xml",
    "customXml/native-powerpoint-extension.xml",
  ]) {
    assert.equal(exportedZip.textFiles.get(path), originalZip.textFiles.get(path), `${path} changed`);
  }
  assert.deepEqual(exportedZip.binaryFiles.get("ppt/media/image1.png"), originalZip.binaryFiles.get("ppt/media/image1.png"));

  const exportedSlide = exportedZip.textFiles.get("ppt/slides/slide1.xml");
  const exportedRelationships = exportedZip.textFiles.get("ppt/slides/_rels/slide1.xml.rels");
  assert.match(exportedSlide, /<c:chart\b/);
  assert.match(exportedSlide, /<a:tbl>/);
  assert.match(exportedSlide, /<p:grpSp>/);
  assert.match(exportedSlide, /<p:timing>/);
  assert.match(exportedSlide, /preserve="unknown-ooxml"/);
  assert.ok(exportedRelationships.includes('Target="https://example.com/native-powerpoint"'));
  assert.match(exportedRelationships, /TargetMode="External"/);
});

test("content validation blocks lossy renderer rewrites of opaque slide markup", async () => {
  const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
  const input = await readDeck("features.pptx");
  const renderer = await createRenderer(input);

  assert.doesNotMatch(renderer.updateShapeText(0, 0, 0, 0, "Edited fixture title"), /^ERROR:/);
  const validation = await validatePowerPointExportContents(toArrayBuffer(input), await renderer.exportPptx());
  assert.equal(validation.ok, false);
  assert.ok(validation.errors.some((error) => error.includes("chart")));
  assert.ok(validation.errors.some((error) => error.includes("unknown OOXML element <np:feature>")));
});

test("macro-enabled fixtures preserve VBA bytes during a renderer round trip", async (t) => {
  const { inspectPowerPointPackage, validatePowerPointExport } = await loadPowerPointPackageModule();

  for (const name of macroFixtures) {
    await t.test(name, async () => {
      const input = await readDeck(name);
      const original = inspectPowerPointPackage(toArrayBuffer(input));
      assert.equal(original.hasVbaProject, true);

      const renderer = await createRenderer(input);
      const output = await renderer.exportPptx();
      const exported = inspectPowerPointPackage(output);
      assert.deepEqual(validatePowerPointExport(original, exported, 1).errors, []);

      const originalZip = await extractZip(toArrayBuffer(input));
      const exportedZip = await extractZip(output);
      assert.deepEqual(exportedZip.binaryFiles.get("ppt/vbaProject.bin"), originalZip.binaryFiles.get("ppt/vbaProject.bin"));
    });
  }
});

test("malformed ZIP fixtures are rejected without crashing", async () => {
  const { inspectPowerPointPackage, validatePowerPointPackageStructure } = await loadPowerPointPackageModule();

  for (const name of ["malformed-random.pptx", "malformed-truncated.pptx"]) {
    const input = await readDeck(name);
    assert.throws(() => inspectPowerPointPackage(toArrayBuffer(input)), /ZIP|Open XML/);
  }

  const unsafe = inspectPowerPointPackage(toArrayBuffer(await readDeck("malformed-unsafe-path.pptx")));
  assert.equal(validatePowerPointPackageStructure(unsafe).ok, false);
  assert.deepEqual(unsafe.unsafePaths, ["../escape.xml"]);

  const duplicate = inspectPowerPointPackage(toArrayBuffer(await readDeck("malformed-duplicate-entry.pptx")));
  assert.equal(validatePowerPointPackageStructure(duplicate).ok, false);
  assert.deepEqual(duplicate.duplicateEntries, ["ppt/presentation.xml"]);
});

test("rapid text edits export one valid final presentation", async () => {
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents,
  } = await loadPowerPointPackageModule();
  const input = await readDeck("simple-edit.pptx");
  const original = inspectPowerPointPackage(toArrayBuffer(input));
  const renderer = await createRenderer(input);

  for (let index = 0; index < 40; index += 1) {
    assert.doesNotMatch(renderer.updateShapeText(0, 0, 0, 0, `Rapid edit ${index}`), /^ERROR:/);
  }

  const output = await renderer.exportPptx();
  const exported = inspectPowerPointPackage(output);
  assert.deepEqual(validatePowerPointExport(original, exported, 1).errors, []);
  assert.deepEqual((await validatePowerPointExportContents(toArrayBuffer(input), output)).errors, []);

  const reloaded = await createRenderer(new Uint8Array(output));
  assert.match(reloaded.getSlideOoxml(0), /Rapid edit 39/);
});

test("slide add, reorder, and delete operations survive export", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("simple-edit.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  assert.equal((await engine.addSlide(0)).slideCount, 2);
  assert.equal((await engine.addSlide(1)).slideCount, 3);
  assert.equal((await engine.moveSlide(2, -1)).slideCount, 3);
  assert.equal((await engine.deleteSlide(1)).slideCount, 2);

  const reloaded = await createRenderer(new Uint8Array(await engine.export()));
  assert.equal(reloaded.getSlideCount(), 2);
  assert.match(reloaded.renderSlideSvg(0), /^<svg\b/);
  assert.match(reloaded.renderSlideSvg(1), /^<svg\b/);
});

test("large 160-slide fixture loads, renders its bounds, and round-trips", async () => {
  const { inspectPowerPointPackage, validatePowerPointExport } = await loadPowerPointPackageModule();
  const input = await readDeck("large-deck.pptx");
  const original = inspectPowerPointPackage(toArrayBuffer(input));
  assert.equal(original.slidePaths.length, 160);

  const renderer = await createRenderer(input);
  assert.equal(renderer.getSlideCount(), 160);
  assert.match(renderer.renderSlideSvg(0), /^<svg\b/);
  assert.match(renderer.getSlideOoxml(159), /Large deck slide 160/);

  const output = await renderer.exportPptx();
  const exported = inspectPowerPointPackage(output);
  assert.deepEqual(validatePowerPointExport(original, exported, 160).errors, []);

  const reloaded = await createRenderer(new Uint8Array(output));
  assert.equal(reloaded.getSlideCount(), 160);
  assert.match(reloaded.getSlideOoxml(159), /Large deck slide 160/);
});
