import assert from "node:assert/strict";
import { test } from "node:test";
import { buildZip, extractZip } from "pptx-svg";
import {
  loadPowerPointPackageModule,
  loadPresentationEngineModule,
  loadShapeClipboardModule,
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

test("pasted and duplicated shapes receive fresh a16:creationId GUIDs", async () => {
  const { createSlideObjectClipboard, pasteSlideObject } = await loadShapeClipboardModule();
  const input = await readDeck("simple-edit.pptx");
  const inputBuffer = toArrayBuffer(input);

  // Seed a creationId on shape 0 (cNvPr id="2") so paste/duplicate must regenerate it.
  const slidePath = "ppt/slides/slide1.xml";
  const sourceZip = await extractZip(inputBuffer);
  const seedGuid = "{11111111-1111-1111-1111-111111111111}";
  const creationExt =
    '<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">' +
    `<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="${seedGuid}"/>` +
    "</a:ext></a:extLst>";
  const slideXml = sourceZip.textFiles
    .get(slidePath)
    .replace('<p:cNvPr id="2" name="Slide 1 title"/>', `<p:cNvPr id="2" name="Slide 1 title">${creationExt}</p:cNvPr>`);
  assert.match(slideXml, /id="\{11111111-1111-1111-1111-111111111111\}"/, "seed creationId was injected onto shape 0");
  const seeded = await buildZip(inputBuffer, new Map([[slidePath, slideXml]]));

  // Paste the seeded shape twice and "duplicate" it (paste back onto the same slide).
  const clipboard = createSlideObjectClipboard(seeded, 0, 0);
  const firstPaste = await pasteSlideObject(seeded, clipboard, 0);
  const secondPaste = await pasteSlideObject(firstPaste.buffer, clipboard, 0);
  const duplicate = await pasteSlideObject(secondPaste.buffer, clipboard, 0);

  const exportedZip = await extractZip(duplicate.buffer);
  const guids = [];
  for (const [partPath, contents] of exportedZip.textFiles) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(partPath)) continue;
    for (const match of contents.matchAll(/<a16:creationId\b[^>]*\bid="([^"]+)"/g)) {
      guids.push(match[1]);
    }
  }

  assert.equal(guids.length, 4, `expected the seed plus three clones to each carry a creationId, found ${guids.length}`);
  assert.equal(new Set(guids).size, guids.length, "every a16:creationId GUID must remain unique");
  assert.equal(guids.filter((guid) => guid === seedGuid).length, 1, "only the source shape keeps the seed GUID");
});

test("setRunStyleForRange bolds only the selected characters within a paragraph", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  const shapeIndex = 0;
  const paragraphIndex = 0;
  await engine.updateParagraphText(0, shapeIndex, paragraphIndex, "Hello world");
  await engine.setRunStyle(0, shapeIndex, { paragraphIndex, runIndex: 0 }, { bold: false });
  await engine.setRunStyleForRange(0, shapeIndex, paragraphIndex, 6, 11, { bold: true });

  assert.equal(engine.getRunStyle(0, shapeIndex, paragraphIndex, 0)?.bold, false);
  assert.equal(engine.getRunStyle(0, shapeIndex, paragraphIndex, 0)?.fontSizePt, 28);
  assert.equal(engine.getRunStyle(0, shapeIndex, paragraphIndex, 1)?.bold, true);
  assert.equal(engine.isRangeStyled(0, shapeIndex, paragraphIndex, 6, 11, "bold"), true);
  assert.equal(engine.isRangeStyled(0, shapeIndex, paragraphIndex, 0, 5, "bold"), false);
});

test("updateParagraphText preserves line breaks within a paragraph", async () => {
  const { createRequire } = await import("node:module");
  const JSZip = createRequire(import.meta.url)("jszip");
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  await engine.updateParagraphText(0, 0, 0, "Line one\nLine two");

  const exported = await engine.export();
  const zip = await JSZip.loadAsync(exported);
  const slideXml = await zip.files["ppt/slides/slide1.xml"].async("string");
  assert.match(slideXml, /<a:br/);
  assert.match(slideXml, /Line one/);
  assert.match(slideXml, /Line two/);

  const reloaded = await PresentationEngine.load(exported);
  const svg = reloaded.renderSlide(0).svg;
  assert.ok(svg.includes(">one<"));
  assert.ok(svg.includes(">two<"));
  assert.ok((svg.match(/data-ooxml-para-idx="0"/g) || []).length >= 2);
});

test("updateShapeTransform allows shapes outside the slide bounds", async () => {
  const { createRequire } = await import("node:module");
  const JSZip = createRequire(import.meta.url)("jszip");
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  await engine.updateShapeTransform(0, 0, {
    x: -9000000,
    y: 342900,
    cx: 5943600,
    cy: 685800,
    rot: 0
  });

  const exported = await engine.export();
  const zip = await JSZip.loadAsync(exported);
  const slideXml = await zip.files["ppt/slides/slide1.xml"].async("string");
  assert.match(slideXml, /x="-9000000"/);
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
