import assert from "node:assert/strict";
import { test } from "node:test";
import { loadNativePowerPointViewModule } from "./helpers/load-plugin-modules.mjs";

function createHarness(NativePowerPointView, { autosaveEnabled = true } = {}) {
  const created = [];
  const modified = [];
  const vault = {
    async createBinary(path, output) {
      created.push({ path, output });
    },
    getAbstractFileByPath() {
      return null;
    },
    async modifyBinary(file, output) {
      modified.push({ file, output });
    },
  };
  const view = new NativePowerPointView({ app: { vault } }, () => ({
    autosaveEnabled,
    yoloMode: false,
  }));
  const file = {
    basename: "fixture",
    extension: "pptx",
    name: "fixture.pptx",
    path: "decks/fixture.pptx",
  };
  const sourcePackage = { hasVbaProject: false };
  view.engine = {
    slideCount: 1,
    async export() {
      return new Uint8Array([1]).buffer;
    },
  };
  view.file = file;
  view.loadedFile = file;
  view.sourcePackage = sourcePackage;
  view.sourceBuffer = new Uint8Array([0]).buffer;
  view.validateExportBeforeSave = async () => sourcePackage;
  return { created, file, modified, sourcePackage, view };
}

test("macro-enabled extensions remain view-only", async () => {
  const {
    isEditablePowerPointExtension,
    isMacroEnabledPowerPointExtension,
    isModernPowerPointExtension,
  } = await loadNativePowerPointViewModule();

  for (const extension of ["pptm", "ppsm", "potm"]) {
    assert.equal(isModernPowerPointExtension(extension), true);
    assert.equal(isMacroEnabledPowerPointExtension(extension), true);
    assert.equal(isEditablePowerPointExtension(extension), false);
  }

  for (const extension of ["pptx", "ppsx", "potx"]) {
    assert.equal(isEditablePowerPointExtension(extension), true);
  }
});

test("rapid edits debounce to one autosave", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { view } = createHarness(NativePowerPointView);
  const timers = new Map();
  const previousWindow = globalThis.window;
  let nextTimer = 0;
  let saves = 0;

  globalThis.window = {
    clearTimeout(timer) {
      timers.delete(timer);
    },
    setTimeout(callback, delay) {
      const timer = ++nextTimer;
      timers.set(timer, { callback, delay });
      return timer;
    },
  };

  try {
    view.saveCurrentPresentation = async () => {
      saves += 1;
      return true;
    };
    view.markDirty();
    view.markDirty();
    view.markDirty();

    assert.equal(timers.size, 1);
    const [{ callback, delay }] = timers.values();
    assert.equal(delay, 1500);
    callback();
    await Promise.resolve();
    assert.equal(saves, 1);
  } finally {
    globalThis.window = previousWindow;
  }
});

test("focusable toolbar controls preserve the active text formatting selection", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { view } = createHarness(NativePowerPointView);
  const snapshot = {
    shapeIndex: 4,
    run: { paragraphIndex: 1, runIndex: 2 },
    ranges: [{ paragraphIndex: 1, start: 3, end: 8 }],
    anchor: { left: 1, top: 2, width: 3, height: 4 },
  };

  view.captureToolbarFormattingSnapshot = () => snapshot;
  view.flushActiveEditor = () => {
    view.toolbarFormattingSnapshot = null;
  };

  view.flushActiveEditorForToolbarInput();

  assert.deepEqual(view.toolbarFormattingSnapshot, snapshot);
});

test("alignment applies to every paragraph touched by a text selection", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { view } = createHarness(NativePowerPointView);
  const ranges = [
    { paragraphIndex: 0, start: 2, end: 7 },
    { paragraphIndex: 1, start: 0, end: 4 },
  ];
  let rangeAlignmentCall = null;
  let wholeParagraphCall = null;

  view.engine = {
    setParagraphAlignment(...args) {
      wholeParagraphCall = args;
      return Promise.resolve();
    },
    setParagraphAlignmentForRanges(...args) {
      rangeAlignmentCall = args;
      return Promise.resolve();
    },
  };
  view.currentSlide = 3;
  view.runTextFormatting = async (_label, apply) => {
    await apply(5, { paragraphIndex: 0, runIndex: 0 }, ranges);
  };

  view.applyAlignment("ctr");
  await Promise.resolve();

  assert.equal(wholeParagraphCall, null);
  assert.deepEqual(rangeAlignmentCall, [3, 5, ranges, "ctr"]);
});

test("queued saves serialize rapid writes and retain the final edit", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { modified, sourcePackage, view } = createHarness(NativePowerPointView, { autosaveEnabled: false });
  let exportCalls = 0;
  let releaseFirstExport;
  const firstExportGate = new Promise((resolve) => {
    releaseFirstExport = resolve;
  });

  view.engine.export = async () => {
    exportCalls += 1;
    if (exportCalls === 1) await firstExportGate;
    return new Uint8Array([exportCalls]).buffer;
  };
  view.validateExportBeforeSave = async () => sourcePackage;
  view.isDirty = true;
  view.editVersion = 1;

  const first = view.saveCurrentPresentation();
  await Promise.resolve();
  assert.equal(exportCalls, 1);

  view.isDirty = true;
  view.editVersion = 2;
  const second = view.saveCurrentPresentation();
  await Promise.resolve();
  assert.equal(exportCalls, 1, "second export started before the first save completed");

  releaseFirstExport();
  assert.deepEqual(await Promise.all([first, second]), [true, true]);
  assert.equal(exportCalls, 2);
  assert.deepEqual(modified.map(({ output }) => new Uint8Array(output)[0]), [1, 2]);
  assert.equal(view.isDirty, false);
});

test("closing while a save is in progress waits, then writes a recovery copy", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { created, view } = createHarness(NativePowerPointView, { autosaveEnabled: false });
  let finishPendingSave;
  view.isDirty = true;
  view.savePromise = new Promise((resolve) => {
    finishPendingSave = resolve;
  });

  const preserve = view.preserveUnsavedChangesForTeardown("closing the view");
  await Promise.resolve();
  assert.equal(created.length, 0);

  finishPendingSave();
  assert.equal(await preserve, true);
  assert.equal(created.length, 1);
  assert.match(created[0].path, /Native PowerPoint recovery/);
});

test("autosave failure during close falls back to a recovery copy", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { created, view } = createHarness(NativePowerPointView);
  view.isDirty = true;
  view.saveCurrentPresentation = async () => false;

  assert.equal(await view.preserveUnsavedChangesForTeardown("closing the view"), true);
  assert.equal(created.length, 1);
  assert.equal(view.isDirty, false);
});

test("failed recovery preserves dirty in-memory edits and prevents close reset", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { view } = createHarness(NativePowerPointView, { autosaveEnabled: false });
  const engine = view.engine;
  view.isDirty = true;
  view.engine.export = async () => {
    throw new Error("simulated vault failure");
  };

  assert.equal(await view.preserveUnsavedChangesForTeardown("closing the view"), false);
  assert.equal(view.isDirty, true);

  view.preserveUnsavedChangesForTeardown = async () => false;
  await view.onClose();
  assert.equal(view.engine, engine);
  assert.notEqual(view.loadedFile, null);
  assert.equal(view.isDirty, true);
});
