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
  const saveController = view.saveController;
  saveController.validateExportOverride = async () => sourcePackage;
  return { created, file, modified, saveController, sourcePackage, view };
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
  const { saveController, view } = createHarness(NativePowerPointView);
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
    saveController.save = async () => {
      saves += 1;
      return true;
    };
    view.markDirty();
    view.markDirty();
    view.markDirty();

    const autosaveTimers = [...timers.values()].filter((entry) => entry.delay === 1500 || entry.delay === 4000);
    assert.equal(autosaveTimers.length, 1, 'rapid markDirty calls should debounce to one autosave timer');
    autosaveTimers[0].callback();
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
  let alignmentCommand = null;

  view.engine = {};
  view.session.applyCommand = async (command) => {
    alignmentCommand = command;
  };
  view.currentSlide = 3;
  view.runTextFormatting = async (_label, apply) => {
    await apply(5, { paragraphIndex: 0, runIndex: 0 }, ranges);
  };

  view.applyAlignment("ctr");
  await Promise.resolve();

  assert.deepEqual(alignmentCommand, {
    type: "set-paragraph-alignment-ranges",
    slideIndex: 3,
    shapeIndex: 5,
    ranges,
    align: "ctr",
  });
});

test("queued saves serialize rapid writes and retain the final edit", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { modified, saveController, sourcePackage, view } = createHarness(NativePowerPointView, { autosaveEnabled: false });
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
  saveController.validateExportOverride = async () => sourcePackage;
  saveController.isDirty = true;
  saveController.editVersion = 1;

  const first = view.saveCurrentPresentation();
  await Promise.resolve();
  assert.equal(exportCalls, 1);

  saveController.isDirty = true;
  saveController.editVersion = 2;
  const second = view.saveCurrentPresentation();
  await Promise.resolve();
  assert.equal(exportCalls, 1, "second export started before the first save completed");

  releaseFirstExport();
  assert.deepEqual(await Promise.all([first, second]), [true, true]);
  assert.equal(exportCalls, 2);
  assert.deepEqual(modified.map(({ output }) => new Uint8Array(output)[0]), [1, 2]);
  assert.equal(saveController.isDirty, false);
});

test("closing while a save is in progress waits, then writes a recovery copy", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { created, saveController, view } = createHarness(NativePowerPointView, { autosaveEnabled: false });
  let finishPendingSave;
  saveController.isDirty = true;
  saveController.savePromise = new Promise((resolve) => {
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
  const { created, saveController, view } = createHarness(NativePowerPointView);
  saveController.isDirty = true;
  saveController.save = async () => false;

  assert.equal(await view.preserveUnsavedChangesForTeardown("closing the view"), true);
  assert.equal(created.length, 1);
  assert.equal(saveController.isDirty, false);
});

test("plugin reload saves dirty edits to the source even when autosave is disabled", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { file, modified, saveController, view } = createHarness(
    NativePowerPointView,
    { autosaveEnabled: false },
  );
  saveController.isDirty = true;

  assert.equal(await view.saveBeforePluginReload(), true);
  assert.equal(modified.length, 1);
  assert.equal(modified[0].file, file);
  assert.deepEqual(new Uint8Array(modified[0].output), new Uint8Array([1]));
  assert.equal(saveController.isDirty, false);
});

test("plugin reload barrier fails closed when the dirty source cannot be saved", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { modified, saveController, view } = createHarness(
    NativePowerPointView,
    { autosaveEnabled: false },
  );
  saveController.isDirty = true;
  view.engine.export = async () => {
    throw new Error("simulated reload save failure");
  };

  assert.equal(await view.saveBeforePluginReload(), false);
  assert.equal(modified.length, 0);
  assert.equal(saveController.isDirty, true);
});

test("failed autosave schedules a delayed retry", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { saveController, view } = createHarness(NativePowerPointView);
  const timers = new Map();
  const previousWindow = globalThis.window;
  let nextTimer = 0;

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
    saveController.isDirty = true;
    view.engine.export = async () => {
      throw new Error('simulated validation failure');
    };

    const saved = await view.saveCurrentPresentation();
    assert.equal(saved, false);
    assert.equal(saveController.saveState, 'failed');
    assert.equal(saveController.lastSaveError, 'simulated validation failure');
    assert.equal(timers.size, 1);

    const [{ delay }] = timers.values();
    assert.equal(delay, 5000);
  } finally {
    timers.clear();
    await saveController.savePromise.catch(() => undefined);
    globalThis.window = previousWindow;
  }
});

test("failed recovery preserves dirty in-memory edits and prevents close reset", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { saveController, view } = createHarness(NativePowerPointView, { autosaveEnabled: false });
  const engine = view.engine;
  saveController.isDirty = true;
  view.engine.export = async () => {
    throw new Error("simulated vault failure");
  };

  assert.equal(await view.preserveUnsavedChangesForTeardown("closing the view"), false);
  assert.equal(saveController.isDirty, true);

  view.preserveUnsavedChangesForTeardown = async () => false;
  await view.onClose();
  assert.equal(view.engine, engine);
  assert.notEqual(view.loadedFile, null);
  assert.equal(saveController.isDirty, true);
});
