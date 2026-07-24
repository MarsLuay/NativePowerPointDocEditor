import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import Module, { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

// Audit companion test for the "PPTX · Present, inspector, fonts, history" group.
//
// SCOPE (deliberately narrow — see PART 2 of the audit):
//   * src/FontFidelity.ts        -> the pure fallback map exposed by
//                                   getRendererFallbacks() (no canvas needed).
//   * src/powerpoint/historyController.ts -> capture/undo/redo + the 20-entry
//                                   cap, driven by a minimal fake HistoryHost.
//
// EXPLICITLY SKIPPED as DOM-only (not exercisable under node:test):
//   * PowerPointPresent.ts overlay/keyboard/fullscreen flow — requires a live
//     `document.body`, fullscreen API, and KeyboardEvent dispatch.
//   * FontFidelity canvas font *probing* (createBrowserFontAvailabilityDetector
//     / createBrowserTextMeasurer) — depends on a 2D canvas context.
//   * Inspector layout/background Apply buttons — DOM + PresentationEngine/WASM.

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

let tempDirectoryPromise;
function getTempDirectory() {
  tempDirectoryPromise ??= mkdtemp(path.join(tmpdir(), "native-powerpoint-audit-fonts-history-"));
  return tempDirectoryPromise;
}

// FontFidelity.ts is bundle-clean on node (the smoke:fonts script does the same):
// `pptx-svg`'s DEFAULT_FONT_FALLBACKS is a plain value, and the constructor's
// canvas detectors degrade to no-ops when `window` is undefined.
let fontFidelityModulePromise;
function loadFontFidelityModule() {
  fontFidelityModulePromise ??= (async () => {
    const outfile = path.join(await getTempDirectory(), "font-fidelity.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/FontFidelity.ts")],
      bundle: true,
      format: "cjs",
      logLevel: "silent",
      outfile,
      platform: "node",
      target: "node22",
    });
    return require(outfile);
  })();
  return fontFidelityModulePromise;
}

// historyController.ts pulls `Notice` and `Platform` from `obsidian` at runtime
// (the PresentationEngine/HistoryEntry imports are type-only and erased, while
// ./constants + ./runtimeCompat are pure). So we externalize `obsidian` and
// hand it a stub via Module._load, exactly like the view loader does.
let historyControllerModulePromise;
function loadHistoryControllerModule() {
  historyControllerModulePromise ??= (async () => {
    const outfile = path.join(await getTempDirectory(), "history-controller.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/powerpoint/historyController.ts")],
      bundle: true,
      external: ["obsidian"],
      format: "cjs",
      logLevel: "silent",
      outfile,
      platform: "node",
      target: "node22",
    });

    const notices = [];
    const originalLoad = Module._load;
    Module._load = function load(request, parent, isMain) {
      if (request === "obsidian") {
        return {
          Notice: class {
            constructor(message) {
              notices.push(message);
            }
          },
          Platform: { isMacOS: false },
        };
      }
      return originalLoad.call(this, request, parent, isMain);
    };
    try {
      return { ...require(outfile), notices };
    } finally {
      Module._load = originalLoad;
    }
  })();
  return historyControllerModulePromise;
}

// ---------------------------------------------------------------------------
// FontFidelity: the renderer fallback map (pure, DOM-free)
// ---------------------------------------------------------------------------

test("getRendererFallbacks maps known PowerPoint fonts to cross-platform stacks", async () => {
  const { FontFidelity } = await loadFontFidelityModule();

  // Inject stub canvas hooks so the constructor never touches a 2D context.
  const fidelity = new FontFidelity({
    isFontAvailable: () => true,
    measureText: (text, _fontFamily, fontSizePx) => text.length * fontSizePx,
  });

  const fallbacks = fidelity.getRendererFallbacks();

  // PowerPoint-specific additions (POWERPOINT_FONT_FALLBACKS).
  assert.deepEqual(fallbacks.Aptos, ["Arial", "Helvetica Neue", "Helvetica"]);
  assert.deepEqual(fallbacks["Aptos Display"], ["Arial", "Helvetica Neue", "Helvetica"]);
  assert.deepEqual(fallbacks["Calibri Light"], ["Calibri", "Arial", "Helvetica"]);
  assert.deepEqual(fallbacks.Consolas, ["Courier New", "monospace"]);
  assert.deepEqual(fallbacks.Helvetica, ["Arial", "sans-serif"]);
  assert.equal(fallbacks.Aptos[0], "Arial", "Aptos resolves to Arial first");
  assert.equal(fallbacks["Calibri Light"][0], "Calibri", "Calibri Light resolves to Calibri first");

  // Inherited from pptx-svg DEFAULT_FONT_FALLBACKS and preserved by the merge.
  assert.deepEqual(fallbacks.Calibri, ["Helvetica Neue", "Helvetica", "Arial"]);
  assert.deepEqual(fallbacks["Segoe UI"], ["SF Pro Text", "Helvetica Neue", "Arial"]);
  assert.ok(Array.isArray(fallbacks["Malgun Gothic"]), "Korean default fallback survives the merge");
});

test("getRendererFallbacks returns a defensive copy", async () => {
  const { FontFidelity } = await loadFontFidelityModule();
  const fidelity = new FontFidelity({ isFontAvailable: () => true, measureText: () => 0 });

  const first = fidelity.getRendererFallbacks();
  first.Aptos = ["mutated"];
  first.InjectedFont = ["nope"];

  const second = fidelity.getRendererFallbacks();
  assert.deepEqual(second.Aptos, ["Arial", "Helvetica Neue", "Helvetica"], "external mutation must not leak back in");
  assert.equal(second.InjectedFont, undefined, "added keys must not persist on the controller");
});

test("constructor merges caller-supplied fallbacks over the PowerPoint defaults", async () => {
  const { FontFidelity } = await loadFontFidelityModule();
  const fidelity = new FontFidelity({
    isFontAvailable: () => true,
    measureText: () => 0,
    fontFallbacks: { Aptos: ["Inter"], "Brand Sans": ["Arial"] },
  });

  const fallbacks = fidelity.getRendererFallbacks();
  assert.deepEqual(fallbacks.Aptos, ["Inter"], "caller override wins");
  assert.deepEqual(fallbacks["Brand Sans"], ["Arial"], "caller additions are present");
  assert.deepEqual(fallbacks.Consolas, ["Courier New", "monospace"], "untouched defaults remain");
});

// ---------------------------------------------------------------------------
// HistoryController: capture / undo / redo / 20-entry cap (fake host)
// ---------------------------------------------------------------------------

function makeBuffer(value) {
  const buffer = new ArrayBuffer(8);
  new DataView(buffer).setFloat64(0, value);
  return buffer;
}

function readBuffer(buffer) {
  return new DataView(buffer).getFloat64(0);
}

// Minimal stand-in for the slice of NativePowerPointView that HistoryHost needs.
// `engine.state` is a single number standing in for the document;
// snapshotAuthoritativePackage()/export() snapshot it and restoreSnapshot()
// reinstates it, so undo/redo are observable.
function createFakeHost() {
  const engine = {
    slideCount: 3,
    state: 0,
    async export() {
      return makeBuffer(this.state);
    },
    async snapshotAuthoritativePackage() {
      return makeBuffer(this.state);
    },
    async restoreSnapshot(buffer) {
      this.state = readBuffer(buffer);
    },
  };
  const calls = { render: 0, thumbnails: 0, markDirty: 0 };
  const host = {
    t: (key) => key,
    engine,
    currentSlide: 0,
    activeEditor: null,
    canEditValue: true,
    ensureEditableValue: true,
    ensureEditable() {
      return this.ensureEditableValue;
    },
    canEdit() {
      return this.canEditValue;
    },
    clearAutosave() {},
    clearDragState() {},
    clearSelection() {},
    markDirty() {
      calls.markDirty += 1;
    },
    async renderCurrentSlide() {
      calls.render += 1;
      return true;
    },
    async renderThumbnails() {
      calls.thumbnails += 1;
    },
    scheduleThumbnailRefresh() {},
  };
  return { host, engine, calls };
}

test("capture snapshots current state; undo then redo round-trips the document", async () => {
  const { HistoryController } = await loadHistoryControllerModule();
  const { host, engine } = createFakeHost();
  const controller = new HistoryController(host);

  // Simulate one edit: snapshot the pre-edit state, mutate, then record it.
  engine.state = 10;
  const entry = await controller.capture("edit");
  assert.equal(readBuffer(entry.buffer), 10, "capture records the live buffer");
  assert.equal(entry.label, "edit");
  engine.state = 20;
  controller.record(entry);

  assert.equal(controller.canUndo, true);
  assert.equal(controller.canRedo, false);

  await controller.undo();
  assert.equal(engine.state, 10, "undo restores the captured snapshot");
  assert.equal(controller.canUndo, false);
  assert.equal(controller.canRedo, true);

  await controller.redo();
  assert.equal(engine.state, 20, "redo re-applies the undone state");
  assert.equal(controller.canUndo, true);
  assert.equal(controller.canRedo, false);
});

test("record() caps the undo stack at HISTORY_LIMIT (20) entries", async () => {
  const { HistoryController } = await loadHistoryControllerModule();
  const { host } = createFakeHost();
  const controller = new HistoryController(host);

  for (let index = 0; index < 25; index += 1) {
    controller.record({ buffer: makeBuffer(index), currentSlide: 0, label: `edit ${index}` });
  }

  // canUndo only exposes a boolean, so drain the stack to count the survivors.
  let undoCount = 0;
  while (controller.canUndo) {
    await controller.undo();
    undoCount += 1;
    assert.ok(undoCount <= 25, "undo must terminate");
  }
  assert.equal(undoCount, 20, "only the most recent 20 entries are retained");
});

test("undo is blocked (and bails out) while a text editor is focused", async () => {
  const { HistoryController } = await loadHistoryControllerModule();
  const { host, engine } = createFakeHost();
  const controller = new HistoryController(host);

  controller.record({ buffer: makeBuffer(99), currentSlide: 0, label: "edit" });
  engine.state = 5;

  let blurred = false;
  host.activeEditor = { blur() { blurred = true; } };

  await controller.undo();
  assert.equal(blurred, true, "the active editor is blurred instead of undoing");
  assert.equal(engine.state, 5, "document state is left untouched while editing");
  assert.equal(controller.canUndo, true, "the history entry is not consumed");
  assert.equal(controller.canRedo, false);
});

test("history toolbar enables an active inline undo entry", async () => {
  const { HistoryController } = await loadHistoryControllerModule();
  const { host } = createFakeHost();
  host.canUndoInlineEdit = () => true;
  host.canRedoInlineEdit = () => false;
  const controller = new HistoryController(host);
  const undoButton = {
    disabled: true,
    toggleClass() {},
    setAttribute() {},
  };
  const redoButton = {
    disabled: true,
    toggleClass() {},
    setAttribute() {},
  };

  controller.attachButtons(undoButton, redoButton);
  controller.updateAvailability();

  assert.equal(undoButton.disabled, false);
  assert.equal(redoButton.disabled, true);
});

test("undo/redo no-op when the host is not editable", async () => {
  const { HistoryController } = await loadHistoryControllerModule();
  const { host, engine } = createFakeHost();
  const controller = new HistoryController(host);

  controller.record({ buffer: makeBuffer(7), currentSlide: 0, label: "edit" });
  engine.state = 3;
  host.ensureEditableValue = false;

  await controller.undo();
  assert.equal(engine.state, 3, "ensureEditable() === false short-circuits undo");
  assert.equal(controller.canUndo, true, "the entry remains for when editing is restored");
});
