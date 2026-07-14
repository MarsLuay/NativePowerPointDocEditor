import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import Module, { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
import { extractZip } from "pptx-svg";

import { exportBytes, loadEngine } from "./helpers/pptx-action-harness.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

let tempDirectoryPromise;
function getTempDirectory() {
  tempDirectoryPromise ??= mkdtemp(path.join(tmpdir(), "native-powerpoint-history-matrix-"));
  return tempDirectoryPromise;
}

// HistoryController is a browser-facing controller, but its snapshot flow only
// needs an engine and the narrow HistoryHost adapter below. Stub `obsidian` so
// this real controller path remains executable under node:test.
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

    const originalLoad = Module._load;
    Module._load = function load(request, parent, isMain) {
      if (request === "obsidian") {
        return {
          Notice: class {},
          Platform: { isMacOS: false },
        };
      }
      return originalLoad.call(this, request, parent, isMain);
    };
    try {
      return require(outfile);
    } finally {
      Module._load = originalLoad;
    }
  })();
  return historyControllerModulePromise;
}

function createHistoryHost(engine) {
  return {
    t: (key) => key,
    engine,
    currentSlide: 0,
    activeEditor: null,
    ensureEditable: () => true,
    canEdit: () => true,
    clearAutosave() {},
    clearDragState() {},
    clearSelection() {},
    markDirty() {},
    async renderCurrentSlide() {
      return true;
    },
    async renderThumbnails() {},
    scheduleThumbnailRefresh() {},
  };
}

async function slideXml(engine) {
  const zip = await extractZip(await exportBytes(engine));
  const xml = zip.textFiles.get("ppt/slides/slide1.xml");
  assert.ok(xml, "export must contain slide1.xml");
  return xml;
}

test("text-box mutation survives HistoryController undo then redo", async () => {
  const { HistoryController } = await loadHistoryControllerModule();
  const engine = await loadEngine();
  const history = new HistoryController(createHistoryHost(engine));
  const beforeXml = await slideXml(engine);

  const entry = await history.capture("insert text box");
  const shapeIndex = await engine.addTextBox(0);
  assert.equal(Number.isInteger(shapeIndex) && shapeIndex >= 0, true, "engine must insert a text box");
  history.record(entry);
  const afterXml = await slideXml(engine);
  assert.notEqual(afterXml, beforeXml, "mutation must change the serialized slide");
  assert.match(afterXml, /New text/, "inserted text must be exported before undo");

  await history.undo();
  assert.equal(await slideXml(engine), beforeXml, "undo must restore the pre-mutation slide snapshot");
  assert.equal(history.canRedo, true, "undo must make the mutation redoable");

  await history.redo();
  assert.equal(await slideXml(engine), afterXml, "redo must restore the mutated slide snapshot");
  assert.equal(history.canUndo, true, "redo must make the mutation undoable again");
});
