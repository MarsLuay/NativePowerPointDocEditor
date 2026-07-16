import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { build } from "esbuild";
import { assertExportRoundTrips, assertShapeIndex, loadEngine } from "./helpers/pptx-action-harness.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
let modulesPromise;

async function loadMutationModules() {
  modulesPromise ??= (async () => {
    const outputDirectory = await mkdtemp(path.join(tmpdir(), "npde-mutations-"));
    const serviceOutfile = path.join(outputDirectory, "PptxMutationService.cjs");
    const sessionOutfile = path.join(outputDirectory, "PresentationSession.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/powerpoint/mutations/PptxMutationService.ts")],
      bundle: true,
      format: "cjs",
      logLevel: "silent",
      outfile: serviceOutfile,
      platform: "node",
      target: "node22",
      external: ["obsidian", "pptx-svg", "pptx-svg/wasm"],
    });
    await build({
      entryPoints: [path.join(projectRoot, "src/powerpoint/session/PresentationSession.ts")],
      bundle: true,
      format: "cjs",
      logLevel: "silent",
      outfile: sessionOutfile,
      platform: "node",
      target: "node22",
      external: ["obsidian", "pptx-svg", "pptx-svg/wasm"],
    });

    const Module = require("node:module");
    const originalLoad = Module._load;
    Module._load = function load(request, parent, isMain) {
      if (request === "obsidian") return { Notice: class Notice {} };
      if (request === "pptx-svg") {
        return {
          PptxRenderer: class PptxRenderer {},
          buildZip: async () => new ArrayBuffer(),
          extractZip: async () => ({}),
        };
      }
      if (request === "pptx-svg/wasm") return new ArrayBuffer();
      return originalLoad.call(this, request, parent, isMain);
    };
    try {
      return { ...require(serviceOutfile), ...require(sessionOutfile) };
    } finally {
      Module._load = originalLoad;
    }
  })();
  return modulesPromise;
}

function createSaveHost() {
  return {
    app: {},
    t: (key) => key,
    statusEl: null,
    getSettings: () => ({ autosaveEnabled: false }),
    getFile: () => null,
    getLoadedFile: () => null,
    getEngine: () => null,
    getSourcePackage: () => null,
    getSourceBuffer: () => null,
    setSource: () => {},
    isCurrentPresentation: () => false,
    ensureEditable: () => true,
    getViewOnlyReason: () => "",
  };
}

test("session command mutations commit a loadable package", async () => {
  const { PresentationSession, PptxMutationService } = await loadMutationModules();
  const engine = await loadEngine("features.pptx");
  globalThis.window ??= {
    setTimeout,
    clearTimeout,
    document: {
      createElement: () => ({
        getContext: () => ({
          font: "",
          measureText: (text) => ({ width: text.length }),
        }),
      }),
    },
  };
  const session = new PresentationSession(createSaveHost(), {
    mutationExecutor: new PptxMutationService(engine),
  });

  const shapeIndex = await session.applyCommand({ type: "insert-text-box", slideIndex: 0 });
  assertShapeIndex("insert text box", shapeIndex);
  await session.applyCommand({
    type: "update-shape-text",
    slideIndex: 0,
    shapeIndex,
    text: "Mutation service text",
  });

  assert.equal(session.dirty, true);
  assert.match(engine.getSlideXml(0), /Mutation service text/);
  await assertExportRoundTrips("mutation service", engine);
  session.reset();
});

test("PptxMutationService commits successful mutations and restores failed ones", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const snapshot = new ArrayBuffer(4);
  const engine = {
    export: async () => {
      calls.push("export");
      return snapshot;
    },
    updateShapeText: async (slideIndex, shapeIndex, text) => {
      calls.push(["updateShapeText", slideIndex, shapeIndex, text]);
      if (text === "fail") throw new Error("engine mutation failed");
    },
    commitMutation: async () => calls.push("commit"),
    restoreSnapshot: async (value) => calls.push(["restore", value]),
  };
  const service = new PptxMutationService(engine);

  await service.execute({ type: "update-shape-text", slideIndex: 2, shapeIndex: 3, text: "saved" });
  await assert.rejects(
    () => service.execute({ type: "update-shape-text", slideIndex: 2, shapeIndex: 3, text: "fail" }),
    /engine mutation failed/,
  );

  assert.deepEqual(calls, [
    "export",
    ["updateShapeText", 2, 3, "saved"],
    "commit",
    "export",
    ["updateShapeText", 2, 3, "fail"],
    ["restore", snapshot],
  ]);
});

test("PptxMutationService routes shape fill color commands", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const engine = {
    export: async () => new ArrayBuffer(4),
    setShapeFillColor: async (slideIndex, shapeIndex, hex) => {
      calls.push(["setShapeFillColor", slideIndex, shapeIndex, hex]);
    },
    commitMutation: async () => calls.push("commit"),
    restoreSnapshot: async () => {},
  };

  await new PptxMutationService(engine).execute({
    type: "set-shape-fill-color",
    slideIndex: 2,
    shapeIndex: 3,
    hex: "AABBCC",
  });

  assert.deepEqual(calls, [["setShapeFillColor", 2, 3, "AABBCC"], "commit"]);
});

test("PptxMutationService routes text-box origins", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const engine = {
    export: async () => new ArrayBuffer(4),
    insertTextBox: async (slideIndex, origin) => {
      calls.push(["insertTextBox", slideIndex, origin]);
      return 7;
    },
    commitMutation: async () => calls.push("commit"),
    restoreSnapshot: async () => {},
  };

  const result = await new PptxMutationService(engine).execute({
    type: "insert-text-box",
    slideIndex: 2,
    origin: { x: 123, y: 456 },
  });

  assert.equal(result, 7);
  assert.deepEqual(calls, [["insertTextBox", 2, { x: 123, y: 456 }], "commit"]);
});

test("PptxMutationService routes paragraph split commands", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const engine = {
    export: async () => new ArrayBuffer(4),
    splitParagraph: async (slideIndex, shapeIndex, paragraphIndex, splitOffset, text) => {
      calls.push(["splitParagraph", slideIndex, shapeIndex, paragraphIndex, splitOffset, text]);
      return { paragraphIndex: paragraphIndex + 1 };
    },
    commitMutation: async () => calls.push("commit"),
    restoreSnapshot: async () => {},
  };

  const result = await new PptxMutationService(engine).execute({
    type: "split-paragraph",
    slideIndex: 2,
    shapeIndex: 3,
    paragraphIndex: 4,
    splitOffset: 5,
    text: "Edited paragraph",
  });

  assert.deepEqual(result, { paragraphIndex: 5 });
  assert.deepEqual(calls, [
    ["splitParagraph", 2, 3, 4, 5, "Edited paragraph"],
    "commit",
  ]);
});

test("PptxMutationService rejects commands before an engine is available", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const service = new PptxMutationService(() => null);

  await assert.rejects(
    () => service.execute({ type: "insert-text-box", slideIndex: 0 }),
    /before the presentation engine loaded/,
  );
});
