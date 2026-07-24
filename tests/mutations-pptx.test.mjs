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

test("PptxMutationService serializes overlapping engine transactions", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  let releaseFirst;
  let signalFirstStarted;
  const firstStarted = new Promise((resolve) => { signalFirstStarted = resolve; });
  const engine = {
    export: async () => {
      calls.push("export");
      return new ArrayBuffer(4);
    },
    updateShapeText: async (_slideIndex, _shapeIndex, text) => {
      calls.push(["updateShapeText", text]);
      if (text === "first") {
        signalFirstStarted();
        await new Promise((resolve) => { releaseFirst = resolve; });
      }
    },
    commitMutation: async () => calls.push("commit"),
    restoreSnapshot: async () => {},
  };
  const service = new PptxMutationService(engine);

  const first = service.execute({ type: "update-shape-text", slideIndex: 0, shapeIndex: 0, text: "first" });
  await firstStarted;
  const second = service.execute({ type: "update-shape-text", slideIndex: 0, shapeIndex: 0, text: "second" });
  await Promise.resolve();
  assert.deepEqual(calls, ["export", ["updateShapeText", "first"]]);

  releaseFirst();
  await Promise.all([first, second]);
  assert.deepEqual(calls, [
    "export",
    ["updateShapeText", "first"],
    "commit",
    "export",
    ["updateShapeText", "second"],
    "commit",
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

test("PptxMutationService routes empty preceding paragraph removal commands", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const engine = {
    export: async () => new ArrayBuffer(4),
    removeEmptyPrecedingParagraph: async (slideIndex, shapeIndex, paragraphIndex) => {
      calls.push(["removeEmptyPrecedingParagraph", slideIndex, shapeIndex, paragraphIndex]);
      return { removed: true, paragraphIndex: paragraphIndex - 1 };
    },
    commitMutation: async () => calls.push("commit"),
    restoreSnapshot: async () => {},
  };

  const result = await new PptxMutationService(engine).execute({
    type: "remove-empty-preceding-paragraph",
    slideIndex: 2,
    shapeIndex: 3,
    paragraphIndex: 4,
  });

  assert.deepEqual(result, { removed: true, paragraphIndex: 3 });
  assert.deepEqual(calls, [
    ["removeEmptyPrecedingParagraph", 2, 3, 4],
    "commit",
  ]);
});

test("PptxMutationService routes preceding paragraph merge commands", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const engine = {
    export: async () => new ArrayBuffer(4),
    mergePrecedingParagraph: async (slideIndex, shapeIndex, paragraphIndex, text) => {
      calls.push(["mergePrecedingParagraph", slideIndex, shapeIndex, paragraphIndex, text]);
      return { merged: true, paragraphIndex: paragraphIndex - 1, caretOffset: 7 };
    },
    commitMutation: async () => calls.push("commit"),
    restoreSnapshot: async () => {},
  };

  const result = await new PptxMutationService(engine).execute({
    type: "merge-preceding-paragraph",
    slideIndex: 2,
    shapeIndex: 3,
    paragraphIndex: 4,
    text: "Edited paragraph",
  });

  assert.deepEqual(result, { merged: true, paragraphIndex: 3, caretOffset: 7 });
  assert.deepEqual(calls, [
    ["mergePrecedingParagraph", 2, 3, 4, "Edited paragraph"],
    "commit",
  ]);
});

test("PptxMutationService uses slide-XML rollback for slide-local text edits", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const engine = {
    export: async () => {
      calls.push("export");
      return new ArrayBuffer(4);
    },
    getSlideXml: (slideIndex) => {
      calls.push(["getSlideXml", slideIndex]);
      return `<slide-${slideIndex}/>`;
    },
    restoreSlideXml: async (slideIndex, xml) => {
      calls.push(["restoreSlideXml", slideIndex, xml]);
    },
    restoreSnapshot: async () => calls.push("restoreSnapshot"),
    splitParagraph: async (slideIndex, shapeIndex, paragraphIndex, splitOffset, text) => {
      calls.push(["splitParagraph", slideIndex, shapeIndex, paragraphIndex, splitOffset, text]);
      if (text === "fail") throw new Error("split failed");
      return { paragraphIndex: paragraphIndex + 1 };
    },
    commitMutation: async () => calls.push("commit"),
  };
  const service = new PptxMutationService(engine);

  const result = await service.execute({
    type: "split-paragraph",
    slideIndex: 2,
    shapeIndex: 3,
    paragraphIndex: 4,
    splitOffset: 5,
    text: "ok",
  });
  assert.deepEqual(result, { paragraphIndex: 5 });

  await assert.rejects(
    () => service.execute({
      type: "split-paragraph",
      slideIndex: 2,
      shapeIndex: 3,
      paragraphIndex: 4,
      splitOffset: 5,
      text: "fail",
    }),
    /split failed/,
  );

  assert.deepEqual(calls, [
    ["getSlideXml", 2],
    ["splitParagraph", 2, 3, 4, 5, "ok"],
    "commit",
    ["getSlideXml", 2],
    ["splitParagraph", 2, 3, 4, 5, "fail"],
    ["restoreSlideXml", 2, "<slide-2/>"],
  ]);
  assert.ok(!calls.includes("export"), "slide-local edits must not export the full deck");
  assert.ok(!calls.includes("restoreSnapshot"), "slide-local rollback must not restore a full-deck snapshot");
});

test("PptxMutationService uses the deferred slide-local commit when the engine supports it", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const engine = {
    export: async () => {
      calls.push("export");
      return new ArrayBuffer(4);
    },
    getSlideXml: (slideIndex) => {
      calls.push(["getSlideXml", slideIndex]);
      return `<slide-${slideIndex}/>`;
    },
    restoreSlideXml: async (slideIndex, xml) => calls.push(["restoreSlideXml", slideIndex, xml]),
    restoreSnapshot: async () => calls.push("restoreSnapshot"),
    splitParagraph: async (slideIndex, shapeIndex, paragraphIndex, splitOffset) => {
      calls.push(["splitParagraph", slideIndex, shapeIndex, paragraphIndex, splitOffset]);
      return { paragraphIndex: paragraphIndex + 1 };
    },
    commitMutation: async () => calls.push("commit"),
    // Deferred commit: recorded as called, but must not export/sync the package.
    commitSlideLocalMutation: async () => calls.push("commit-slide-local"),
  };
  const service = new PptxMutationService(engine);

  const result = await service.execute({
    type: "split-paragraph",
    slideIndex: 2,
    shapeIndex: 3,
    paragraphIndex: 4,
    splitOffset: 5,
  });
  assert.deepEqual(result, { paragraphIndex: 5 });

  assert.deepEqual(calls, [
    ["getSlideXml", 2],
    ["splitParagraph", 2, 3, 4, 5],
    "commit-slide-local",
  ]);
  assert.ok(!calls.includes("export"), "slide-local commit must not export the full deck");
  assert.ok(!calls.includes("commit"), "slide-local commit must not run the full-export commit");
});

test("deferred slide-local commit keeps edits through a later reorder", async () => {
  // Regression for the residual keystroke lag fix: slide-local text edits no
  // longer zip-sync `currentBuffer` on every commit. `reorderSlides` must fold
  // pending slide XML before permuting, or the edit vanishes when slides move.
  const { PptxMutationService } = await loadMutationModules();
  const engine = await loadEngine("features.pptx");
  const service = new PptxMutationService(engine);

  // Ensure a second slide exists so a reorder is meaningful (full-export path).
  await service.execute({ type: "add-slide", afterIndex: 0 });
  assert.ok(engine.slideCount >= 2, "add-slide should yield at least two slides");

  const marker = "CHEAP_COMMIT_MARKER_42";
  await service.execute({
    type: "update-paragraph-text",
    slideIndex: 0,
    shapeIndex: 0,
    paragraphIndex: 0,
    text: marker,
  });
  // Renderer model reflects the edit immediately.
  assert.equal(engine.getParagraphRunText(0, 0, 0), marker);

  // Move slide 0 to position 1. reorderSlides must sync pending first.
  await engine.reorderSlides([1, 0]);
  assert.equal(
    engine.getParagraphRunText(1, 0, 0),
    marker,
    "reorder must fold deferred slide-local edits into currentBuffer before permuting",
  );

  await assertExportRoundTrips("deferred slide-local commit + reorder", engine, { expectedSlides: 2 });
});

test("PptxMutationService uses slide-local commit for set-run-style-ranges", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const engine = {
    export: async () => {
      calls.push("export");
      return new ArrayBuffer(4);
    },
    getSlideXml: (slideIndex) => {
      calls.push(["getSlideXml", slideIndex]);
      return `<slide-${slideIndex}/>`;
    },
    restoreSlideXml: async (slideIndex, xml) => calls.push(["restoreSlideXml", slideIndex, xml]),
    restoreSnapshot: async () => calls.push("restoreSnapshot"),
    setRunStyleForRanges: async (slideIndex, shapeIndex, ranges, change) => {
      calls.push(["setRunStyleForRanges", slideIndex, shapeIndex, ranges, change]);
      return undefined;
    },
    commitMutation: async () => calls.push("commit"),
    commitSlideLocalMutation: async () => calls.push("commit-slide-local"),
  };
  const service = new PptxMutationService(engine);

  await service.execute({
    type: "set-run-style-ranges",
    slideIndex: 0,
    shapeIndex: 2,
    ranges: [{ paragraphIndex: 0, start: 0, end: 4 }],
    change: { fontSizePt: 18 },
  });

  assert.deepEqual(calls, [
    ["getSlideXml", 0],
    ["setRunStyleForRanges", 0, 2, [{ paragraphIndex: 0, start: 0, end: 4 }], { fontSizePt: 18 }],
    "commit-slide-local",
  ]);
  assert.ok(!calls.includes("export"), "font-size formatting must not export the full deck");
  assert.ok(!calls.includes("commit"), "font-size formatting must not run the full-export commit");
});

test("PptxMutationService keeps the full-export commit for non-slide-local commands", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const calls = [];
  const engine = {
    export: async () => {
      calls.push("export");
      return new ArrayBuffer(4);
    },
    getSlideXml: () => {
      calls.push("getSlideXml");
      return "<unused/>";
    },
    restoreSlideXml: async () => calls.push("restoreSlideXml"),
    restoreSnapshot: async () => calls.push("restoreSnapshot"),
    duplicateSlide: async (slideIndex) => {
      calls.push(["duplicateSlide", slideIndex]);
      return slideIndex + 1;
    },
    commitMutation: async () => calls.push("commit"),
    commitSlideLocalMutation: async () => calls.push("commit-slide-local"),
  };
  const service = new PptxMutationService(engine);

  const result = await service.execute({ type: "duplicate-slide", slideIndex: 1 });
  assert.equal(result, 2);

  assert.deepEqual(calls, [
    "export",
    ["duplicateSlide", 1],
    "commit",
  ]);
  assert.ok(!calls.includes("getSlideXml"), "structural commands must snapshot the whole package");
  assert.ok(
    !calls.includes("commit-slide-local"),
    "structural commands must not use the slide-local commit",
  );
});

test("PptxMutationService rejects commands before an engine is available", async () => {
  const { PptxMutationService } = await loadMutationModules();
  const service = new PptxMutationService(() => null);

  await assert.rejects(
    () => service.execute({ type: "insert-text-box", slideIndex: 0 }),
    /before the presentation engine loaded/,
  );
});
