import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

import {
  assertExportRoundTrips,
  loadEngine,
} from "./helpers/pptx-action-harness.mjs";
import {
  loadNativePowerPointViewModule,
  loadPowerPointPackageModule,
} from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

let coordinatorModulePromise;
function loadDocumentSaveCoordinator() {
  coordinatorModulePromise ??= (async () => {
    const outputDirectory = await mkdtemp(path.join(tmpdir(), "native-powerpoint-save-matrix-"));
    const outfile = path.join(outputDirectory, "document-save-coordinator.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/save/DocumentSaveCoordinator.ts")],
      bundle: true,
      format: "cjs",
      logLevel: "silent",
      outfile,
      platform: "node",
      target: "node22",
    });
    return require(outfile);
  })();
  return coordinatorModulePromise;
}

function createSaveHarness(NativePowerPointView, { autosaveEnabled = false } = {}) {
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

  view.file = file;
  view.loadedFile = file;
  return { created, file, modified, view };
}

test("matrix export remains structurally and content-valid for editable decks", async (t) => {
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents,
  } = await loadPowerPointPackageModule();

  for (const fixture of ["features.pptx", "simple-edit.pptx"]) {
    await t.test(fixture, async () => {
      const source = toArrayBuffer(await readDeck(fixture));
      const engine = await loadEngine(fixture);
      const exported = await assertExportRoundTrips(fixture, engine);

      assert.deepEqual(
        validatePowerPointExport(
          inspectPowerPointPackage(source),
          inspectPowerPointPackage(exported),
          engine.slideCount,
        ).errors,
        [],
        `${fixture}: export must preserve required package parts`,
      );
      assert.deepEqual(
        (await validatePowerPointExportContents(source, exported)).errors,
        [],
        `${fixture}: export must preserve required package bytes`,
      );
    });
  }
});

test("matrix coordinator serializes queued saves and retains the final dirty version", async () => {
  const { DocumentSaveCoordinator } = await loadDocumentSaveCoordinator();
  const writes = [];
  const events = [];
  let releaseFirstSerialization;
  const firstSerialization = new Promise((resolve) => {
    releaseFirstSerialization = resolve;
  });
  let serializations = 0;
  const coordinator = new DocumentSaveCoordinator({
    adapter: {
      async serialize(_context, request) {
        serializations += 1;
        events.push(`serialize:${request.targetVersion}`);
        if (serializations === 1) await firstSerialization;
        return `deck-${request.targetVersion}`;
      },
      async prepareForWrite(serialized) {
        return serialized;
      },
      async validate(prepared) {
        return prepared;
      },
      async persist(prepared) {
        writes.push(prepared);
      },
    },
    getContext: () => ({ deck: "fixture" }),
    autosave: {
      enabled: () => false,
      delayMs: () => 1,
      source: "autosave",
    },
    onStateChange: (state) => events.push(`state:${state}`),
  });

  coordinator.markDirty();
  const first = coordinator.save("manual");
  await Promise.resolve();
  assert.deepEqual(events, ["state:dirty", "state:saving", "serialize:1"]);

  coordinator.markDirty();
  const second = coordinator.save("manual");
  await Promise.resolve();
  assert.equal(serializations, 1, "a queued save must wait for the active serialization");

  releaseFirstSerialization();
  assert.deepEqual(await Promise.all([first, second]), [true, true]);
  assert.deepEqual(writes, ["deck-1", "deck-2"]);
  assert.equal(coordinator.version, 2);
  assert.equal(coordinator.state, "clean");
  await coordinator.waitForIdle();
});

test("matrix SaveController persists valid exports and preserves unvalidated recovery copies", async (t) => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const { inspectPowerPointPackage, validatePowerPointPackageStructure } =
    await loadPowerPointPackageModule();

  await t.test("validated save writes a reloadable package", async () => {
    const { file, modified, view } = createSaveHarness(NativePowerPointView);
    const source = toArrayBuffer(await readDeck("features.pptx"));
    view.engine = await loadEngine("features.pptx");
    view.sourceBuffer = source;
    view.sourcePackage = inspectPowerPointPackage(source);
    view.saveController.validateExportOverride = async (output) => inspectPowerPointPackage(output);
    view.saveController.isDirty = true;

    assert.equal(await view.saveCurrentPresentation(), true);
    assert.equal(modified.length, 1);
    assert.equal(modified[0].file, file);
    assert.equal(view.saveController.isDirty, false);
    assert.equal(
      validatePowerPointPackageStructure(
        inspectPowerPointPackage(modified[0].output),
        view.engine.slideCount,
      ).ok,
      true,
    );
  });

  await t.test("validation failure still writes a labeled recovery copy", async () => {
    const { created, view } = createSaveHarness(NativePowerPointView);
    const source = new Uint8Array([0]).buffer;
    view.engine = {
      slideCount: 1,
      async export() {
        return new Uint8Array([1]).buffer;
      },
    };
    view.sourceBuffer = source;
    view.sourcePackage = { hasVbaProject: false };
    view.saveController.validateExportOverride = async () => {
      throw new Error("simulated validation failure");
    };
    view.saveController.isDirty = true;

    assert.equal(
      await view.saveController.preserveUnsavedChangesForTeardown("closing the view"),
      true,
    );
    assert.equal(created.length, 1);
    assert.match(created[0].path, /Native PowerPoint unvalidated recovery/);
    assert.deepEqual(new Uint8Array(created[0].output), new Uint8Array([1]));
    assert.equal(view.saveController.isDirty, false);
    assert.equal(view.saveController.saveState, "recovered");
  });
});
