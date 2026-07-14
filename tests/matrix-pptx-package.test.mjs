import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { test } from "node:test";
import {
  assertExportRoundTrips,
  assertSlideCount,
  exportBytes,
  loadEngine,
} from "./helpers/pptx-action-harness.mjs";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";

function sha256(bytes) {
  return createHash("sha256").update(new Uint8Array(bytes)).digest("hex");
}

async function mutateExportAndReload(label, engine) {
  const before = await exportBytes(engine);
  const expectedSlides = engine.slideCount + 1;

  await engine.addSlide(engine.slideCount - 1);
  assertSlideCount(engine, expectedSlides);

  const exported = await assertExportRoundTrips(label, engine, { expectedSlides });
  assert.notEqual(sha256(exported), sha256(before), `${label}: mutation must change exported package`);
}

test("package-authoritative mutation exports and reloads structurally", async () => {
  await mutateExportAndReload(
    "package-authoritative mutation",
    await loadEngine("features.pptx"),
  );
});

test("forced-JS package mutation exports and reloads structurally", async (t) => {
  const {
    PresentationEngine,
    resetForceJsBackendOverride,
    setForceJsBackendOverride,
  } = await loadPresentationEngineModule();

  if (typeof setForceJsBackendOverride !== "function") {
    t.skip("forced-JS override is unavailable");
    return;
  }

  setForceJsBackendOverride(true);
  try {
    const fixtureBytes = await (await loadEngine("features.pptx")).export();
    const engine = await PresentationEngine.load(
      fixtureBytes,
    );
    assert.equal(engine.getRendererBackend(), "js");
    await mutateExportAndReload("forced-JS package mutation", engine);
  } finally {
    resetForceJsBackendOverride?.();
  }
});
