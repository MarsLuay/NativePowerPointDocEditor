import assert from "node:assert/strict";
import { test } from "node:test";
import {
  loadPowerPointPackageModule,
  loadPresentationEngineModule,
} from "./helpers/load-plugin-modules.mjs";
import { createRenderer, readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

// Audit coverage: "PPTX · Save / recovery / export".
//
// This exercises the headless-feasible core of the save/export pipeline that
// NativePowerPointView.saveCurrentPresentation() and writeRecoveryCopy() rely
// on through validateExportBeforeSave():
//   1. engine.export()                       (PresentationEngine.exportRendererState)
//   2. validatePowerPointExport(...)         (package structure + preserved parts)
//   3. validatePowerPointExportContents(...) (byte-level preserved parts)
//   4. PresentationEngine.validateRoundTrip  (slide count + slide renders)
//
// PNG/PDF/zip rasterization (PowerPointExport.ts, exportController.ts) and the
// print path are DOM <canvas> bound (CanvasRenderingContext2D / jsPDF / Blob),
// so they cannot run under node:test and remain manual-only.

const FIXTURE = "features.pptx";
const EXPECTED_SLIDE_COUNT = 1;

test("save/export round trip preserves slides and passes export validation", async (t) => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents,
  } = await loadPowerPointPackageModule();

  const input = await readDeck(FIXTURE);
  const sourceBuffer = toArrayBuffer(input);
  const original = inspectPowerPointPackage(sourceBuffer);

  // Mirror saveCurrentPresentation(): export the live engine state.
  const engine = await PresentationEngine.load(sourceBuffer);
  assert.equal(engine.slideCount, EXPECTED_SLIDE_COUNT);
  const output = await engine.export();

  await t.test("validatePowerPointExport finds no errors", () => {
    const exported = inspectPowerPointPackage(output);
    const result = validatePowerPointExport(original, exported, EXPECTED_SLIDE_COUNT);
    assert.deepEqual(result.errors, []);
    assert.equal(result.ok, true);
  });

  await t.test("validatePowerPointExportContents finds no errors", async () => {
    const result = await validatePowerPointExportContents(sourceBuffer, output);
    assert.deepEqual(result.errors, []);
    assert.equal(result.ok, true);
  });

  await t.test("validateRoundTrip accepts the exported buffer", async () => {
    await assert.doesNotReject(
      PresentationEngine.validateRoundTrip(output, EXPECTED_SLIDE_COUNT)
    );
  });

  await t.test("reloaded export preserves slide count and renders", async () => {
    const reloaded = await createRenderer(new Uint8Array(output));
    assert.equal(reloaded.getSlideCount(), EXPECTED_SLIDE_COUNT);
    assert.match(reloaded.renderSlideSvg(0), /^<svg\b/);
  });
});

// PNG/PDF/print remain DOM-canvas bound and are intentionally not covered here.
test("PNG/PDF/print rasterization is DOM-canvas bound (manual-only)", () => {
  assert.ok(true, "exportSlideToPng/exportSlidesToPdf/printPresentation need a DOM canvas");
});
