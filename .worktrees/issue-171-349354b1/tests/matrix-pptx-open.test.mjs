import assert from "node:assert/strict";
import { test } from "node:test";
import {
  assertExportRoundTrips,
  assertSlideCount,
  loadEngine,
} from "./helpers/pptx-action-harness.mjs";
import { loadPowerPointPackageModule } from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const EDITABLE_DECKS = ["features.pptx", "simple-edit.pptx"];

test("modern editable decks load, render, and round-trip", async (t) => {
  for (const fixture of EDITABLE_DECKS) {
    await t.test(fixture, async () => {
      const engine = await loadEngine(fixture);

      assertSlideCount(engine, 1);
      await assertExportRoundTrips(fixture, engine);
    });
  }
});

test("unreadable malformed decks reject during engine load", async (t) => {
  for (const fixture of ["malformed-random.pptx", "malformed-truncated.pptx"]) {
    await t.test(fixture, async () => {
      await assert.rejects(
        loadEngine(fixture),
        /ZIP|Open XML|ppt\/presentation\.xml/i,
        `${fixture} should reject with a readable package error`,
      );
    });
  }
});

test("structurally malformed packages report validation errors cleanly", async (t) => {
  const { inspectPowerPointPackage, validatePowerPointPackageStructure } =
    await loadPowerPointPackageModule();

  for (const [fixture, expectedError] of [
    ["malformed-unsafe-path.pptx", /Unsafe ZIP entry paths: \.\.\/escape\.xml/],
    ["malformed-duplicate-entry.pptx", /Duplicate ZIP entries: ppt\/presentation\.xml/],
  ]) {
    await t.test(fixture, async () => {
      const inspection = inspectPowerPointPackage(toArrayBuffer(await readDeck(fixture)));
      const validation = validatePowerPointPackageStructure(inspection);

      assert.equal(validation.ok, false, `${fixture} must not validate as a PowerPoint package`);
      assert.ok(
        validation.errors.some((error) => expectedError.test(error)),
        `${fixture} should report its structural defect: ${validation.errors.join("; ")}`,
      );
    });
  }
});
