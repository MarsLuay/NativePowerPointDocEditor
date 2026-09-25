import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);
let diagnosticsModulePromise;
async function loadDiagnosticsModule() {
  diagnosticsModulePromise ??= bundleSource(
    "src/powerpoint/selectionOverlayDiagnostics.ts",
    "selection-overlay-diagnostics.cjs",
  ).then((outfile) => require(outfile));
  return diagnosticsModulePromise;
}

test("selection overlay diagnostics identify materially inflated group bounds", async () => {
  const { getSelectionOverlayBoundsAnomaly } = await loadDiagnosticsModule();

  const anomaly = getSelectionOverlayBoundsAnomaly(
    { left: 96, top: 44, width: 452, height: 134 },
    { left: 100, top: 48, width: 216, height: 64 },
  );

  assert.equal(anomaly?.materiallyDifferent, true);
  assert.equal(anomaly?.widthDelta, 236);
  assert.equal(anomaly?.heightDelta, 70);
  assert.ok((anomaly?.widthRatio ?? 0) > 2);
});

test("selection overlay diagnostics report only bounded descendants outside the frame", async () => {
  const { summarizeSelectionOverlayContributors } = await loadDiagnosticsModule();
  const frame = { left: 100, top: 50, width: 200, height: 100 };
  const candidate = (identity, localRect, uiKind = "other") => ({
    tagName: "rect",
    identity,
    clientRect: localRect,
    localRect,
    transform: null,
    visibility: "visible",
    display: "block",
    uiKind,
  });

  const summary = summarizeSelectionOverlayContributors([
    candidate("rect[data-ooxml-run-idx=1]", { left: 110, top: 60, width: 20, height: 20 }),
    candidate("rect.selection", { left: 90, top: 50, width: 230, height: 100 }, "selection"),
    candidate("line.caret", { left: 150, top: 70, width: 1, height: 20 }, "caret"),
  ], frame, 2);

  assert.equal(summary.descendantCount, 3);
  assert.equal(summary.outlierCount, 1);
  assert.equal(summary.contributors.length, 1);
  assert.equal(summary.contributors[0].identity, "rect.selection");
  assert.equal(summary.contributors[0].uiKind, "selection");
  assert.equal(summary.contributors[0].extendsBeyondFrame, true);
});
