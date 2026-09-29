import assert from "node:assert/strict";
import { test } from "node:test";
import {
  runDocxSaveExportBenchmarkTier,
  runPptxSaveExportBenchmarkTier,
  runSaveExportBenchmarks,
} from "../scripts/lib/save-export-benchmarks.mjs";

function assertOperation(operation) {
  assert.ok(operation.totalMs >= 0);
  assert.ok(operation.synchronousMs >= 0);
  assert.ok(operation.maxFrameGapMs >= 0);
  assert.ok(operation.outputBytes > 0);
}

test("DOCX save/export benchmark separates total and blocking measurements", async () => {
  const result = await runDocxSaveExportBenchmarkTier(3);
  assert.equal(result.format, "docx");
  assertOperation(result.save);
  assertOperation(result.pdfExport);
  assert.equal(result.complexity.pages, 3);
});

test("PPTX save/export benchmark uses real package and slide render paths", async () => {
  const result = await runPptxSaveExportBenchmarkTier(3);
  assert.equal(result.format, "pptx");
  assertOperation(result.save);
  assertOperation(result.pdfExport);
  assert.equal(result.complexity.slides, 3);
});

test("save/export benchmark runner preserves aligned complexity tiers", async () => {
  const report = await runSaveExportBenchmarks({ tiers: [2] });
  assert.equal(report.schemaVersion, 1);
  assert.deepEqual(report.results.docx.map((result) => result.tier), [2]);
  assert.deepEqual(report.results.pptx.map((result) => result.tier), [2]);
});
