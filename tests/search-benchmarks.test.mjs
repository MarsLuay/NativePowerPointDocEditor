import assert from "node:assert/strict";
import { test } from "node:test";
import {
  runDocxSearchBenchmarkTier,
  runPptxSearchBenchmarkTier,
  runSearchBenchmarks,
} from "../scripts/lib/search-benchmarks.mjs";

test("DOCX search benchmark separates index, query, navigation, replace, and reveal timing", async () => {
  const result = await runDocxSearchBenchmarkTier(5);
  assert.equal(result.format, "docx");
  assert.equal(result.tier, 5);
  assert.ok(result.indexedRowCount > 0);
  assert.ok(result.resultCount > 0);
  assert.ok(result.indexBuildMs >= 0);
  assert.ok(result.firstQueryMs >= 0);
  assert.ok(result.repeatedQuery.p95Ms >= result.repeatedQuery.p50Ms);
  assert.ok(result.navigationMs >= 0);
  assert.ok(result.replaceMs >= 0);
  assert.ok(result.reveal.synchronousMs >= 0);
  assert.ok(result.reveal.scheduledDelayMs >= result.reveal.synchronousMs);
});

test("PPTX search benchmark uses the real PresentationEngine index and reveal path", async () => {
  const result = await runPptxSearchBenchmarkTier(5);
  assert.equal(result.format, "pptx");
  assert.equal(result.tier, 5);
  assert.equal(result.indexedRowCount, 5);
  assert.ok(result.resultCount > 0);
  assert.ok(result.indexBuildMs >= 0);
  assert.ok(result.firstQueryMs >= 0);
  assert.ok(result.repeatedQuery.p95Ms >= result.repeatedQuery.p50Ms);
  assert.ok(result.replaceEngineMs >= 0);
  assert.ok(result.reveal.synchronousMs >= 0);
  assert.ok(result.reveal.scheduledDelayMs >= result.reveal.synchronousMs);
});

test("search benchmark runner keeps DOCX and PPTX tiers aligned", async () => {
  const report = await runSearchBenchmarks({ tiers: [3, 5] });
  assert.equal(report.schemaVersion, 1);
  assert.deepEqual(report.results.docx.map((result) => result.tier), [3, 5]);
  assert.deepEqual(report.results.pptx.map((result) => result.tier), [3, 5]);
});
