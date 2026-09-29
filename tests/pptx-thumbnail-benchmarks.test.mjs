import assert from "node:assert/strict";
import { test } from "node:test";
import {
  runPptxThumbnailBenchmarkTier,
  runPptxThumbnailBenchmarks,
} from "../scripts/lib/thumbnail-render-benchmarks.mjs";

test("PPTX thumbnail benchmark measures active, nearby, and background render boundaries", async () => {
  const result = await runPptxThumbnailBenchmarkTier(9);

  assert.equal(result.fixture, "deterministic-pptx-thumbnails-v1");
  assert.equal(result.tier, 9);
  assert.equal(result.activeSlideIndex, 4);
  assert.deepEqual(result.priorityIndices, [2, 3, 4, 5, 6]);
  assert.deepEqual(result.nearbyIndices, [2, 3, 5, 6]);
  assert.equal(result.backgroundCount, 4);
  assert.ok(result.activeSlideMs >= 0);
  assert.ok(result.nearbyCompletionMs >= 0);
  assert.ok(result.backgroundCompletionMs >= 0);
  assert.equal(result.perThumbnail.count, 9);
  assert.equal(result.cache.misses, 9);
  assert.equal(result.cache.hits, 2);
  assert.equal(result.cache.renderedSvgCount, 9);
  assert.equal(result.cache.entries, 9);
  assert.equal(result.concurrency.configuredBatchSize, 1);
  assert.equal(result.concurrency.observedMax, 1);
  assert.ok(Number.isFinite(result.interactionImpact.maxEventLoopGapMs));
  assert.ok(result.interactionImpact.eventLoopGapCount >= 1);

  const coldIndices = result.measurements
    .filter((measurement) => !measurement.cacheHit)
    .map((measurement) => measurement.index);
  assert.deepEqual(coldIndices, [4, 2, 3, 5, 6, 0, 1, 7, 8]);
  assert.deepEqual(
    result.measurements.filter((measurement) => measurement.cacheHit).map((measurement) => measurement.index),
    [4, 2],
  );
});

test("PPTX thumbnail benchmark runner reports requested tiers", async () => {
  const report = await runPptxThumbnailBenchmarks({ tiers: [5, 7] });
  assert.equal(report.schemaVersion, 1);
  assert.equal(report.results.length, 2);
  assert.deepEqual(report.results.map((result) => result.tier), [5, 7]);
  assert.ok(report.results.every((result) => result.cache.renderedSvgCount === result.tier));
});
