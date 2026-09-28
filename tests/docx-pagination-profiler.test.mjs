import assert from "node:assert/strict";
import { test } from "node:test";
import { createRequire } from "node:module";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);
let modulePromise;

async function loadProfiler() {
  modulePromise ??= bundleSource(
    "src/docxPaginationProfiler.ts",
    "docx-pagination-profiler.cjs"
  ).then((outfile) => require(outfile));
  return modulePromise;
}

test("pagination profiler records transaction-to-stable boundaries and bounded timing stats", async () => {
  const {
    DocxPaginationProfiler,
    createAffectedDocxRange,
    inferDocxReflowLocation,
  } = await loadProfiler();
  let now = 100;
  const profiler = new DocxPaginationProfiler({ maxSamples: 2, now: () => now });
  const range = createAffectedDocxRange(10, 20, 100, 10);

  assert.equal(inferDocxReflowLocation(1, 100), "top");
  assert.equal(inferDocxReflowLocation(50, 100), "middle");
  assert.equal(inferDocxReflowLocation(99, 100), "end");
  assert.deepEqual(range, {
    from: 10,
    to: 20,
    documentSize: 100,
    pageStart: 2,
    pageEnd: 3,
  });

  profiler.beginEdit({
    editKind: "paragraph-insert",
    location: "top",
    affectedRange: range,
    initialPageCount: 10,
  });
  now = 110;
  profiler.markSynchronousWorkComplete();
  profiler.recordPaginationPass(10);
  profiler.recordPaginationPass(12);
  profiler.recordLayoutCallback();
  now = 125;
  const measurement = profiler.complete({ stablePageCount: 12 });

  assert.equal(measurement.totalMs, 25);
  assert.equal(measurement.synchronousWorkMs, 10);
  assert.equal(measurement.frameSchedulingDelayMs, 15);
  assert.equal(measurement.paginationPasses, 2);
  assert.equal(measurement.layoutCallbackCount, 1);
  assert.equal(measurement.pagesRecalculated, 2);
  assert.equal(measurement.stablePageCount, 12);
  assert.equal(measurement.status, "stable");

  const summary = profiler.getSummary();
  assert.equal(summary.sampleCount, 1);
  assert.equal(summary.retainedSampleCount, 1);
  assert.equal(summary.timing.p50, 25);
  assert.equal(summary.timing.p95, 25);
  assert.equal(summary.timing.max, 25);
  assert.equal(summary.recentMeasurements.length, 1);
});

test("pagination profiler caps retained samples and never keeps an active transaction after completion", async () => {
  const { DocxPaginationProfiler } = await loadProfiler();
  let now = 0;
  const profiler = new DocxPaginationProfiler({ maxSamples: 2, now: () => now });
  const base = {
    editKind: "transaction",
    location: "middle",
    affectedRange: { from: 1, to: 2, documentSize: 10, pageStart: 1, pageEnd: 1 },
  };

  for (const duration of [5, 10, 20]) {
    profiler.beginEdit(base);
    now += duration;
    profiler.complete();
  }

  const summary = profiler.getSummary();
  assert.equal(summary.sampleCount, 3);
  assert.equal(summary.retainedSampleCount, 2);
  assert.deepEqual(summary.recentMeasurements.map((sample) => sample.totalMs), [10, 20]);
  assert.equal(profiler.complete(), null);
});

test("pagination profiler reports timeouts and cancels session state", async () => {
  const { DocxPaginationProfiler } = await loadProfiler();
  let now = 50;
  const profiler = new DocxPaginationProfiler({ now: () => now });
  profiler.beginEdit({
    editKind: "page-break-change",
    location: "end",
    affectedRange: { from: 8, to: 8, documentSize: 8, pageStart: 2, pageEnd: 2 },
  });
  now = 80;
  const timeout = profiler.complete({ status: "timeout" });
  assert.equal(timeout.status, "timeout");
  profiler.beginEdit({
    editKind: "image-resize",
    location: "middle",
    affectedRange: { from: 3, to: 4, documentSize: 8, pageStart: 1, pageEnd: 2 },
  });
  profiler.cancel();
  assert.equal(profiler.complete(), null);
});
