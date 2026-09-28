import assert from "node:assert/strict";
import { test } from "node:test";
import { createRequire } from "node:module";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);
let modulePromise;

async function loadProfiler() {
  modulePromise ??= bundleSource(
    "src/slideSwitchProfiler.ts",
    "slide-switch-profiler.cjs"
  ).then((outfile) => require(outfile));
  return modulePromise;
}

function start(profiler, fromSlide, toSlide, cacheState = "cold") {
  profiler.begin({
    fromSlide,
    toSlide,
    reason: "test-navigation",
    slideCount: 100,
    cacheState,
  });
}

test("slide switch profiler records warm/cold stage boundaries through the first frame", async () => {
  const { SlideSwitchProfiler } = await loadProfiler();
  let now = 100;
  const profiler = new SlideSwitchProfiler({ now: () => now, maxSamples: 10, maxWorstSwitches: 2 });

  start(profiler, 0, 1, "cold");
  profiler.recordStage("selection", 2);
  profiler.recordStage("cache-lookup", 1);
  profiler.recordStage("render", 20);
  profiler.recordStage("svg-parse", 3);
  profiler.recordStage("dom-swap", 2);
  now = 130;
  assert.equal(profiler.finish("stable"), null, "stable completion waits for a rendered frame");
  now = 136;
  const measurement = profiler.markFrame();

  assert.equal(measurement.distance, "adjacent");
  assert.equal(measurement.distanceSlides, 1);
  assert.equal(measurement.cacheState, "cold");
  assert.equal(measurement.inputToVisibleMs, 36);
  assert.equal(measurement.stageTimingsMs.render, 20);
  assert.equal(measurement.stageTimingsMs.domSwap, 2);
  assert.equal(measurement.status, "stable");
  assert.ok(measurement.frameSchedulingDelayMs >= 0);
});

test("slide switch profiler distinguishes distant navigation and retains only bounded worst switches", async () => {
  const { SlideSwitchProfiler } = await loadProfiler();
  let now = 0;
  const profiler = new SlideSwitchProfiler({ now: () => now, maxSamples: 5, maxWorstSwitches: 2 });

  for (const [toSlide, duration] of [[10, 10], [20, 40], [30, 20], [40, 80], [50, 60]]) {
    start(profiler, 0, toSlide, toSlide === 10 ? "warm" : "cold");
    now += duration;
    profiler.finish("failed");
  }

  const summary = profiler.getSummary();
  assert.equal(summary.sampleCount, 5);
  assert.equal(summary.retainedSampleCount, 5);
  assert.equal(summary.worstSwitches.length, 2);
  assert.equal(summary.worstSwitches[0].inputToVisibleMs, 80);
  assert.equal(summary.worstSwitches[1].inputToVisibleMs, 60);
  assert.equal(summary.byCacheAndDistance["warm-distant"].p95, 10);
  assert.equal(summary.timing.max, 80);
});

test("slide switch profiler exposes superseded switches without retaining an active trace", async () => {
  const { SlideSwitchProfiler } = await loadProfiler();
  let now = 20;
  const profiler = new SlideSwitchProfiler({ now: () => now });
  start(profiler, 4, 9, "cold");
  now = 28;
  const measurement = profiler.finish("superseded");
  assert.equal(measurement.status, "superseded");
  assert.equal(profiler.finish("stable"), null);
});
