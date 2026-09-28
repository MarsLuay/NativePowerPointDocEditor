import assert from "node:assert/strict";
import { test } from "node:test";
import {
  CACHE_STATES,
  DISTANCES,
  comparePptxSlideSwitchBenchmarks,
  runPptxSlideSwitchBenchmarks,
} from "../scripts/lib/slide-switch-benchmarks.mjs";

test("PPTX slide-switch benchmark separates warm/cold and adjacent/distant navigation", async () => {
  const report = await runPptxSlideSwitchBenchmarks({ tiers: [10], iterations: 1 });
  assert.equal(report.schemaVersion, 1);
  assert.equal(report.fixture, "deterministic-pptx-slides-v1");
  const scenarios = report.tiers[0].scenarios;
  assert.equal(scenarios.length, CACHE_STATES.length * DISTANCES.length);
  for (const cacheState of CACHE_STATES) {
    for (const distance of DISTANCES) {
      const scenario = scenarios.find((entry) => entry.cacheState === cacheState && entry.distance === distance);
      assert.ok(scenario, `${cacheState} ${distance} scenario should be measured`);
      assert.equal(scenario.iterations, 1);
      assert.ok(scenario.timing.p95 >= scenario.timing.p50);
      assert.ok(scenario.timing.max >= scenario.timing.p95);
      assert.ok(scenario.stageTimingsMs.render.p50 >= 0);
      assert.ok(scenario.stageTimingsMs.svgParse.p50 >= 0);
      assert.ok(scenario.stageTimingsMs.frameSchedulingDelay.p50 >= 0);
      assert.ok(scenario.complexity.svgCharacters > 0);
      assert.ok(scenario.complexity.shapeCount >= 0);
    }
  }
  assert.equal(report.tiers[0].worstSwitches.length, 4);
});

test("PPTX slide-switch benchmark covers representative large deck sizes", async () => {
  const report = await runPptxSlideSwitchBenchmarks({ tiers: [10, 50, 100, 250, 500], iterations: 1 });
  assert.deepEqual(report.tiers.map((tier) => tier.slideCount), [10, 50, 100, 250, 500]);
  for (const tier of report.tiers) {
    assert.equal(tier.scenarios.length, 4);
    assert.ok(tier.worstSwitches.length <= 10);
    assert.ok(tier.scenarios.every((scenario) => scenario.samples[0].complexity.slideXmlCharacters > 0));
  }
});

test("PPTX slide-switch comparisons expose before/after p95 and max evidence", async () => {
  const before = await runPptxSlideSwitchBenchmarks({ tiers: [10], iterations: 1 });
  const after = await runPptxSlideSwitchBenchmarks({ tiers: [10], iterations: 1 });
  const comparison = comparePptxSlideSwitchBenchmarks(before, after);
  assert.equal(comparison.schemaVersion, 1);
  assert.equal(comparison.comparisons.length, 4);
  assert.ok(comparison.comparisons.every((entry) =>
    typeof entry.beforeP95Ms === "number"
      && typeof entry.afterP95Ms === "number"
      && typeof entry.beforeMaxMs === "number"
      && typeof entry.afterMaxMs === "number"
  ));
});
