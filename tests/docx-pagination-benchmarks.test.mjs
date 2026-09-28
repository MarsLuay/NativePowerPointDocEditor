import assert from "node:assert/strict";
import { test } from "node:test";
import {
  EDIT_KINDS,
  LOCATIONS,
  compareDocxPaginationBenchmarks,
  runDocxPaginationBenchmark,
} from "../scripts/lib/docx-pagination-benchmarks.mjs";

test("DOCX pagination benchmark covers every edit kind at top, middle, and end", async () => {
  const report = await runDocxPaginationBenchmark({ tiers: [10], iterations: 1 });
  assert.equal(report.schemaVersion, 1);
  assert.equal(report.fixture, "deterministic-docx-pages-v1");
  assert.deepEqual(report.tiers.map((tier) => tier.pageCount), [10]);

  const scenarios = report.tiers[0].scenarios;
  assert.equal(scenarios.length, EDIT_KINDS.length * LOCATIONS.length);
  for (const editKind of EDIT_KINDS) {
    for (const location of LOCATIONS) {
      const scenario = scenarios.find((entry) => entry.editKind === editKind && entry.location === location);
      assert.ok(scenario, `${editKind} at ${location} should be measured`);
      assert.equal(scenario.iterations, 1);
      assert.ok(scenario.affectedRange.documentSize > 0);
      assert.ok(scenario.timing.p50 >= 0);
      assert.ok(scenario.timing.p95 >= scenario.timing.p50);
      assert.ok(scenario.timing.max >= scenario.timing.p95);
      assert.ok(scenario.paginationPasses.p50 >= 1);
      assert.ok(scenario.layoutCallbackCount.p50 >= 1);
      assert.ok(scenario.pagesRecalculated.p50 >= 1);
      assert.equal(scenario.initialPageCount, 10);
      assert.ok(scenario.stablePageCount >= 1);
    }
  }

  const locations = new Map(scenarios
    .filter((scenario) => scenario.editKind === "paragraph-insert")
    .map((scenario) => [scenario.location, scenario.affectedRange.pageStart]));
  assert.equal(locations.get("top"), 1);
  assert.equal(locations.get("middle"), 6);
  assert.equal(locations.get("end"), 10);
});

test("DOCX pagination benchmark scales across 10, 50, 100, and 250 page fixtures", async () => {
  const report = await runDocxPaginationBenchmark({ tiers: [10, 50, 100, 250], iterations: 1 });
  assert.deepEqual(report.tiers.map((tier) => tier.pageCount), [10, 50, 100, 250]);
  for (const tier of report.tiers) {
    assert.equal(tier.scenarios.length, EDIT_KINDS.length * LOCATIONS.length);
    const top = tier.scenarios.find((scenario) => scenario.editKind === "paragraph-insert" && scenario.location === "top");
    const end = tier.scenarios.find((scenario) => scenario.editKind === "paragraph-insert" && scenario.location === "end");
    assert.equal(top.affectedRange.pageStart, 1);
    assert.equal(end.affectedRange.pageStart, tier.pageCount);
    assert.ok(end.pagesRecalculated.p50 <= tier.pageCount);
  }
});

test("DOCX pagination benchmark comparisons expose before/after p95 evidence", async () => {
  const before = await runDocxPaginationBenchmark({ tiers: [10], iterations: 1 });
  const after = await runDocxPaginationBenchmark({ tiers: [10], iterations: 1 });
  const comparison = compareDocxPaginationBenchmarks(before, after);
  assert.equal(comparison.schemaVersion, 1);
  assert.equal(comparison.comparisons.length, EDIT_KINDS.length * LOCATIONS.length);
  assert.ok(comparison.comparisons.every((entry) => typeof entry.beforeP95Ms === "number" && typeof entry.afterP95Ms === "number"));
});
