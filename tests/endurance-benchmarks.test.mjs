import assert from "node:assert/strict";
import { test } from "node:test";
import {
  ENDURANCE_OPERATIONS,
  createDeterministicWorkload,
  runEnduranceHarness,
} from "../scripts/lib/endurance-benchmarks.mjs";

test("endurance workload is deterministic and emits periodic summaries", async () => {
  const workloadA = createDeterministicWorkload({ seed: 42, actionCount: 40 });
  const workloadB = createDeterministicWorkload({ seed: 42, actionCount: 40 });
  assert.deepEqual(workloadA, workloadB);
  assert.ok(workloadA.every((action) => ENDURANCE_OPERATIONS.includes(action.kind)));

  const report = await runEnduranceHarness({ seed: 42, actionCount: 40, summaryEveryActions: 10 });
  assert.equal(report.schemaVersion, 1);
  assert.equal(report.actionCount, 40);
  assert.equal(report.summaries.length, 4);
  assert.equal(report.summaries.at(-1).summary.actions, 40);
  assert.ok(Number.isFinite(report.comparison.typingP95DeltaMs));
  assert.ok(Number.isFinite(report.comparison.frameMaxDeltaMs));
});

test("endurance accounting includes custom resource growth for regression comparison", async () => {
  const report = await runEnduranceHarness({
    actionCount: 12,
    summaryEveryActions: 4,
    workload: createDeterministicWorkload({ seed: 7, actionCount: 12 }),
    measureAction: async () => ({
      typingMs: 2,
      frameGapMs: 3,
      saveMs: 4,
      resources: { domNodes: 2, mutationObservers: 1, timers: 1, cacheEntries: 1, editorViews: 0 },
    }),
  });
  assert.equal(report.comparison.domNodeGrowth, 16);
  assert.equal(report.comparison.observerGrowth, 8);
  assert.equal(report.comparison.timerGrowth, 8);
  assert.equal(report.comparison.cacheEntryGrowth, 8);
  assert.equal(report.comparison.editorViewGrowth, 0);
});
