import assert from "node:assert/strict";
import { test } from "node:test";
import {
  runGrammarBenchmarkScenario,
  runGrammarBenchmarks,
} from "../scripts/lib/grammar-benchmarks.mjs";

test("grammar benchmark compares disabled and enabled rapid edits without request buildup", async () => {
  const disabled = await runGrammarBenchmarkScenario({
    kind: "docx",
    units: 20,
    enabled: false,
    editCount: 8,
  });
  const enabled = await runGrammarBenchmarkScenario({
    kind: "docx",
    units: 20,
    enabled: true,
    editCount: 8,
    lintDelayMs: 1,
  });

  assert.equal(disabled.lintRequests, 0);
  assert.equal(disabled.completedResults, 0);
  assert.equal(disabled.staleOrCanceledResults, 8);
  assert.equal(enabled.lintRequests, 1);
  assert.equal(enabled.completedResults, 1);
  assert.equal(enabled.staleOrCanceledResults, 7);
  assert.equal(enabled.outstandingLintRequests, 1);
  assert.equal(enabled.logEntryCount, 2);
  assert.equal(enabled.slowDiagnosticsRetained, 0);
  assert.ok(enabled.typingLatency.maxMs >= 0);
});

test("grammar benchmark reports large DOCX/PPTX text scenarios and bounded slow summaries", async () => {
  const report = await runGrammarBenchmarks({ units: [5], editCount: 4, lintDelayMs: 1 });
  assert.equal(report.schemaVersion, 1);
  assert.equal(report.scenarios.length, 4);
  assert.deepEqual(
    report.scenarios.map((scenario) => [scenario.kind, scenario.enabled]),
    [["docx", false], ["docx", true], ["pptx", false], ["pptx", true]],
  );
  assert.ok(report.scenarios.every((scenario) => scenario.textLength > 100));
  assert.ok(report.scenarios.every((scenario) => scenario.frameImpact.maxGapMs >= 0));
});
