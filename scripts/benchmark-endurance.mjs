import { runEnduranceHarness } from "./lib/endurance-benchmarks.mjs";

const report = await runEnduranceHarness({
  actionCount: 60_000,
  summaryEveryActions: 1_000,
});
process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
