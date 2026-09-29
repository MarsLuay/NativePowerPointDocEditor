import { runSearchBenchmarks } from "./lib/search-benchmarks.mjs";

const report = await runSearchBenchmarks({
  onProgress: (message) => process.stderr.write(`${message}\n`),
});
process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
