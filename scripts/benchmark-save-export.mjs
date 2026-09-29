import { runSaveExportBenchmarks } from "./lib/save-export-benchmarks.mjs";

const report = await runSaveExportBenchmarks({
  onProgress: (message) => process.stderr.write(`${message}\n`),
});
process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
