import { runGrammarBenchmarks } from "./lib/grammar-benchmarks.mjs";

const report = await runGrammarBenchmarks({
  onProgress: (message) => process.stderr.write(`${message}\n`),
});
process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
