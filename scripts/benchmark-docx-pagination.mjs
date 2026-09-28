#!/usr/bin/env node
import { writeFile } from "node:fs/promises";
import { runDocxPaginationBenchmark } from "./lib/docx-pagination-benchmarks.mjs";

const args = process.argv.slice(2);
const outputIndex = args.indexOf("--output");
const outputPath = outputIndex >= 0 ? args[outputIndex + 1] : null;
const quick = args.includes("--quick");
const iterationsIndex = args.indexOf("--iterations");
const requestedIterations = iterationsIndex >= 0 ? Number(args[iterationsIndex + 1]) : undefined;
const tiers = quick ? [10, 50] : undefined;

const report = await runDocxPaginationBenchmark({
  tiers,
  iterations: requestedIterations,
  onProgress: outputPath ? undefined : (message) => process.stderr.write(`${message}\n`),
});
const serialized = `${JSON.stringify(report, null, 2)}\n`;
if (outputPath) {
  await writeFile(outputPath, serialized, "utf8");
} else {
  process.stdout.write(serialized);
}
