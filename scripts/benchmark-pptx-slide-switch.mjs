#!/usr/bin/env node
import { writeFile } from "node:fs/promises";
import { runPptxSlideSwitchBenchmarks } from "./lib/slide-switch-benchmarks.mjs";

const args = process.argv.slice(2);
const quick = args.includes("--quick");
const outputIndex = args.indexOf("--output");
const outputPath = outputIndex >= 0 ? args[outputIndex + 1] : null;
const iterationsIndex = args.indexOf("--iterations");
const requestedIterations = iterationsIndex >= 0 ? Number(args[iterationsIndex + 1]) : undefined;
const report = await runPptxSlideSwitchBenchmarks({
  tiers: quick ? [10, 50] : undefined,
  iterations: requestedIterations,
  onProgress: outputPath ? undefined : (message) => process.stderr.write(`${message}\n`),
});
const serialized = `${JSON.stringify(report, null, 2)}\n`;
if (outputPath) await writeFile(outputPath, serialized, "utf8");
else process.stdout.write(serialized);
