import { runPptxThumbnailBenchmarks } from "./lib/thumbnail-render-benchmarks.mjs";

const report = await runPptxThumbnailBenchmarks({
  onProgress: (message) => process.stderr.write(`${message}\n`),
});
process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
