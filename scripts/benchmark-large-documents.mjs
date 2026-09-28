#!/usr/bin/env node
import { writeFile } from "node:fs/promises";
import {
  runLargeDocumentBenchmarks,
  DEFAULT_DOCX_TIERS,
  DEFAULT_PPTX_TIERS,
} from "./lib/large-document-benchmarks.mjs";

function parseArgs(args) {
  const options = {
    json: false,
    output: null,
    allTiers: false,
    docxOnly: false,
    pptxOnly: false,
    tiers: null,
  };

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === "--json") {
      options.json = true;
    } else if (arg === "--all-tiers") {
      options.allTiers = true;
    } else if (arg === "--docx-only") {
      options.docxOnly = true;
    } else if (arg === "--pptx-only") {
      options.pptxOnly = true;
    } else if (arg === "--output" && i + 1 < args.length) {
      options.output = args[++i];
    } else if (arg === "--tier" && i + 1 < args.length) {
      const parsed = args[++i].split(",").map((s) => parseInt(s.trim(), 10)).filter((n) => !isNaN(n));
      options.tiers = parsed;
    }
  }

  return options;
}

function formatTable(title, rows) {
  if (rows.length === 0) return "";
  const header = `=== ${title} ===\n` +
    `Tier | Size (KB) | 1st Vis (ms) | Usable (ms) | Cold Nav | Warm Nav | Typing | Search | Save | Bg Work | Heap Delta (KB)\n` +
    `-----+-----------+--------------+-------------+----------+----------+--------+--------+------+---------+----------------\n`;
  const body = rows.map((r) => {
    const sizeKb = (r.documentSizeBytes / 1024).toFixed(1).padStart(9);
    const firstVis = r.openToFirstVisibleMs.toFixed(1).padStart(12);
    const usable = r.openToUsableMs.toFixed(1).padStart(11);
    const cold = r.coldNavigationMs.toFixed(1).padStart(8);
    const warm = r.warmNavigationMs.toFixed(1).padStart(8);
    const typing = r.typingMutationMs.toFixed(1).padStart(6);
    const search = r.searchIndexMs.toFixed(1).padStart(6);
    const save = r.saveMs.toFixed(1).padStart(4);
    const bg = r.backgroundCompletionMs.toFixed(1).padStart(7);
    const heap = r.heapUsedDeltaBytes ? (r.heapUsedDeltaBytes / 1024).toFixed(0).padStart(15) : "N/A".padStart(15);
    return `${String(r.tier).padStart(4)} | ${sizeKb} | ${firstVis} | ${usable} | ${cold} | ${warm} | ${typing} | ${search} | ${save} | ${bg} | ${heap}`;
  }).join("\n");
  return `${header}${body}\n\n`;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));

  const scopes = [];
  if (options.docxOnly) scopes.push("docx");
  else if (options.pptxOnly) scopes.push("pptx");
  else scopes.push("docx", "pptx");

  const tierConfig = {};
  if (options.tiers && options.tiers.length > 0) {
    if (scopes.includes("docx")) tierConfig.docx = options.tiers;
    if (scopes.includes("pptx")) tierConfig.pptx = options.tiers;
  }

  const report = await runLargeDocumentBenchmarks({
    scopes,
    tiers: Object.keys(tierConfig).length > 0 ? tierConfig : undefined,
    includeOptInTiers: options.allTiers,
    onProgress: (msg) => {
      if (!options.json) process.stderr.write(`${msg}\n`);
    },
  });

  const jsonOutput = JSON.stringify(report, null, 2);

  if (options.output) {
    await writeFile(options.output, jsonOutput, "utf8");
    if (!options.json) console.log(`Report written to ${options.output}`);
  }

  if (options.json) {
    console.log(jsonOutput);
  } else {
    console.log(`\nLarge Document Scaling Benchmark Report`);
    console.log(`Platform: ${report.environment.platform} (${report.environment.arch}), Node: ${report.environment.nodeVersion}\n`);
    if (report.results.docx.length > 0) {
      console.log(formatTable("DOCX Benchmark Scaling (Pages)", report.results.docx));
    }
    if (report.results.pptx.length > 0) {
      console.log(formatTable("PPTX Benchmark Scaling (Slides)", report.results.pptx));
    }
  }
}

main().catch((err) => {
  console.error("Benchmark failed:", err);
  process.exit(1);
});
