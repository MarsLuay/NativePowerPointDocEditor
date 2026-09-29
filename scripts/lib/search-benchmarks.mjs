import { performance } from "node:perf_hooks";
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import { createDeck } from "../../tests/helpers/fixture-builder.mjs";
import { generateDeterministicDocx } from "./large-document-benchmarks.mjs";
import { loadPresentationEngineModule } from "../../tests/helpers/load-plugin-modules.mjs";


export const DEFAULT_SEARCH_TIERS = [10, 25, 50];

function round(value) {
  return Math.round(value * 10) / 10;
}

function percentile(values, percentileValue) {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.min(sorted.length - 1, Math.floor((sorted.length - 1) * percentileValue))] ?? 0;
}

function latencySummary(values) {
  return {
    count: values.length,
    p50Ms: round(percentile(values, 0.5)),
    p95Ms: round(percentile(values, 0.95)),
    maxMs: round(Math.max(0, ...values)),
  };
}

function nextImmediate() {
  return new Promise((resolve) => setImmediate(resolve));
}

function queryRows(rows, query) {
  const normalized = query.toLocaleLowerCase();
  return rows.flatMap((row, rowIndex) => {
    const text = row.text.toLocaleLowerCase();
    const matches = [];
    let offset = text.indexOf(normalized);
    while (offset >= 0) {
      matches.push({ rowIndex, offset });
      offset = text.indexOf(normalized, offset + Math.max(1, normalized.length));
    }
    return matches;
  });
}

function measureQuery(rows, query, count = 1) {
  const timings = [];
  let matches = [];
  for (let iteration = 0; iteration < count; iteration += 1) {
    const started = performance.now();
    matches = queryRows(rows, query);
    timings.push(performance.now() - started);
  }
  return { matches, timings };
}

async function measureReveal(reveal) {
  const started = performance.now();
  const value = reveal();
  const synchronousMs = performance.now() - started;
  await nextImmediate();
  return {
    value,
    synchronousMs: round(synchronousMs),
    scheduledDelayMs: round(performance.now() - started),
  };
}

function benchmarkRows(rows, query, replace) {
  const indexStarted = performance.now();
  const indexedRows = rows.map((row) => ({ ...row, text: String(row.text) }));
  const indexBuildMs = round(performance.now() - indexStarted);

  const first = measureQuery(indexedRows, query);
  const repeated = measureQuery(indexedRows, query, 5);
  const navigationStarted = performance.now();
  let current = 0;
  for (let iteration = 0; iteration < Math.min(100, Math.max(1, first.matches.length * 2)); iteration += 1) {
    current = first.matches.length === 0 ? 0 : (current + 1) % first.matches.length;
  }
  const navigationMs = performance.now() - navigationStarted;

  const replaceStarted = performance.now();
  const replacedRows = indexedRows.map((row) => ({
    ...row,
    text: row.text.split(query).join(replace),
  }));
  const replaceMs = round(performance.now() - replaceStarted);

  return {
    indexedRows,
    replacedRows,
    resultCount: first.matches.length,
    indexBuildMs,
    firstQueryMs: round(first.timings[0] ?? 0),
    repeatedQuery: latencySummary(repeated.timings),
    navigationMs: round(navigationMs),
    replaceMs,
    currentMatch: first.matches[current] ?? null,
  };
}

export async function runDocxSearchBenchmarkTier(pageCount, options = {}) {
  const documentBuffer = await generateDeterministicDocx(pageCount);
  const zip = await JSZip.loadAsync(documentBuffer.slice(0));
  const xml = await zip.file("word/document.xml").async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const paragraphs = Array.from(doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p"));
  const rows = paragraphs.map((paragraph, index) => ({
    index,
    text: paragraph.textContent ?? "",
  }));
  const query = options.query ?? "quick brown fox";
  const replace = options.replace ?? "quick green fox";
  const measured = benchmarkRows(rows, query, replace);
  const reveal = await measureReveal(() => measured.currentMatch ? rows[measured.currentMatch.rowIndex]?.text ?? "" : "");

  return {
    schemaVersion: 1,
    format: "docx",
    tier: pageCount,
    unit: "pages",
    fixture: "deterministic-docx-search-v1",
    documentSizeBytes: documentBuffer.byteLength,
    indexedRowCount: measured.indexedRows.length,
    resultCount: measured.resultCount,
    indexBuildMs: measured.indexBuildMs,
    firstQueryMs: measured.firstQueryMs,
    repeatedQuery: measured.repeatedQuery,
    navigationMs: measured.navigationMs,
    replaceMs: measured.replaceMs,
    reveal: {
      synchronousMs: reveal.synchronousMs,
      scheduledDelayMs: reveal.scheduledDelayMs,
      selectedRow: measured.currentMatch?.rowIndex ?? null,
    },
  };
}

export async function runPptxSearchBenchmarkTier(slideCount, options = {}) {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const deck = createDeck({ format: "pptx", slideCount, richFirstSlide: true });
  const buffer = deck.buffer.slice(deck.byteOffset, deck.byteOffset + deck.byteLength);
  const engine = await PresentationEngine.load(buffer.slice(0));
  const query = options.query ?? "Slide";
  const replace = options.replace ?? "Page";

  const indexStarted = performance.now();
  const rows = [];
  for (let index = 0; index < engine.slideCount; index += 1) {
    rows.push({ index, text: engine.getSlideXml(index) });
  }
  const indexBuildMs = round(performance.now() - indexStarted);
  const measured = benchmarkRows(rows, query, replace);
  const selectedSlide = measured.currentMatch?.rowIndex ?? 0;
  const reveal = await measureReveal(() => engine.renderSlide(selectedSlide).svg);

  const replaceStarted = performance.now();
  const replacedCount = await engine.replaceText(query, replace);
  const replaceEngineMs = round(performance.now() - replaceStarted);

  return {
    schemaVersion: 1,
    format: "pptx",
    tier: slideCount,
    unit: "slides",
    fixture: "deterministic-pptx-search-v1",
    documentSizeBytes: buffer.byteLength,
    indexedRowCount: rows.length,
    resultCount: measured.resultCount,
    indexBuildMs,
    firstQueryMs: measured.firstQueryMs,
    repeatedQuery: measured.repeatedQuery,
    navigationMs: measured.navigationMs,
    replaceMs: measured.replaceMs,
    replaceEngineMs,
    replacedCount,
    reveal: {
      synchronousMs: reveal.synchronousMs,
      scheduledDelayMs: reveal.scheduledDelayMs,
      selectedSlide,
    },
  };
}

export async function runSearchBenchmarks(options = {}) {
  const tiers = options.tiers ?? DEFAULT_SEARCH_TIERS;
  const results = { docx: [], pptx: [] };
  for (const tier of tiers) {
    options.onProgress?.(`Running DOCX search tier ${tier} pages...`);
    results.docx.push(await runDocxSearchBenchmarkTier(tier, options));
    options.onProgress?.(`Running PPTX search tier ${tier} slides...`);
    results.pptx.push(await runPptxSearchBenchmarkTier(tier, options));
  }
  return {
    schemaVersion: 1,
    timestamp: new Date().toISOString(),
    environment: {
      platform: process.platform,
      arch: process.arch,
      nodeVersion: process.version,
    },
    results,
  };
}

