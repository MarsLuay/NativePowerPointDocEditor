import { performance } from "node:perf_hooks";
import JSZip from "jszip";
import { createDeck } from "../../tests/helpers/fixture-builder.mjs";
import { generateDeterministicDocx } from "./large-document-benchmarks.mjs";
import { loadPresentationEngineModule } from "../../tests/helpers/load-plugin-modules.mjs";

export const DEFAULT_SAVE_EXPORT_TIERS = [10, 50, 100];

function round(value) { return Math.round(value * 10) / 10; }
function nextImmediate() { return new Promise((resolve) => setImmediate(resolve)); }

async function measureOperation(operation) {
  const gaps = [];
  let stopped = false;
  let last = performance.now();
  const tick = () => {
    const now = performance.now();
    gaps.push(now - last);
    last = now;
    if (!stopped) setImmediate(tick);
  };
  setImmediate(tick);
  const started = performance.now();
  const value = await operation();
  const totalMs = performance.now() - started;
  stopped = true;
  await nextImmediate();
  return {
    value,
    totalMs: round(totalMs),
    synchronousMs: round(Math.max(0, ...gaps)),
    maxFrameGapMs: round(Math.max(0, ...gaps)),
    gapsOver16Ms: gaps.filter((gap) => gap > 16.7).length,
  };
}

export async function runDocxSaveExportBenchmarkTier(pageCount) {
  const documentBuffer = await generateDeterministicDocx(pageCount);
  const save = await measureOperation(async () => {
    const zip = await JSZip.loadAsync(documentBuffer.slice(0));
    const xml = await zip.file("word/document.xml").async("string");
    const serializeStarted = performance.now();
    zip.file("word/document.xml", xml.replace("Benchmark Section Heading", "Saved Benchmark Section Heading"));
    const serialized = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
    return { serializedBytes: serialized.byteLength, serializationMs: round(performance.now() - serializeStarted) };
  });
  const pdfExport = await measureOperation(async () => {
    const zip = await JSZip.loadAsync(documentBuffer.slice(0));
    const xml = await zip.file("word/document.xml").async("string");
    return { renderedTextBytes: Buffer.byteLength(xml, "utf8") };
  });
  return {
    format: "docx", tier: pageCount, unit: "pages", documentSizeBytes: documentBuffer.byteLength,
    save: { totalMs: save.totalMs, synchronousMs: save.synchronousMs, maxFrameGapMs: save.maxFrameGapMs, outputBytes: save.value.serializedBytes, serializationMs: save.value.serializationMs },
    pdfExport: { totalMs: pdfExport.totalMs, synchronousMs: pdfExport.synchronousMs, maxFrameGapMs: pdfExport.maxFrameGapMs, outputBytes: pdfExport.value.renderedTextBytes },
    complexity: { pages: pageCount },
  };
}

export async function runPptxSaveExportBenchmarkTier(slideCount) {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const deck = createDeck({ format: "pptx", slideCount, richFirstSlide: true });
  const buffer = deck.buffer.slice(deck.byteOffset, deck.byteOffset + deck.byteLength);
  const engine = await PresentationEngine.load(buffer.slice(0));
  const save = await measureOperation(async () => {
    const output = await engine.export();
    return { outputBytes: output.byteLength };
  });
  const pdfExport = await measureOperation(async () => {
    let renderedBytes = 0;
    for (let index = 0; index < engine.slideCount; index += 1) renderedBytes += engine.renderSlide(index).svg.length;
    return { renderedBytes };
  });
  return {
    format: "pptx", tier: slideCount, unit: "slides", documentSizeBytes: buffer.byteLength,
    save: { totalMs: save.totalMs, synchronousMs: save.synchronousMs, maxFrameGapMs: save.maxFrameGapMs, outputBytes: save.value.outputBytes },
    pdfExport: { totalMs: pdfExport.totalMs, synchronousMs: pdfExport.synchronousMs, maxFrameGapMs: pdfExport.maxFrameGapMs, outputBytes: pdfExport.value.renderedBytes },
    complexity: { slides: slideCount },
  };
}

export async function runSaveExportBenchmarks(options = {}) {
  const tiers = options.tiers ?? DEFAULT_SAVE_EXPORT_TIERS;
  const results = { docx: [], pptx: [] };
  for (const tier of tiers) {
    options.onProgress?.(`Running DOCX save/export tier ${tier}...`);
    results.docx.push(await runDocxSaveExportBenchmarkTier(tier));
    options.onProgress?.(`Running PPTX save/export tier ${tier}...`);
    results.pptx.push(await runPptxSaveExportBenchmarkTier(tier));
  }
  return { schemaVersion: 1, timestamp: new Date().toISOString(), results };
}
