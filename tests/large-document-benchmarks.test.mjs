import assert from "node:assert/strict";
import { test } from "node:test";
import {
  generateDeterministicDocx,
  generateDeterministicPptx,
  runDocxBenchmarkTier,
  runPptxBenchmarkTier,
  runLargeDocumentBenchmarks,
} from "../scripts/lib/large-document-benchmarks.mjs";

function sameBytes(bufA, bufB) {
  if (bufA.byteLength !== bufB.byteLength) return false;
  const viewA = new Uint8Array(bufA);
  const viewB = new Uint8Array(bufB);
  for (let i = 0; i < viewA.length; i++) {
    if (viewA[i] !== viewB[i]) return false;
  }
  return true;
}

test("deterministic fixture generation produces bit-for-bit identical fixtures for DOCX and PPTX", async () => {
  const docx1 = await generateDeterministicDocx(10);
  const docx2 = await generateDeterministicDocx(10);
  assert.equal(docx1.byteLength, docx2.byteLength);
  assert.ok(sameBytes(docx1, docx2), "DOCX generation must be bit-for-bit deterministic");

  const pptx1 = generateDeterministicPptx(10);
  const pptx2 = generateDeterministicPptx(10);
  assert.equal(pptx1.byteLength, pptx2.byteLength);
  assert.ok(sameBytes(pptx1, pptx2), "PPTX generation must be bit-for-bit deterministic");
});

test("DOCX benchmark exercises real stages and records non-zero timings", async () => {
  const result = await runDocxBenchmarkTier(5);
  assert.equal(result.tier, 5);
  assert.equal(result.unit, "pages");
  assert.ok(result.documentSizeBytes > 1000, "Document size should reflect 5 pages of OOXML");

  assert.ok(Number.isFinite(result.openToFirstVisibleMs) && result.openToFirstVisibleMs > 0);
  assert.ok(Number.isFinite(result.openToUsableMs) && result.openToUsableMs >= result.openToFirstVisibleMs);
  assert.ok(Number.isFinite(result.coldNavigationMs));
  assert.ok(Number.isFinite(result.warmNavigationMs));
  assert.ok(Number.isFinite(result.typingMutationMs) && result.typingMutationMs > 0);
  assert.ok(Number.isFinite(result.searchIndexMs) && result.searchIndexMs > 0);
  assert.ok(Number.isFinite(result.saveMs) && result.saveMs > 0);
  assert.ok(Number.isFinite(result.backgroundCompletionMs));
});

test("PPTX benchmark exercises real PresentationEngine and records real timings", async () => {
  const result = await runPptxBenchmarkTier(5);
  assert.equal(result.tier, 5);
  assert.equal(result.unit, "slides");
  assert.ok(result.documentSizeBytes > 10000, "Presentation size should reflect 5 slides of OOXML");

  assert.ok(Number.isFinite(result.openToFirstVisibleMs) && result.openToFirstVisibleMs > 0);
  assert.ok(Number.isFinite(result.openToUsableMs) && result.openToUsableMs >= result.openToFirstVisibleMs);
  assert.ok(Number.isFinite(result.coldNavigationMs));
  assert.ok(Number.isFinite(result.warmNavigationMs));
  assert.ok(Number.isFinite(result.typingMutationMs) && result.typingMutationMs > 0);
  assert.ok(Number.isFinite(result.searchIndexMs) && result.searchIndexMs > 0);
  assert.ok(Number.isFinite(result.saveMs) && result.saveMs > 0);
  assert.ok(Number.isFinite(result.backgroundCompletionMs));
});

test("benchmark runner produces machine-readable schema and demonstrates tier scaling", async () => {
  const report = await runLargeDocumentBenchmarks({
    tiers: {
      docx: [5, 15, 30],
      pptx: [5, 15, 30],
    },
  });

  assert.equal(report.schemaVersion, 1);
  assert.ok(report.timestamp);
  assert.ok(report.environment.platform);
  assert.ok(report.environment.nodeVersion);

  assert.equal(report.results.docx.length, 3);
  assert.equal(report.results.pptx.length, 3);

  // Validate scaling across the 3 tiers
  const [d5, d15, d30] = report.results.docx;
  assert.ok(d15.documentSizeBytes > d5.documentSizeBytes, "DOCX size scales with tier");
  assert.ok(d30.documentSizeBytes > d15.documentSizeBytes, "DOCX size scales with tier");

  const [p5, p15, p30] = report.results.pptx;
  assert.ok(p15.documentSizeBytes > p5.documentSizeBytes, "PPTX size scales with tier");
  assert.ok(p30.documentSizeBytes > p15.documentSizeBytes, "PPTX size scales with tier");
  assert.ok(p30.saveMs > p5.saveMs, "PPTX save latency scales with tier size");
});
