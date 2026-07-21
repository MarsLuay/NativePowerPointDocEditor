import assert from "node:assert/strict";
import { docxEditorAliases } from './helpers/docx-esbuild-aliases.mjs';
import { mkdtemp } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const require = createRequire(import.meta.url);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

async function loadHelper() {
  const outputDirectory = await mkdtemp(path.join(tmpdir(), "coalesce-pptx-ops-"));
  const outfile = path.join(outputDirectory, "coalesce-pptx-ops.cjs");
  await build({
		alias: docxEditorAliases,
    entryPoints: [path.join(projectRoot, "src/ai/coalescePptxOps.ts")],
    bundle: true,
    format: "cjs",
    logLevel: "silent",
    outfile,
    platform: "node",
    target: "node22",
  });
  return require(outfile);
}

const { coalescePptxOps } = await loadHelper();

test("coalescePptxOps merges consecutive same-slide deletes descending", () => {
  const coalesced = coalescePptxOps([
    { op: "pptx.deleteShape", slideIndex: 0, shapeIndex: 2 },
    { op: "pptx.deleteShape", slideIndex: 0, shapeIndex: 0 },
    { op: "pptx.deleteShape", slideIndex: 0, shapeIndex: 2 },
    { op: "pptx.updateShapeText", slideIndex: 0, shapeIndex: 1, text: "x" },
  ]);

  assert.equal(coalesced.length, 2);
  assert.deepEqual(coalesced[0], {
    op: "pptx.deleteShapes",
    slideIndex: 0,
    shapeIndexes: [2, 0],
  });
  assert.equal(coalesced[1].op, "pptx.updateShapeText");
});

test("coalescePptxOps leaves single deletes and other slides alone", () => {
  const coalesced = coalescePptxOps([
    { op: "pptx.deleteShape", slideIndex: 0, shapeIndex: 1 },
    { op: "pptx.deleteShape", slideIndex: 1, shapeIndex: 1 },
  ]);
  assert.equal(coalesced.length, 2);
  assert.equal(coalesced[0].op, "pptx.deleteShape");
  assert.equal(coalesced[1].op, "pptx.deleteShape");
});
