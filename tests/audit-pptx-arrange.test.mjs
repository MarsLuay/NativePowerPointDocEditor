// Audit test for the "PPTX · Shapes: arrange" feature group.
//
// Scope note (see the chat audit for the full write-up):
//   * Z-order reorder (front/forward/backward/back), group, and ungroup are
//     PUBLIC, deck-level methods on PresentationEngine (reorderShapes /
//     groupShapes / ungroupShapes). They edit slide OOXML directly via
//     mutateSlideTree and round-trip through export, so they ARE headless-
//     testable. This file exercises reorderShapes.
//   * Object distribute and object align (left/right/top/bottom/center/middle)
//     live ONLY in the view layer (NativePowerPointView.distributeSelectedShapes
//     / alignSelectedShapes); they mutate the live SVG via collectSelectedTransforms
//     and have NO PresentationEngine entry point, so they are manual/UI-only and
//     intentionally not covered here. alignSelectedShapes is additionally dead
//     code (no caller / no UI trigger) — documented in the audit, not testable.
//
// The assertions below confirm reorderShapes (a) actually changes top-level
// shape order in the exported slide XML and (b) leaves a deck that still
// exports and re-loads/renders.

import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");

// Mirrors PresentationEngine's SHAPE_ELEMENT_NAMES — the spTree children that
// count as top-level, index-addressable shapes.
const SHAPE_ELEMENT_NAMES = new Set(["cxnSp", "graphicFrame", "grpSp", "pic", "sp"]);

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function firstByLocalName(root, localName) {
  const all = root.getElementsByTagName("*");
  for (let i = 0; i < all.length; i++) {
    if (all[i].localName === localName) return all[i];
  }
  return null;
}

function elementChildren(element) {
  const result = [];
  for (let node = element.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1) result.push(node);
  }
  return result;
}

// Identity-preserving order of a slide's top-level shapes, keyed by each
// shape's first cNvPr id. Index position alone can't prove a reorder (indices
// are just slots), so we track stable ids.
async function topLevelShapes(zip, slideNumber) {
  const file = zip.files[`ppt/slides/slide${slideNumber}.xml`];
  assert.ok(file, `missing ppt/slides/slide${slideNumber}.xml`);
  const xml = await file.async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const spTree = firstByLocalName(doc, "spTree");
  assert.ok(spTree, `slide ${slideNumber} has no spTree`);
  return elementChildren(spTree)
    .filter((child) => SHAPE_ELEMENT_NAMES.has(child.localName))
    .map((shape) => {
      const cNvPr = firstByLocalName(shape, "cNvPr");
      return {
        id: cNvPr ? cNvPr.getAttribute("id") : null,
        kind: shape.localName,
      };
    });
}

async function topLevelShapeIds(zip, slideNumber) {
  return (await topLevelShapes(zip, slideNumber)).map((shape) => shape.id);
}

async function loadFixtureEngine(PresentationEngine) {
  const fixturePath = path.join(projectRoot, "tests/fixtures/decks/features.pptx");
  assert.ok(existsSync(fixturePath), `missing fixture: ${fixturePath}`);
  const fileBuffer = await readFile(fixturePath);
  return PresentationEngine.load(toArrayBuffer(fileBuffer));
}

// Locate a slide carrying >= 2 top-level shapes so a reorder is observable.
async function findMultiShapeSlide(engine) {
  const exported = await engine.export();
  const zip = await JSZip.loadAsync(exported);
  for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
    const ids = await topLevelShapeIds(zip, slideIndex + 1);
    if (ids.length >= 2) return { slideIndex, ids };
  }
  return null;
}

async function insertTextBoxAt(engine, slideIndex, x, y) {
  return engine.insertTextBox(slideIndex, { x, y });
}

async function addBlankTestSlide(engine) {
  const { slideIndex } = await engine.addSlide(engine.slideCount - 1);
  const shapeIds = await topLevelShapeIds(await JSZip.loadAsync(await engine.export()), slideIndex + 1);
  assert.deepEqual(shapeIds, [], "new test slide must not introduce pre-existing objects");
  return slideIndex;
}

test("reorderShapes('front') moves the first shape to the top of the z-order", async () => {
  const { PresentationEngine, setForceJsBackendOverride, resetForceJsBackendOverride } =
    await loadPresentationEngineModule();
  setForceJsBackendOverride(true);
  try {
    const engine = await loadFixtureEngine(PresentationEngine);

    const target = await findMultiShapeSlide(engine);
    assert.ok(target, "no slide in features.pptx has >= 2 top-level shapes");
    const { slideIndex, ids: before } = target;
    const count = before.length;

    // Bring the bottom-most shape (index 0) to the front.
    const newIndices = await engine.reorderShapes(slideIndex, [0], "front", { intersectingOnly: true });
    assert.deepEqual(
      newIndices,
      [count - 1],
      `expected shape 0 to move to the last slot (${count - 1})`,
    );

    const exported = await engine.export();
    const after = await topLevelShapeIds(await JSZip.loadAsync(exported), slideIndex + 1);

    // Same shapes, reordered: it must be a true permutation, not an add/drop.
    assert.deepEqual([...after].sort(), [...before].sort(), "shape set changed");
    assert.notDeepEqual(after, before, "shape order did not change");
    assert.equal(after[after.length - 1], before[0], "first shape did not move to front");
    assert.deepEqual(after.slice(0, count - 1), before.slice(1), "other shapes did not keep order");
  } finally {
    resetForceJsBackendOverride();
  }
});

test("reorderShapes('backward'/'forward') skips unrelated layers and stops at an overlap", async () => {
  const { PresentationEngine, setForceJsBackendOverride, resetForceJsBackendOverride } =
    await loadPresentationEngineModule();
  setForceJsBackendOverride(true);
  try {
    const engine = await loadFixtureEngine(PresentationEngine);
    const slideIndex = await addBlankTestSlide(engine);
    const overlappingBelow = await insertTextBoxAt(engine, slideIndex, 1_000_000, 1_000_000);
    await insertTextBoxAt(engine, slideIndex, 6_000_000, 6_000_000);
    const selected = await insertTextBoxAt(engine, slideIndex, 1_000_000, 1_000_000);

    const backward = await engine.reorderShapes(slideIndex, [selected], "backward", { intersectingOnly: true });
    assert.deepEqual(backward, [overlappingBelow], "backward must skip the unrelated middle layer");

    const forward = await engine.reorderShapes(slideIndex, backward, "forward", { intersectingOnly: true });
    assert.deepEqual(forward, [overlappingBelow + 1], "forward must stop at that same overlapping object");
  } finally {
    resetForceJsBackendOverride();
  }
});

test("reorderShapes('backward'/'forward') leaves isolated objects in place", async () => {
  const { PresentationEngine, setForceJsBackendOverride, resetForceJsBackendOverride } =
    await loadPresentationEngineModule();
  setForceJsBackendOverride(true);
  try {
    const engine = await loadFixtureEngine(PresentationEngine);
    const slideIndex = await addBlankTestSlide(engine);
    const isolated = await insertTextBoxAt(engine, slideIndex, 1_000_000, 1_000_000);
    const before = await topLevelShapeIds(await JSZip.loadAsync(await engine.export()), slideIndex + 1);

    const backward = await engine.reorderShapes(slideIndex, [isolated], "backward", { intersectingOnly: true });
    assert.equal(backward, null);
    const forward = await engine.reorderShapes(slideIndex, [isolated], "forward", { intersectingOnly: true });

    assert.equal(forward, null);
    const after = await topLevelShapeIds(await JSZip.loadAsync(await engine.export()), slideIndex + 1);
    assert.deepEqual(after, before, "isolated objects must retain their structural z-order");
  } finally {
    resetForceJsBackendOverride();
  }
});

test("a reordered deck still exports and re-loads/renders", async () => {
  const { PresentationEngine, setForceJsBackendOverride, resetForceJsBackendOverride } =
    await loadPresentationEngineModule();
  setForceJsBackendOverride(true);
  try {
    const engine = await loadFixtureEngine(PresentationEngine);

    const target = await findMultiShapeSlide(engine);
    assert.ok(target, "no slide in features.pptx has >= 2 top-level shapes");
    const { slideIndex } = target;

    await engine.reorderShapes(slideIndex, [0], "front");
    const exported = await engine.export();
    assert.ok(exported.byteLength > 0, "export produced an empty buffer");

    const reloaded = await PresentationEngine.load(exported);
    assert.ok(reloaded.slideCount > 0, "reloaded deck has no slides");
    const { svg } = reloaded.renderSlide(slideIndex);
    assert.equal(typeof svg, "string");
    assert.ok(svg.includes("<svg"), "reloaded deck did not render SVG markup");
    assert.ok(!svg.startsWith("ERROR:"), `renderSlide returned ${svg.slice(0, 80)}`);
  } finally {
    resetForceJsBackendOverride();
  }
});
