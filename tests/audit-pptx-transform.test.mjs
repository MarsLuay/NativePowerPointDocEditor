// Headless audit tests for the "PPTX · Shapes: transform" feature group.
//
// Strategy (mirrors tests/helpers/load-plugin-modules.mjs `bundleSource`):
//   * src/powerpoint/svgUtils.ts is bundled with esbuild to a temp CJS file and
//     required. Only the genuinely pure exports are exercised here
//     (cloneTransform / parseSvgLength / transformsMatch) — the DOM-bound helpers
//     (getSvgIntrinsicSize / ensureSvgViewBox) need a real SVGSVGElement.viewBox
//     and are out of scope for a pure unit test.
//   * The unit-conversion functions (emuToPx / pxToEmu / degreesToOoxml /
//     ooxmlToDegrees) are NOT standalone exports of PresentationEngine.ts — that
//     module only re-exposes them as instance methods on a class whose
//     constructor is private and needs a real WASM renderer. They are imported by
//     PresentationEngine.ts from the "pptx-svg" package, so this test imports them
//     from the exact same specifier the plugin source uses.

import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { createRequire } from "node:module";
import Module from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { build } from "esbuild";

// Same unit-conversion functions PresentationEngine.ts imports (see its line 1-12
// `import { ... emuToPx, pxToEmu, degreesToOoxml, ooxmlToDegrees ... } from 'pptx-svg'`).
import {
  degreesToOoxml,
  emuToPx,
  ooxmlToDegrees,
  pxToEmu,
} from "pptx-svg";

const require = createRequire(import.meta.url);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

// Mirror the helper's xmldom globals so a bundled module that touches DOM types
// at load time would not crash (the pure exports under test do not, but this
// keeps parity with tests/helpers/load-plugin-modules.mjs).
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
globalThis.DOMParser ??= DOMParser;
globalThis.XMLSerializer ??= XMLSerializer;

async function bundleSvgUtils() {
  const outputDirectory = await mkdtemp(path.join(tmpdir(), "audit-pptx-transform-"));
  const outfile = path.join(outputDirectory, "svg-utils.cjs");
  await build({
    entryPoints: [path.join(projectRoot, "src/powerpoint/svgUtils.ts")],
    bundle: true,
    // "obsidian" is only reachable via textUtils.ts (Platform) and is stubbed below.
    external: ["obsidian"],
    format: "cjs",
    logLevel: "silent",
    outfile,
    platform: "node",
    target: "node22",
  });

  const originalLoad = Module._load;
  Module._load = function load(request, parent, isMain) {
    if (request === "obsidian") {
      return { Platform: { isDesktop: true, isMobile: false, isMacOS: false } };
    }
    return originalLoad.call(this, request, parent, isMain);
  };
  try {
    return require(outfile);
  } finally {
    Module._load = originalLoad;
  }
}

const svgUtils = await bundleSvgUtils();

test("emuToPx and pxToEmu round-trip for 96-DPI-aligned pixel values", () => {
  // EMU_PER_PX_96DPI = 9525, so pixel values map to whole EMU and back exactly.
  for (const px of [0, 1, 12, 100, 720, 960, 1280]) {
    const emu = pxToEmu(px);
    assert.equal(emu, px * 9525, `pxToEmu(${px}) should be ${px} * 9525`);
    assert.equal(emuToPx(emu), px, `emuToPx(pxToEmu(${px})) should round-trip`);
  }
});

test("emuToPx is the rounding inverse of pxToEmu for arbitrary EMU values", () => {
  for (const emu of [0, 9525, 12700, 914400, 6858000]) {
    assert.equal(pxToEmu(emuToPx(emu)), Math.round(emuToPx(emu)) * 9525);
  }
  // Spot-check the documented examples.
  assert.equal(emuToPx(914400), 96); // 1 inch at 96 DPI
  assert.equal(pxToEmu(96), 914400);
});

test("degreesToOoxml and ooxmlToDegrees round-trip", () => {
  for (const deg of [0, 15, 45, 90, 180, 270, 359, 360]) {
    const ooxml = degreesToOoxml(deg);
    assert.equal(ooxml, deg * 60000, `degreesToOoxml(${deg}) should be ${deg} * 60000`);
    assert.equal(ooxmlToDegrees(ooxml), deg, `ooxmlToDegrees round-trips ${deg}`);
  }
  // Documented examples.
  assert.equal(degreesToOoxml(90), 5400000);
  assert.equal(ooxmlToDegrees(5400000), 90);
});

test("parseSvgLength parses unitless and px values, rejects %", () => {
  assert.equal(svgUtils.parseSvgLength("100"), 100);
  assert.equal(svgUtils.parseSvgLength("100px"), 100);
  assert.equal(svgUtils.parseSvgLength("12.5"), 12.5);
  assert.equal(svgUtils.parseSvgLength("12.5px"), 12.5);
  assert.equal(svgUtils.parseSvgLength("  48px  "), 48); // trimmed
  assert.equal(svgUtils.parseSvgLength("50%"), null);
  assert.equal(svgUtils.parseSvgLength(null), null);
  assert.equal(svgUtils.parseSvgLength(""), null);
  assert.equal(svgUtils.parseSvgLength("0"), null); // requires > 0
  assert.equal(svgUtils.parseSvgLength("-5"), null); // leading minus not matched
});

test("parseSvgLength does NOT parse pt (documented audit limitation)", () => {
  // AUDIT FINDING: the regex only allows an optional `px` suffix, so any explicit
  // `pt`/`em`/etc. unit returns null. This is fine for SVG width/height (which are
  // px or unitless), but the value is not a general CSS length parser.
  assert.equal(svgUtils.parseSvgLength("12pt"), null);
  assert.equal(svgUtils.parseSvgLength("2em"), null);
});

test("transformsMatch is true only when every field is equal", () => {
  const base = { x: 10, y: 20, cx: 100, cy: 50, rot: 0 };
  assert.equal(svgUtils.transformsMatch(base, { ...base }), true);
  assert.equal(svgUtils.transformsMatch(base, { ...base, x: 11 }), false);
  assert.equal(svgUtils.transformsMatch(base, { ...base, y: 21 }), false);
  assert.equal(svgUtils.transformsMatch(base, { ...base, cx: 101 }), false);
  assert.equal(svgUtils.transformsMatch(base, { ...base, cy: 51 }), false);
  assert.equal(svgUtils.transformsMatch(base, { ...base, rot: 60000 }), false);
});

test("cloneTransform makes an independent deep copy of the five fields", () => {
  const source = { x: 1, y: 2, cx: 3, cy: 4, rot: 5 };
  const copy = svgUtils.cloneTransform(source);
  assert.deepEqual(copy, source);
  assert.notEqual(copy, source); // new object reference
  copy.x = 999;
  copy.rot = 888;
  assert.equal(source.x, 1, "mutating the copy must not affect the source");
  assert.equal(source.rot, 5);
});

test("isSelectableShapeIndex rejects layout/master decorations and null", () => {
  assert.equal(svgUtils.isSelectableShapeIndex(0), true);
  assert.equal(svgUtils.isSelectableShapeIndex(3), true);
  assert.equal(svgUtils.isSelectableShapeIndex(-1), false);
  assert.equal(svgUtils.isSelectableShapeIndex(-10), false);
  assert.equal(svgUtils.isSelectableShapeIndex(null), false);
});
