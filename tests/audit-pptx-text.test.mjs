import assert from "node:assert/strict";
import { test } from "node:test";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import Module, { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const require = createRequire(import.meta.url);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

// src/powerpoint/textUtils.ts imports `Platform` from "obsidian" at module
// scope, so bundle it standalone with obsidian marked external and stub the
// dependency at require time (mirrors load-plugin-modules.mjs without editing it).
let textUtilsPromise;
function loadTextUtils() {
  textUtilsPromise ??= (async () => {
    const outputDirectory = await mkdtemp(path.join(tmpdir(), "audit-pptx-text-"));
    const outfile = path.join(outputDirectory, "text-utils.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/powerpoint/textUtils.ts")],
      bundle: true,
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
        return { Platform: { isMacOS: false, isDesktop: true, isMobile: false } };
      }
      return originalLoad.call(this, request, parent, isMain);
    };
    try {
      return require(outfile);
    } finally {
      Module._load = originalLoad;
    }
  })();
  return textUtilsPromise;
}

test("parsePrimaryFontFamily extracts the first real font face", async () => {
  const { parsePrimaryFontFamily } = await loadTextUtils();

  assert.equal(parsePrimaryFontFamily('"Calibri", sans-serif'), "Calibri");
  assert.equal(parsePrimaryFontFamily("'Times New Roman', serif"), "Times New Roman");
  assert.equal(parsePrimaryFontFamily("Arial, Helvetica, sans-serif"), "Arial");
  assert.equal(parsePrimaryFontFamily("Helvetica Neue"), "Helvetica Neue");

  // Generic families, empties, and CSS-wide keywords resolve to null.
  assert.equal(parsePrimaryFontFamily("sans-serif"), null);
  assert.equal(parsePrimaryFontFamily("SYSTEM-UI, sans-serif"), null);
  assert.equal(parsePrimaryFontFamily(""), null);
  assert.equal(parsePrimaryFontFamily("   "), null);
  assert.equal(parsePrimaryFontFamily("inherit"), null);
  assert.equal(parsePrimaryFontFamily("initial"), null);
  assert.equal(parsePrimaryFontFamily("unset"), null);
});

test("normalizeSearchText collapses whitespace and trims", async () => {
  const { normalizeSearchText } = await loadTextUtils();

  assert.equal(normalizeSearchText("  hello   world  "), "hello world");
  assert.equal(normalizeSearchText("\t\nfoo\n  bar \r"), "foo bar");
  assert.equal(normalizeSearchText("single"), "single");
  assert.equal(normalizeSearchText("   "), "");
  assert.equal(normalizeSearchText(""), "");
});

test("PresentationEngine run-style edit applies, renders, and round-trips", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  await assert.doesNotReject(
    engine.setRunStyle(0, 0, { paragraphIndex: 0, runIndex: 0 }, {
      bold: true,
      italic: true,
      underline: true,
      fontFamily: "Georgia",
      fontSizePt: 33,
      color: "112233",
      highlight: "FFEE00",
    }),
  );

  const style = engine.getRunStyle(0, 0, 0, 0);
  assert.equal(style?.bold, true);
  assert.equal(style?.italic, true);
  assert.equal(style?.underline, true);
  assert.equal(style?.fontFamily, "Georgia");
  assert.equal(style?.fontSizePt, 33);
  assert.equal(style?.color, "112233");
  assert.equal(style?.highlight, "FFEE00");

  const exported = await engine.export();
  const reloaded = await PresentationEngine.load(exported);
  assert.match(reloaded.renderSlide(0).svg, /^<svg\b/);

  const reloadedStyle = reloaded.getRunStyle(0, 0, 0, 0);
  assert.equal(reloadedStyle?.bold, true);
  assert.equal(reloadedStyle?.italic, true);
  assert.equal(reloadedStyle?.underline, true);
  assert.equal(reloadedStyle?.fontFamily, "Georgia");
  assert.equal(reloadedStyle?.fontSizePt, 33);
  assert.equal(reloadedStyle?.color, "112233");
});

test("PresentationEngine paragraph alignment edit survives export", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  await assert.doesNotReject(engine.setParagraphAlignment(0, 0, 0, "ctr"));
  assert.equal(engine.getRunStyle(0, 0, 0, 0)?.alignment, "ctr");

  const reloaded = await PresentationEngine.load(await engine.export());
  assert.match(reloaded.renderSlide(0).svg, /^<svg\b/);
  assert.equal(reloaded.getRunStyle(0, 0, 0, 0)?.alignment, "ctr");
});
