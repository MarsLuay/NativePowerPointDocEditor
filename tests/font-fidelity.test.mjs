import assert from "node:assert/strict";
import { test } from "node:test";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";
import { createRequire } from "node:module";
import Module from "node:module";

const originalLoad = Module._load;
Module._load = function load(request, parent, isMain) {
  if (request === "pptx-svg") {
    return {
      DEFAULT_FONT_FALLBACKS: {},
    };
  }
  return originalLoad.call(this, request, parent, isMain);
};

const require = createRequire(import.meta.url);
let fontFidelityModulePromise;

function loadFontFidelityModule() {
  fontFidelityModulePromise ??= bundleSource(
    "src/FontFidelity.ts",
    "font-fidelity.cjs",
    ["pptx-svg"],
    []
  ).then((outfile) => require(outfile));
  return fontFidelityModulePromise;
}

test("splitCssFontFamilies", async () => {
  const { splitCssFontFamilies } = await loadFontFidelityModule();

  // Single unquoted
  assert.deepEqual(splitCssFontFamilies("Arial"), ["Arial"]);

  // Comma separated
  assert.deepEqual(splitCssFontFamilies("Arial, sans-serif"), ["Arial", "sans-serif"]);

  // Double quotes
  assert.deepEqual(splitCssFontFamilies('"Times New Roman", Times, serif'), ["Times New Roman", "Times", "serif"]);

  // Single quotes
  assert.deepEqual(splitCssFontFamilies("'Courier New', Courier, monospace"), ["Courier New", "Courier", "monospace"]);

  // Whitespace trimming
  assert.deepEqual(splitCssFontFamilies("  'Comic Sans MS'  ,  Arial  "), ["Comic Sans MS", "Arial"]);

  // Quotes containing commas
  assert.deepEqual(splitCssFontFamilies('"Aptos, Display", Arial'), ["Aptos, Display", "Arial"]);
  assert.deepEqual(splitCssFontFamilies('Multiple, "quoted families", together'), ["Multiple", "quoted families", "together"]);

  // Empty inputs
  assert.deepEqual(splitCssFontFamilies(""), []);
  assert.deepEqual(splitCssFontFamilies("   "), []);
});
