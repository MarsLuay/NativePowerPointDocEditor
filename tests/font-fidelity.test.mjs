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

test("FontFidelity constructor options", async () => {
  const { FontFidelity, POWERPOINT_FONT_FALLBACKS } = await loadFontFidelityModule();

  const customFallbacks = { 'Custom Font': ['Fallback1', 'sans-serif'] };
  const fidelity = new FontFidelity({ fontFallbacks: customFallbacks });

  const fallbacks = fidelity.getRendererFallbacks();
  assert.deepEqual(fallbacks['Custom Font'], ['Fallback1', 'sans-serif']);
  assert.deepEqual(fallbacks['Aptos'], POWERPOINT_FONT_FALLBACKS['Aptos']);
});

test("FontFidelity.prototype.measureText delegates to measureTextFn and safely fallbacks", async () => {
  const { FontFidelity } = await loadFontFidelityModule();

  const mockMeasureText = (text, fontFamily, fontSizePx) => {
    if (text === "invalid") return NaN;
    if (text === "negative") return -1;
    return text.length * fontSizePx;
  };

  const fidelity = new FontFidelity({
    measureText: mockMeasureText,
    isFontAvailable: () => true
  });

  // Delegates properly
  assert.equal(fidelity.measureText("hello", "Arial", 10), 50);

  // Safely falls back on NaN
  assert.equal(fidelity.measureText("invalid", "Arial", 10), 7 * 10 * 0.6);

  // Safely falls back on negative
  assert.equal(fidelity.measureText("negative", "Arial", 10), 8 * 10 * 0.6);
});

class MockElement {
  constructor(tagName) {
    this.tagName = tagName;
    this.attributes = new Map();
    this.children = [];
  }
  getAttribute(name) {
    return this.attributes.get(name) || null;
  }
  setAttribute(name, value) {
    this.attributes.set(name, value);
  }
  getElementsByTagName(name) {
    let results = [];
    for (const child of this.children) {
      if (name === '*' || child.tagName === name) {
        results.push(child);
      }
      results.push(...child.getElementsByTagName(name));
    }
    return results;
  }
}

test("FontFidelity.prototype.applySvgSubstitutions substitutes fonts and styles", async () => {
  const { FontFidelity } = await loadFontFidelityModule();

  const fidelity = new FontFidelity({
    isFontAvailable: (font) => font.toLowerCase() === "arial" || font.toLowerCase() === "sans-serif"
  });

  const svg = new MockElement("svg");
  svg.setAttribute("font-family", "Unknown Font");

  const textElement = new MockElement("text");
  textElement.setAttribute("style", "font-family: 'Another Unknown'; color: red;");
  svg.children.push(textElement);

  const substitutions = fidelity.applySvgSubstitutions(svg);

  // Check svg attributes
  assert.equal(svg.getAttribute("data-native-powerpoint-requested-font"), "Unknown Font");
  assert.equal(svg.getAttribute("data-native-powerpoint-font-substitution"), "Arial");
  assert.ok(svg.getAttribute("font-family").includes("Arial"));

  // Check text element style
  assert.equal(textElement.getAttribute("data-native-powerpoint-requested-font"), "Another Unknown");
  assert.equal(textElement.getAttribute("data-native-powerpoint-font-substitution"), "Arial");
  assert.ok(textElement.getAttribute("style").includes("font-family:"));
  assert.ok(textElement.getAttribute("style").includes("color: red;"));

  // Check substitutions array
  assert.equal(substitutions.length, 2);
  // Sorted by requested font string
  assert.equal(substitutions[0].requested, "Another Unknown");
  assert.equal(substitutions[0].substitute, "Arial");
  assert.equal(substitutions[1].requested, "Unknown Font");
  assert.equal(substitutions[1].substitute, "Arial");
});

test("FontFidelity caches isFontAvailableFn results", async () => {
  const { FontFidelity } = await loadFontFidelityModule();

  let callCount = 0;
  const mockIsFontAvailable = (font) => {
    callCount++;
    return true;
  };

  const fidelity = new FontFidelity({
    isFontAvailable: mockIsFontAvailable
  });

  fidelity.measureText("hello", "CustomCacheFont", 10);
  const callsAfterFirst = callCount;

  // Should not trigger again
  fidelity.measureText("hello", "CustomCacheFont", 10);
  assert.equal(callCount, callsAfterFirst);

  // Same font with quotes and mixed casing should hit cache
  fidelity.measureText("hello", "'customcachefont'", 10);
  assert.equal(callCount, callsAfterFirst);
});
