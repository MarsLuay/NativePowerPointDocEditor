import assert from "node:assert/strict";
import test from "node:test";
import { loadParseRenderedSlideSvgModule } from "./helpers/load-plugin-modules.mjs";

test("parseRenderedSlideSvg accepts SVG roots and augments missing query APIs", async () => {
  const { parseRenderedSlideSvg } = await loadParseRenderedSlideSvgModule();
  const root = parseRenderedSlideSvg(
    '<svg xmlns="http://www.w3.org/2000/svg"><g data-shape="one"><text>Hi</text></g></svg>',
  );

  assert.equal(root.localName, "svg");
  assert.equal(typeof root.querySelector, "function");
  assert.equal(root.querySelector('g[data-shape="one"]')?.getAttribute("data-shape"), "one");
});

test("parseRenderedSlideSvg rejects non-SVG document roots", async () => {
  const { parseRenderedSlideSvg } = await loadParseRenderedSlideSvgModule();

  assert.throws(
    () => parseRenderedSlideSvg("<html><body /></html>"),
    /Could not parse slide SVG/,
  );
});
