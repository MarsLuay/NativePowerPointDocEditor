import assert from "node:assert/strict";
import test from "node:test";
import { loadTextUtilsModule } from "./helpers/load-plugin-modules.mjs";

test("inline paragraph preview preserves its wrapped SVG line containers after a backspace", async () => {
  const { redistributeTextAcrossVisualRuns } = await loadTextUtilsModule();
  const previousLines = [
    "They make setup easier for newer users and reduce the need to ",
    "watch the whole process.",
  ];
  const nextText = "They make setup easier for newer users and reduce the nee to watch the whole process.";

  const previewLines = redistributeTextAcrossVisualRuns(previousLines, nextText);

  assert.equal(previewLines.join(""), nextText);
  assert.equal(previewLines.length, previousLines.length);
  assert.equal(previewLines[0], previousLines[0].replace("need", "nee"));
  assert.equal(previewLines[1], previousLines[1]);
});

test("inline paragraph preview does not duplicate replacement text across wrapped lines", async () => {
  const { redistributeTextAcrossVisualRuns } = await loadTextUtilsModule();
  const previewLines = redistributeTextAcrossVisualRuns(["first ", "second ", "third"], "replacement");

  assert.deepEqual(previewLines, ["replacement", "", ""]);
});

test("inline paragraph preview keeps an end insertion inside the final visual line", async () => {
  const { redistributeTextAcrossVisualRuns } = await loadTextUtilsModule();
  const previewLines = redistributeTextAcrossVisualRuns(["first ", "second"], "first second!");

  assert.deepEqual(previewLines, ["first ", "second!"]);
});
