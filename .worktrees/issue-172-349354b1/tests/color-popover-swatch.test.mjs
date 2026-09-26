import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import test from "node:test";

const projectFile = (path) => resolve(path);

test("shape fill context action uses concise change-color copy", async () => {
  const locale = JSON.parse(await readFile(projectFile("locales/en/powerpoint.json"), "utf8"));

  assert.equal(locale.contextMenu.shapeFillColor, "Change color");
});

test("color pickers render a dedicated visible fill inside every swatch button", async () => {
  const [viewSource, toolbarSource, css] = await Promise.all([
    readFile(projectFile("src/powerpoint/ui/NativePowerPointView.ts"), "utf8"),
    readFile(projectFile("src/powerpoint/textToolbarController.ts"), "utf8"),
    readFile(projectFile("styles.css"), "utf8"),
  ]);

  assert.match(viewSource, /native-powerpoint-color-popover-swatch-fill/);
  assert.match(toolbarSource, /native-powerpoint-color-popover-swatch-fill/);
  assert.match(css, /\.native-powerpoint-color-popover-swatch-fill\s*\{[\s\S]*?background(?:-color)?:\s*var\(--np-swatch-color/);
  assert.match(css, /\.native-powerpoint-color-popover-swatch-fill\s*\{[\s\S]*?pointer-events:\s*none/);
});
