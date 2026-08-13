import assert from "node:assert/strict";
import { readFile, readdir } from "node:fs/promises";
import { test } from "node:test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("shipped empty-list Enter clears the marker without splitting another paragraph", async () => {
  const distDir = path.join(projectRoot, "vendor/docx-editor-runtime/core/dist");
  const marker = "numPr:null,listIsBullet:null,listNumFmt:null,listMarker:null,borders:null";
  const matches = [];

  for (const filename of await readdir(distDir)) {
    if (!filename.endsWith(".js")) continue;
    const source = await readFile(path.join(distDir, filename), "utf8");
    const markerIndex = source.indexOf(marker);
    if (markerIndex === -1) continue;
    const functionStart = source.lastIndexOf("function ", markerIndex);
    const functionEnd = source.indexOf("function ", markerIndex);
    matches.push({
      filename,
      handler: source.slice(functionStart, functionEnd),
    });
  }

  assert.equal(matches.length, 1, "expected one shipped empty-list Enter handler");
  assert.match(matches[0].handler, /setNodeMarkup\(/);
  assert.doesNotMatch(
    matches[0].handler,
    /\.split\(/,
    `${matches[0].filename} must not add a second empty paragraph while exiting the list`,
  );
});

test("DOCX Enter diagnostics record empty-list state without document text", async () => {
  const source = await readFile(path.join(projectRoot, "src/DocxReactView.tsx"), "utf8");

  assert.match(source, /debugLog\('text-input', 'DOCX Enter key received'/);
  assert.match(source, /paragraphContentSize: paragraph\.content\.size/);
  assert.match(source, /isList: Boolean\(listProperties\)/);
  assert.doesNotMatch(source, /DOCX Enter key received'[\s\S]{0,700}textContent/);
});
