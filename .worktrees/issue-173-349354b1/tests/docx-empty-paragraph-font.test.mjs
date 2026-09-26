import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("empty-paragraph font family uses core setFontFamily + flush on tab deactivate", async () => {
	const react = await readFile(path.join(projectRoot, "src/DocxReactView.tsx"), "utf8");
	const view = await readFile(path.join(projectRoot, "src/DocxView.tsx"), "utf8");
	const markUtils = await readFile(
		path.join(projectRoot, "docx-editor/packages/core/src/prosemirror/extensions/marks/markUtils.ts"),
		"utf8",
	);
	const fromPara = await readFile(
		path.join(projectRoot, "docx-editor/packages/core/src/prosemirror/conversion/fromProseDoc/paragraph.ts"),
		"utf8",
	);

	assert.match(react, /setFontFamily\(fontFamily\)\(view\.state, view\.dispatch\)/);
	assert.match(react, /DOCX empty-paragraph font family applied/);
	assert.match(react, /defaults\?\.fontFamily\?\.ascii \?\? defaults\?\.fontFamily\?\.hAnsi/);
	assert.match(react, /insertPlainTextAsParagraphs/);
	assert.match(view, /flushPendingSave/);
	assert.match(view, /flush on tab deactivate|Await flush before remount/);
	assert.match(view, /await this\.getReactHandle\(\)\?\.flushPendingSave/);
	assert.match(markUtils, /runProperties: defaultTextFormatting/);
	assert.match(fromPara, /runProperties,/);
	assert.match(fromPara, /attrs\.defaultTextFormatting/);
});
