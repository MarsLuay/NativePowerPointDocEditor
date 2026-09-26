import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("DOCX close flushes pending save before unsaved prompt", async () => {
	const view = await readFile(path.join(projectRoot, "src/DocxView.tsx"), "utf8");
	const react = await readFile(path.join(projectRoot, "src/DocxReactView.tsx"), "utf8");

	assert.match(react, /flushPendingSave:\s*\(\)\s*=>\s*Promise<boolean>/);
	assert.match(react, /DOCX flush pending save before close/);
	assert.match(react, /await current\.save\('autosave'\)/);
	assert.match(react, /await current\.waitForIdle\(\)/);
	assert.match(view, /handle\?\.flushPendingSave/);
	assert.match(view, /Prompting for unsaved changes/);
	assert.match(
		react,
		/reason === 'change' && \(!dirtyTrackingEnabled \|\| !editorReady\)/,
	);
});
