import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("DOCX load traces terminate across lifecycle exits and render failures", async () => {
	const view = await readFile(path.join(projectRoot, "src/DocxView.tsx"), "utf8");
	const react = await readFile(path.join(projectRoot, "src/DocxReactView.tsx"), "utf8");
	const mount = await readFile(path.join(projectRoot, "src/DocxReactMount.tsx"), "utf8");

	const close = view.slice(view.indexOf("async onClose()"), view.indexOf("async onLoadFile("));
	const unload = view.slice(view.indexOf("async onUnloadFile("), view.indexOf("async onRename("));

	assert.ok(close.indexOf("if (!await this.promptToSaveIfDirty())") < close.indexOf("this.finishOpenLoadTrace("));
	assert.match(unload, /this\.finishOpenLoadTrace\('file-unloaded-before-ready'/);
	assert.match(view, /this\.finishOpenLoadTrace\('view-open-failed'/);
	assert.match(view, /this\.finishOpenLoadTrace\('file-load-canceled'/);
	assert.match(view, /this\.finishOpenLoadTrace\('file-load-failed'/);
	assert.match(view, /this\.finishOpenLoadTrace\('react-mount-failed'/);
	assert.match(view, /this\.finishOpenLoadTrace\('editor-view-ready'/);
	assert.match(view, /createDocxReactMount\(this\.hostEl, \(renderError\) =>/);
	assert.match(react, /onLoadPhase\?\.\('editor-render-error'/);
	assert.match(mount, /onRenderError\?\.\(error, errorInfo\)/);
});
