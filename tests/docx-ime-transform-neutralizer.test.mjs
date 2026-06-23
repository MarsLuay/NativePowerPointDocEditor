import assert from "node:assert/strict";
import { test } from "node:test";
import { build } from "esbuild";
import { mkdtemp } from "node:fs/promises";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

let neutralizerModulePromise;

async function loadNeutralizerModule() {
	neutralizerModulePromise ??= (async () => {
		const tempDir = await mkdtemp(path.join(os.tmpdir(), "docx-ime-neutralizer-"));
		const outfile = path.join(tempDir, "neutralizer.cjs");
		await build({
			entryPoints: [path.join(projectRoot, "src/docxImeTransformNeutralizer.ts")],
			bundle: true,
			format: "cjs",
			logLevel: "silent",
			outfile,
			platform: "node",
			target: "node22",
		});
		return require(outfile);
	})();

	return neutralizerModulePromise;
}

test("parseEditorZoomTransform reads translateX and scale from eigenpal-style transforms", async () => {
	const {
		parseEditorZoomTransform,
		editorZoomTransformNeedsNeutralization,
	} = await loadNeutralizerModule();

	assert.deepEqual(parseEditorZoomTransform("none"), { translateXPx: 0, scale: 1 });
	assert.deepEqual(parseEditorZoomTransform("translateX(-176px) scale(1.25)"), {
		translateXPx: -176,
		scale: 1.25,
	});
	assert.deepEqual(parseEditorZoomTransform("scale(1)"), { translateXPx: 0, scale: 1 });
	assert.equal(editorZoomTransformNeedsNeutralization("translateX(-176px) scale(1.25)"), true);
	assert.equal(editorZoomTransformNeedsNeutralization("none"), false);
});
