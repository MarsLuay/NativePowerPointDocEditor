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

	assert.deepEqual(parseEditorZoomTransform("none"), { translateXPx: 0, translateYPx: 0, scale: 1 });
	assert.deepEqual(parseEditorZoomTransform("translateX(-176px) scale(1.25)"), {
		translateXPx: -176,
		translateYPx: 0,
		scale: 1.25,
	});
	assert.deepEqual(parseEditorZoomTransform("matrix(1.25, 0, 0, 1.25, -176, 24)"), {
		translateXPx: -176,
		translateYPx: 24,
		scale: 1.25,
	});
	assert.deepEqual(parseEditorZoomTransform("scale(1)"), { translateXPx: 0, translateYPx: 0, scale: 1 });
	assert.equal(editorZoomTransformNeedsNeutralization("translateX(-176px) scale(1.25)"), true);
	assert.equal(editorZoomTransformNeedsNeutralization("translate(0px, 24px)"), true);
	assert.equal(editorZoomTransformNeedsNeutralization("none"), false);
});

test("calculateHiddenImeAnchorPosition aligns hidden IME caret to visible caret", async () => {
	const { calculateHiddenImeAnchorPosition } = await loadNeutralizerModule();

	const next = calculateHiddenImeAnchorPosition(
		-9999,
		0,
		{ left: -9400, top: 200, bottom: 220, height: 20 },
		{ left: 300, top: 400, bottom: 424, height: 24 },
	);

	assert.deepEqual(next, { leftPx: -299, topPx: 204 });

	const hiddenCaretLeftAfterMove = -9400 + (next.leftPx - -9999);
	const hiddenCaretBottomAfterMove = 220 + (next.topPx - 0);
	assert.equal(hiddenCaretLeftAfterMove, 300);
	assert.equal(hiddenCaretBottomAfterMove, 424);
});
