import assert from "node:assert/strict";
import { test } from "node:test";
import { loadDocxFloatingLayerLayoutModule } from "./helpers/load-plugin-modules.mjs";

test("fullscreen dialog classification excludes non-modal color-picker dialogs", async () => {
	const { isFullscreenDialogLayer } = await loadDocxFloatingLayerLayoutModule();

	assert.equal(isFullscreenDialogLayer("dialog", null, false), false);
	assert.equal(isFullscreenDialogLayer("dialog", "false", false), false);
	assert.equal(isFullscreenDialogLayer("dialog", "true", false), true);
	assert.equal(isFullscreenDialogLayer(null, null, true), true);
});

test("formatting dropdown scroll transition preserves and restores the toolbar offset", async () => {
	const { getFormattingDropdownScrollTransition } = await loadDocxFloatingLayerLayoutModule();

	assert.deepEqual(
		getFormattingDropdownScrollTransition(true, 300, null),
		{ savedScrollLeft: 300, visualOffset: 300, restoreScrollLeft: null },
	);
	assert.deepEqual(
		getFormattingDropdownScrollTransition(true, 0, 300),
		{ savedScrollLeft: 300, visualOffset: 300, restoreScrollLeft: null },
	);
	assert.deepEqual(
		getFormattingDropdownScrollTransition(false, 0, 300),
		{ savedScrollLeft: null, visualOffset: 0, restoreScrollLeft: 300 },
	);
});

test("floating dropdown position clamps right and bottom overflow into the viewport", async () => {
	const { calculateClampedFloatingLayerPosition } = await loadDocxFloatingLayerLayoutModule();
	const next = calculateClampedFloatingLayerPosition(
		{ left: 1003.375, right: 1223.375, top: 600, bottom: 918, width: 220, height: 318 },
		{ width: 1097, height: 800 },
		{ left: 1003.88, top: 600 },
	);

	assert.ok(Math.abs(next.left - 869.505) < 0.001);
	assert.equal(next.top, 474);
});

test("floating dropdown position preserves visible layers and clamps oversized layers to the margin", async () => {
	const { calculateClampedFloatingLayerPosition } = await loadDocxFloatingLayerLayoutModule();

	assert.deepEqual(
		calculateClampedFloatingLayerPosition(
			{ left: 100, right: 320, top: 120, bottom: 438, width: 220, height: 318 },
			{ width: 1097, height: 800 },
			{ left: 100, top: 120 },
		),
		{ left: 100, top: 120 },
	);
	assert.deepEqual(
		calculateClampedFloatingLayerPosition(
			{ left: -20, right: 1180, top: -10, bottom: 890, width: 1200, height: 900 },
			{ width: 1097, height: 800 },
			{ left: -20, top: -10 },
		),
		{ left: 8, top: 8 },
	);
});
