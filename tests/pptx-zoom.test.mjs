import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadNativePowerPointViewModule } from './helpers/load-plugin-modules.mjs';

function createView(NativePowerPointView) {
	return new NativePowerPointView({ app: { vault: {} } }, () => ({
		autosaveEnabled: false,
		yoloMode: false,
	}));
}

function createWheelEvent({ deltaX = 0, deltaY = 0, ctrlKey = false } = {}) {
	let defaultPrevented = false;
	let propagationStopped = false;
	return {
		clientX: 100,
		clientY: 100,
		ctrlKey,
		deltaMode: 0,
		deltaX,
		deltaY,
		preventDefault() {
			defaultPrevented = true;
		},
		stopPropagation() {
			propagationStopped = true;
		},
		get defaultPrevented() {
			return defaultPrevented;
		},
		get propagationStopped() {
			return propagationStopped;
		},
	};
}

test('PowerPoint canvas scroll gestures pan, while pinch directions zoom', async () => {
	const { NativePowerPointView } = await loadNativePowerPointViewModule();
	const view = createView(NativePowerPointView);
	const originalWheelEvent = globalThis.WheelEvent;
	globalThis.WheelEvent = { DOM_DELTA_LINE: 1, DOM_DELTA_PAGE: 2 };

	try {
		const zooms = [];
		view.canvasPane = { scrollLeft: 20, scrollTop: 30 };
		view.engine = {};
		view.isActivePowerPointView = () => true;
		view.slideSurface = {};
		view.svgEl = {};
		view.setZoom = (zoom) => zooms.push(zoom);
		view.zoomLevel = 1;

		const pan = createWheelEvent({ deltaX: 120, deltaY: 80 });
		view.handleCanvasWheel(pan);

		assert.equal(view.canvasPane.scrollLeft, 140, 'horizontal two-finger scroll must pan the canvas');
		assert.equal(view.canvasPane.scrollTop, 110, 'vertical two-finger scroll must pan the canvas');
		assert.deepEqual(zooms, [], 'ordinary two-finger scroll must not zoom');
		assert.equal(pan.defaultPrevented, true);
		assert.equal(pan.propagationStopped, true);

		view.handleCanvasWheel(createWheelEvent({ ctrlKey: true, deltaY: -120 }));
		view.handleCanvasWheel(createWheelEvent({ ctrlKey: true, deltaY: 120 }));

		assert.ok(zooms[0] > 1, 'spreading fingers must zoom in');
		assert.ok(zooms[1] < 1, 'closing fingers must zoom out');
	} finally {
		if (originalWheelEvent === undefined) {
			delete globalThis.WheelEvent;
		} else {
			globalThis.WheelEvent = originalWheelEvent;
		}
	}
});

test('right-click text-box origins use the original slide point rather than the menu click', async () => {
	const { NativePowerPointView } = await loadNativePowerPointViewModule();
	const view = createView(NativePowerPointView);
	view.engine = { getSlideScale: () => 9144 };
	view.svgEl = {
		getScreenCTM: () => ({
			inverse: () => ({ id: 'inverse' }),
		}),
		createSVGPoint: () => ({
			x: 0,
			y: 0,
			matrixTransform: () => ({ x: 125.25, y: 42.5 }),
		}),
	};

	assert.deepEqual(
		view.getTextBoxInsertOrigin({ clientX: 400, clientY: 300 }),
		{ x: 1145286, y: 388620 },
	);
});
