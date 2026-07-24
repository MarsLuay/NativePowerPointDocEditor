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
	const hadWindow = Object.prototype.hasOwnProperty.call(globalThis, 'window');
	const originalWindow = hadWindow ? globalThis.window : undefined;
	const rafQueue = [];
	const rafHost = {
		requestAnimationFrame: (cb) => {
			rafQueue.push(cb);
			return rafQueue.length;
		},
		cancelAnimationFrame: (id) => {
			rafQueue[id - 1] = null;
		},
	};
	globalThis.WheelEvent = { DOM_DELTA_LINE: 1, DOM_DELTA_PAGE: 2 };
	// View prefers window/activeWindow for RAF (popout-safe); Node has neither by default.
	globalThis.window = rafHost;

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

		// Two-finger scroll uses native overflow scrolling (no preventDefault /
		// JS scroll mutation). Handler only stops propagation so Obsidian
		// workspace chrome does not steal the gesture.
		assert.equal(view.canvasPane.scrollLeft, 20, 'pan handler must not mutate scrollLeft');
		assert.equal(view.canvasPane.scrollTop, 30, 'pan handler must not mutate scrollTop');
		assert.deepEqual(zooms, [], 'ordinary two-finger scroll must not zoom');
		assert.equal(pan.defaultPrevented, false);
		assert.equal(pan.propagationStopped, true);

		view.handleCanvasWheel(createWheelEvent({ ctrlKey: true, deltaY: -120 }));
		view.handleCanvasWheel(createWheelEvent({ ctrlKey: true, deltaY: 120 }));
		assert.deepEqual(zooms, [], 'pinch must wait for animation frame coalesce');
		for (const cb of rafQueue.splice(0)) {
			if (cb) cb(0);
		}
		assert.equal(zooms.length, 1, 'pinch wheel ticks must coalesce to one zoom apply');
		// -120 then +120 from base 1 → net factor 2^(-120/600)*2^(120/600) = 1
		// but second tick chains from pending zoom of first, so final ≈ 1.
		assert.ok(Math.abs(zooms[0] - 1) < 0.001, 'equal in/out pinch should net near 1x');

		zooms.length = 0;
		view.zoomLevel = 1;
		view.handleCanvasWheel(createWheelEvent({ ctrlKey: true, deltaY: -120 }));
		for (const cb of rafQueue.splice(0)) {
			if (cb) cb(0);
		}
		assert.ok(zooms[0] > 1, 'spreading fingers must zoom in');

		zooms.length = 0;
		view.zoomLevel = 1;
		view.handleCanvasWheel(createWheelEvent({ ctrlKey: true, deltaY: 120 }));
		for (const cb of rafQueue.splice(0)) {
			if (cb) cb(0);
		}
		assert.ok(zooms[0] < 1, 'closing fingers must zoom out');
	} finally {
		if (originalWheelEvent === undefined) {
			delete globalThis.WheelEvent;
		} else {
			globalThis.WheelEvent = originalWheelEvent;
		}
		if (hadWindow) {
			globalThis.window = originalWindow;
		} else {
			delete globalThis.window;
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
