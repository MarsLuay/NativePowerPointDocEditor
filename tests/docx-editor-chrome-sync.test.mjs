import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import test from 'node:test';
import { JSDOM } from 'jsdom';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
const {
	createEditorChromeReconciler,
	observeEditorChromeAttributes,
} = require(await bundleSource('src/docxEditorChromeSync.ts', 'docx-editor-chrome-sync.cjs'));

function createFrames() {
	let nextId = 1;
	const pending = new Map();
	return {
		requestFrame(callback) {
			const id = nextId++;
			pending.set(id, callback);
			return id;
		},
		cancelFrame(id) {
			pending.delete(id);
		},
		get size() {
			return pending.size;
		},
		flushAll(max = 30) {
			let ran = 0;
			while (ran < max) {
				const id = pending.keys().next().value;
				if (id === undefined) {
					break;
				}
				const callback = pending.get(id);
				pending.delete(id);
				callback();
				ran += 1;
			}
			return ran;
		},
	};
}

function createHost(html) {
	const dom = new JSDOM(`<!doctype html><div id="host">${html}</div>`);
	return dom.window.document.getElementById('host');
}

function settleDom() {
	return new Promise((resolve) => {
		setImmediate(resolve);
	});
}

test('one vendor title mutation reconciles and then settles', async () => {
	const host = createHost('<button title="Bold"></button>');
	const button = host.querySelector('button');
	const frames = createFrames();
	let syncCount = 0;
	createEditorChromeReconciler({
		observe: (listener) => observeEditorChromeAttributes(host, () => listener()),
		syncTarget: host,
		requestFrame: frames.requestFrame,
		cancelFrame: frames.cancelFrame,
		now: () => 0,
		sync() {
			syncCount += 1;
			const title = button.getAttribute('title');
			if (title && button.getAttribute('aria-label') !== title) {
				button.setAttribute('aria-label', title);
			}
			if (button.hasAttribute('title')) {
				button.removeAttribute('title');
			}
		},
	});

	button.setAttribute('title', 'Italic');
	await settleDom();
	const framesRun = frames.flushAll();
	await settleDom();
	const extraFrames = frames.flushAll();

	assert.equal(button.getAttribute('title'), null);
	assert.equal(button.getAttribute('aria-label'), 'Italic');
	assert.equal(syncCount, 1);
	assert.equal(framesRun, 1);
	assert.equal(extraFrames, 0);
});

test('one vendor aria-label mutation reconciles and then settles', async () => {
	const host = createHost('<button aria-label="Bold"></button>');
	const button = host.querySelector('button');
	const frames = createFrames();
	let syncCount = 0;
	createEditorChromeReconciler({
		observe: (listener) => observeEditorChromeAttributes(host, () => listener()),
		syncTarget: host,
		requestFrame: frames.requestFrame,
		cancelFrame: frames.cancelFrame,
		now: () => 0,
		sync() {
			syncCount += 1;
			if (button.getAttribute('aria-label') === 'vendor') {
				button.setAttribute('aria-label', 'normalized');
			}
		},
	});

	button.setAttribute('aria-label', 'vendor');
	await settleDom();
	frames.flushAll();
	await settleDom();
	const extraFrames = frames.flushAll();

	assert.equal(button.getAttribute('aria-label'), 'normalized');
	assert.equal(syncCount, 1);
	assert.equal(extraFrames, 0);
});

test('plugin-authored chrome normalization does not schedule another sync', async () => {
	const host = createHost('<button></button>');
	const button = host.querySelector('button');
	const frames = createFrames();
	let syncCount = 0;
	const reconciler = createEditorChromeReconciler({
		observe: (listener) => observeEditorChromeAttributes(host, () => listener()),
		syncTarget: host,
		requestFrame: frames.requestFrame,
		cancelFrame: frames.cancelFrame,
		now: () => 0,
		sync() {
			syncCount += 1;
			const next = syncCount % 2 === 0 ? 'normalized-even' : 'normalized-odd';
			button.setAttribute('aria-label', next);
		},
	});

	reconciler.schedule();
	const framesRun = frames.flushAll();
	await settleDom();
	const extraFrames = frames.flushAll();

	assert.equal(syncCount, 1);
	assert.equal(framesRun, 1);
	assert.equal(extraFrames, 0);
	assert.equal(button.getAttribute('aria-label'), 'normalized-odd');
});

test('unchanged chrome does not schedule repeated animation frames', async () => {
	const host = createHost('<button aria-label="Bold"></button>');
	const frames = createFrames();
	let syncCount = 0;
	const reconciler = createEditorChromeReconciler({
		observe: (listener) => observeEditorChromeAttributes(host, () => listener()),
		syncTarget: host,
		requestFrame: frames.requestFrame,
		cancelFrame: frames.cancelFrame,
		now: () => 0,
		sync() {
			syncCount += 1;
		},
	});

	reconciler.schedule();
	assert.equal(frames.flushAll(), 1);
	await settleDom();

	assert.equal(syncCount, 1);
	assert.equal(frames.size, 0);
});

test('chrome observer teardown cancels a queued sync', async () => {
	const host = createHost('<button title="Bold"></button>');
	const frames = createFrames();
	let syncCount = 0;
	const reconciler = createEditorChromeReconciler({
		observe: (listener) => observeEditorChromeAttributes(host, () => listener()),
		requestFrame: frames.requestFrame,
		cancelFrame: frames.cancelFrame,
		sync() {
			syncCount += 1;
		},
	});

	reconciler.schedule();
	assert.equal(frames.size, 1);
	reconciler.dispose();
	assert.equal(frames.flushAll(), 0);
	assert.equal(syncCount, 0);

	host.querySelector('button').setAttribute('title', 'Italic');
	await settleDom();
	assert.equal(frames.flushAll(), 0);
	assert.equal(syncCount, 0);
});

test('a chrome feedback storm logs one summary instead of every frame', () => {
	const frames = createFrames();
	const storms = [];
	const reconciler = createEditorChromeReconciler({
		stormThreshold: 4,
		now: () => 1000,
		onStorm: (summary) => {
			storms.push(summary);
		},
		observe(listener) {
			return {
				suspend() {},
				resume() {
					listener();
				},
				dispose() {},
			};
		},
		requestFrame: frames.requestFrame,
		cancelFrame: frames.cancelFrame,
		sync() {
			return true;
		},
	});

	reconciler.schedule();
	const framesRun = frames.flushAll(12);

	assert.equal(storms.length, 1);
	assert.ok(framesRun > 1);
	assert.ok(storms[0].syncCount >= 4 || storms[0].observerCallbacks >= 4);
	assert.ok(storms[0].syncCount >= 1);
	assert.ok(storms[0].observerCallbacks >= 1);
	assert.equal(storms[0].lastSyncChangedDom, true);
	assert.deepEqual([...storms[0].attributeNames], ['title', 'aria-label']);
	assert.equal(typeof storms[0].elapsedMs, 'number');
	assert.equal(typeof storms[0].pluginAuthoredMutations, 'number');
});
