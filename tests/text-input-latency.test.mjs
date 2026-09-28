import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createRequire } from 'node:module';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
let modulePromise;

async function loadTracker() {
	modulePromise ??= bundleSource('src/textInputLatency.ts', 'text-input-latency.cjs').then((outfile) => require(outfile));
	return modulePromise;
}

function trustedKey(key = 'a') {
	return { isTrusted: true, key };
}

function trustedInput(inputType = 'insertText') {
	return { isTrusted: true, inputType };
}

test('typing latency records separate model and frame delays without per-key logs', async () => {
	const { createTextInputLatencyTracker } = await loadTracker();
	let now = 0;
	const frames = [];
	const summaries = [];
	const tracker = createTextInputLatencyTracker({
		scope: 'docx',
		now: () => now,
		requestFrame: (callback) => { frames.push(callback); return frames.length; },
		onSummary: (summary) => summaries.push(summary),
	});

	const correlationId = tracker.begin(trustedKey(), 'keydown');
	assert.ok(correlationId);
	now = 4;
	assert.equal(tracker.begin(trustedInput(), 'beforeinput'), correlationId);
	now = 7;
	tracker.markModelUpdated();
	assert.equal(frames.length, 1);
	frames.shift()(16.67);

	const summary = tracker.getSummary();
	assert.equal(summary.scope, 'docx');
	assert.equal(summary.sampleCount, 1);
	assert.equal(summary.slowInteractionCount, 0);
	assert.equal(summary.inputToModelMs.p50, 7);
	assert.equal(summary.frameSchedulingDelayMs.p50, 9.7);
	assert.equal(summary.inputToVisibleMs.p50, 16.7);
	assert.equal(summaries.length, 0, 'summaries are retained rather than emitted per keystroke');
});

test('typing latency rejects untrusted events and retains slow correlation data only', async () => {
	const { createTextInputLatencyTracker } = await loadTracker();
	let now = 0;
	const frames = [];
	const slow = [];
	const tracker = createTextInputLatencyTracker({
		scope: 'pptx',
		now: () => now,
		requestFrame: (callback) => { frames.push(callback); return frames.length; },
		onSlowInteraction: (entry) => slow.push(entry),
	});

	assert.equal(tracker.begin({ isTrusted: false, key: 'a' }, 'keydown'), null);
	assert.equal(tracker.begin({ isTrusted: true, key: 'ArrowLeft' }, 'keydown'), null);
	const correlationId = tracker.begin(trustedInput('deleteContentBackward'), 'beforeinput');
	assert.ok(correlationId);
	now = 12;
	tracker.markModelUpdated(correlationId);
	frames.shift()(55);

	assert.equal(tracker.getSummary().slowInteractionCount, 1);
	assert.deepEqual(slow[0], {
		correlationId,
		source: 'beforeinput',
		inputToModelMs: 12,
		frameSchedulingDelayMs: 43,
		inputToVisibleMs: 55,
		slowThresholdMs: 50,
	});
});

test('typing latency summarizes bounded rapid input samples', async () => {
	const { createTextInputLatencyTracker } = await loadTracker();
	let now = 0;
	const frames = [];
	const tracker = createTextInputLatencyTracker({
		scope: 'docx',
		now: () => now,
		requestFrame: (callback) => { frames.push(callback); return frames.length; },
	});

	for (let index = 0; index < 140; index += 1) {
		const id = tracker.begin(trustedInput(index % 2 ? 'insertText' : 'insertFromPaste'), 'beforeinput');
		now += 2;
		tracker.markModelUpdated(id);
		frames.shift()(now + 8);
		now += 10;
	}
	const summary = tracker.getSummary();
	assert.equal(summary.sampleCount, 128);
	assert.equal(summary.inputToModelMs.p50, 2);
	assert.equal(summary.frameSchedulingDelayMs.p95, 8);
	assert.equal(summary.inputToVisibleMs.max, 10);
});
