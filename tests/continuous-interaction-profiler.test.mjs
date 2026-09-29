import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createRequire } from 'node:module';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
let modulePromise;

async function loadProfiler() {
	modulePromise ??= bundleSource('src/continuousInteractionProfiler.ts', 'continuous-interaction-profiler.cjs').then((outfile) => require(outfile));
	return modulePromise;
}

function mockProfile(resolvedRefreshHz = 60) {
	const budget = 1000 / resolvedRefreshHz;
	return {
		measuredRefreshHz: resolvedRefreshHz,
		measuredFrameBudgetMs: budget,
		resolvedRefreshHz,
		resolvedFrameBudgetMs: budget,
		sampleCount: 30,
		confidence: 'stable',
		thresholdSource: 'measured-raf',
		platform: 'desktop',
		appMode: 'desktop',
		thresholds: {
			synchronousWorkMs: 8,
			lateFrameGapMs: Math.round(budget * 1.5 * 100) / 100,
			missedFrameGapMs: Math.round(budget * 2 * 100) / 100,
			substantialStallMs: Math.round(budget * 3 * 100) / 100,
		},
	};
}

test('continuous profiler measures 60Hz scroll frames and detects late and missed frames', async () => {
	const { ContinuousInteractionProfiler } = await loadProfiler();
	let currentTime = 1000;
	const rafCallbacks = [];
	const timeoutCallbacks = [];
	const summaries = [];
	const slowAlerts = [];

	const profiler = new ContinuousInteractionProfiler({
		now: () => currentTime,
		requestFrame: (cb) => {
			rafCallbacks.push(cb);
			return rafCallbacks.length;
		},
		cancelFrame: () => {},
		scheduleTimeout: (cb, ms) => {
			timeoutCallbacks.push({ cb, triggerTime: currentTime + ms });
			return timeoutCallbacks.length;
		},
		cancelTimeout: () => {},
		getFrameProfile: () => mockProfile(60), // 16.67ms budget, late > 25ms, missed > 33.3ms, stall > 50ms
		settleTimeoutMs: 150,
		onSummary: (s) => summaries.push(s),
		onSlowInteraction: (s) => slowAlerts.push(s),
	});

	// User initiates docx scrolling
	profiler.recordInteractionEvent('docx-scroll');
	assert.equal(rafCallbacks.length, 1);

	// Frame 1: normal frame (16.6ms)
	currentTime = 1016.6;
	const frame1Cb = rafCallbacks.shift();
	profiler.measureSynchronousWork('docx-scroll', () => {
		// simulate 3ms synchronous plugin work
		currentTime += 3;
	});
	frame1Cb(1016.6);

	// Frame 2: delayed frame (35ms -> missed frame)
	assert.equal(rafCallbacks.length, 1);
	currentTime = 1051.6;
	const frame2Cb = rafCallbacks.shift();
	profiler.measureSynchronousWork('docx-scroll', () => {
		currentTime += 12; // slow sync work
	});
	frame2Cb(1051.6);

	// Frame 3: normal frame
	assert.equal(rafCallbacks.length, 1);
	currentTime = 1068.2;
	const frame3Cb = rafCallbacks.shift();
	frame3Cb(1068.2);

	// Trigger settle timeout
	currentTime = 1250;
	assert.equal(timeoutCallbacks.length > 0, true);
	const settle = timeoutCallbacks[timeoutCallbacks.length - 1];
	settle.cb();

	assert.equal(summaries.length, 1);
	const summary = summaries[0];
	assert.equal(summary.interactionType, 'docx-scroll');
	assert.equal(summary.scope, 'docx');
	assert.equal(summary.frameCount, 3);
	assert.equal(summary.lateFrameCount, 1); // frame 2 is late & missed
	assert.equal(summary.missedFrameEstimate, 1);
	assert.equal(summary.longestFrameGapMs, 35);
	assert.equal(slowAlerts.length, 1);

	// Verify worst frame records distinguish sync work vs raf delay
	assert.ok(summary.worstFrames.length > 0);
	const worst = summary.worstFrames[0];
	assert.equal(worst.intervalMs, 35);
	assert.equal(worst.synchronousWorkMs, 12);
	assert.equal(worst.classification, 'missed');
});

test('direct-manipulation summaries separate coalesced input, observable loss, and preview mutations', async () => {
	const { ContinuousInteractionProfiler } = await loadProfiler();
	let currentTime = 1000;
	const rafCallbacks = [];
	const timeoutCallbacks = [];
	const summaries = [];

	const profiler = new ContinuousInteractionProfiler({
		now: () => currentTime,
		requestFrame: (cb) => {
			rafCallbacks.push(cb);
			return rafCallbacks.length;
		},
		cancelFrame: () => {},
		scheduleTimeout: (cb, ms) => {
			timeoutCallbacks.push({ cb, triggerTime: currentTime + ms });
			return timeoutCallbacks.length;
		},
		cancelTimeout: () => {},
		getFrameProfile: () => mockProfile(60),
		settleTimeoutMs: 150,
		onSummary: (summary) => summaries.push(summary),
	});

	profiler.recordInteractionEvent('pptx-text-box-resize', { coalescedInputCount: 2 });
	profiler.recordInteractionEvent('pptx-text-box-resize', { coalescedInputCount: 3 });
	profiler.recordDomMutations('pptx-text-box-resize', 4);
	currentTime = 1016.7;
	rafCallbacks.shift()(1016.7);
	currentTime = 1200;
	timeoutCallbacks[timeoutCallbacks.length - 1].cb();

	assert.equal(summaries.length, 1);
	assert.equal(summaries[0].interactionType, 'pptx-text-box-resize');
	assert.equal(summaries[0].inputEventCount, 2);
	assert.equal(summaries[0].coalescedInputCount, 5);
	assert.equal(summaries[0].droppedInputEstimate, 0);
	assert.equal(summaries[0].domMutationCount, 4);
});

test('continuous profiler uses 120Hz frame budget correctly without false positives on 8.3ms frames', async () => {
	const { ContinuousInteractionProfiler } = await loadProfiler();
	let currentTime = 1000;
	const rafCallbacks = [];
	const timeoutCallbacks = [];
	const summaries = [];

	const profiler = new ContinuousInteractionProfiler({
		now: () => currentTime,
		requestFrame: (cb) => {
			rafCallbacks.push(cb);
			return rafCallbacks.length;
		},
		cancelFrame: () => {},
		scheduleTimeout: (cb, ms) => {
			timeoutCallbacks.push({ cb, triggerTime: currentTime + ms });
			return timeoutCallbacks.length;
		},
		cancelTimeout: () => {},
		getFrameProfile: () => mockProfile(120), // 8.33ms budget, late > 12.5ms
		settleTimeoutMs: 150,
		onSummary: (s) => summaries.push(s),
	});

	profiler.recordInteractionEvent('pptx-zoom');

	// 5 smooth frames at 8.33ms
	for (let i = 0; i < 5; i++) {
		currentTime += 8.33;
		const cb = rafCallbacks.shift();
		cb(currentTime);
	}

	currentTime += 200;
	timeoutCallbacks[timeoutCallbacks.length - 1].cb();

	assert.equal(summaries.length, 1);
	const summary = summaries[0];
	assert.equal(summary.interactionType, 'pptx-zoom');
	assert.equal(summary.lateFrameCount, 0);
	assert.equal(summary.missedFrameEstimate, 0);
	assert.equal(summary.frameIntervalsMs.p50, 8.3);
});

test('continuous profiler respects bounded history and worst frame limits', async () => {
	const { ContinuousInteractionProfiler } = await loadProfiler();
	let currentTime = 1000;
	let rafCallbacks = [];
	const timeoutCallbacks = [];

	const profiler = new ContinuousInteractionProfiler({
		now: () => currentTime,
		requestFrame: (cb) => {
			rafCallbacks.push(cb);
			return rafCallbacks.length;
		},
		cancelFrame: (h) => {
			rafCallbacks = [];
		},
		scheduleTimeout: (cb, ms) => {
			timeoutCallbacks.push({ cb, triggerTime: currentTime + ms });
			return timeoutCallbacks.length;
		},
		cancelTimeout: () => {},
		getFrameProfile: () => mockProfile(60),
		maxHistorySummaries: 3,
		maxWorstFrames: 2,
		settleTimeoutMs: 100,
	});

	for (let i = 0; i < 5; i++) {
		profiler.recordInteractionEvent('pptx-canvas-scroll');
		currentTime += 20;
		if (rafCallbacks.length > 0) {
			const cb = rafCallbacks.shift();
			cb(currentTime);
		}
		currentTime += 150;
		if (timeoutCallbacks.length > 0) {
			const timeout = timeoutCallbacks.shift();
			timeout.cb();
		}
	}

	const history = profiler.getRecentSummaries();
	assert.equal(history.length, 3, 'capped at maxHistorySummaries');
	const latest = profiler.getLatestSummary('pptx-canvas-scroll');
	assert.ok(latest);
	assert.ok(latest.worstFrames.length <= 2, 'capped at maxWorstFrames');
});
