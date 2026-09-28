import assert from 'node:assert/strict';
import { test } from 'node:test';
import { bundleSource } from './helpers/load-plugin-modules.mjs';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
let modulePromise;

async function loadProfilerModule() {
	modulePromise ??= bundleSource(
		'src/runtimeFrameProfiler.ts',
		'runtime-frame-profiler.cjs',
	).then((outfile) => require(outfile));
	return modulePromise;
}

function observeCadence(profiler, interval, count = 40) {
	let timestamp = 0;
	profiler.observeFrame(timestamp);
	for (let index = 0; index < count; index += 1) {
		timestamp += interval;
		profiler.observeFrame(timestamp);
	}
	return timestamp;
}

test('runtime frame profiler resolves 60 Hz, 120 Hz, and 144 Hz rAF cadences', async () => {
	const { RuntimeFrameProfiler } = await loadProfilerModule();
	for (const [interval, expectedHz] of [
		[1000 / 60, 60],
		[1000 / 120, 120],
		[1000 / 144, 144],
	]) {
		const profiler = new RuntimeFrameProfiler({ isVisible: () => true });
		observeCadence(profiler, interval);
		const profile = profiler.getProfile();
		assert.equal(profile.thresholdSource, 'measured-raf');
		assert.equal(profile.confidence, 'stable');
		assert.ok(Math.abs(profile.measuredRefreshHz - expectedHz) < 0.2, `${expectedHz} Hz cadence should be measured`);
		assert.ok(Math.abs(profile.measuredFrameBudgetMs - interval) < 0.02, `${expectedHz} Hz frame budget should be measured`);
	}
});

test('runtime frame profiler excludes stalled frames from the cadence baseline', async () => {
	const { RuntimeFrameProfiler } = await loadProfilerModule();
	const profiler = new RuntimeFrameProfiler({ isVisible: () => true });
	let timestamp = observeCadence(profiler, 1000 / 120, 20);
	timestamp += 250;
	profiler.observeFrame(timestamp);
	for (let index = 0; index < 20; index += 1) {
		timestamp += 1000 / 120;
		profiler.observeFrame(timestamp);
	}
	const profile = profiler.getProfile();
	assert.equal(profile.thresholdSource, 'measured-raf');
	assert.ok(Math.abs(profile.measuredRefreshHz - 120) < 0.2);
});

test('runtime frame profiler falls back until enough visible samples are available', async () => {
	const { RuntimeFrameProfiler } = await loadProfilerModule();
	let visible = true;
	const profiler = new RuntimeFrameProfiler({
		appMode: 'mobile',
		platform: 'ios',
		isVisible: () => visible,
	});
	observeCadence(profiler, 1000 / 120, 4);
	let profile = profiler.getProfile();
	assert.equal(profile.thresholdSource, 'platform-fallback');
	assert.equal(profile.measuredRefreshHz, null);
	assert.equal(profile.resolvedRefreshHz, 60);
	assert.equal(profile.appMode, 'mobile');
	assert.equal(profile.platform, 'ios');

	visible = false;
	profiler.observeFrame(1000);
	visible = true;
	observeCadence(profiler, 1000 / 60, 40);
	profile = profiler.getProfile();
	assert.equal(profile.thresholdSource, 'measured-raf');
	assert.ok(Math.abs(profile.measuredRefreshHz - 60) < 0.2);
});

test('frame-gap thresholds scale from the measured frame budget while synchronous work remains strict', async () => {
	const { RuntimeFrameProfiler, SYNCHRONOUS_WORK_THRESHOLD_MS } = await loadProfilerModule();
	const sixty = new RuntimeFrameProfiler({ isVisible: () => true });
	observeCadence(sixty, 1000 / 60);
	const oneTwenty = new RuntimeFrameProfiler({ isVisible: () => true });
	observeCadence(oneTwenty, 1000 / 120);

	assert.equal(sixty.getThresholds().synchronousWorkMs, SYNCHRONOUS_WORK_THRESHOLD_MS);
	assert.equal(oneTwenty.getThresholds().synchronousWorkMs, SYNCHRONOUS_WORK_THRESHOLD_MS);
	assert.ok(sixty.getThresholds().lateFrameGapMs > oneTwenty.getThresholds().lateFrameGapMs * 1.9);
	assert.ok(sixty.getThresholds().missedFrameGapMs > oneTwenty.getThresholds().missedFrameGapMs * 1.9);
	assert.ok(sixty.getThresholds().substantialStallMs > oneTwenty.getThresholds().substantialStallMs * 1.9);
});

test('runtime frame profiler bounds retained samples and profile-change notifications', async () => {
	const { RuntimeFrameProfiler } = await loadProfilerModule();
	const changes = [];
	const profiler = new RuntimeFrameProfiler({
		isVisible: () => true,
		minimumSamples: 4,
		stableSamples: 8,
		maxSamples: 16,
		onProfileChange: (profile) => changes.push(profile),
	});
	observeCadence(profiler, 1000 / 60, 80);

	assert.ok(profiler.getProfile().sampleCount <= 16);
	assert.equal(profiler.getProfile().confidence, 'stable');
	assert.ok(changes.length <= 2, 'stable cadence must not emit a per-frame update');
});

test('shared rAF helpers preserve a view-specific window for popout DOCX views', async () => {
	const { startRuntimeFrameProfiler, stopRuntimeFrameProfiler, requestRuntimeFrame, cancelRuntimeFrame } = await loadProfilerModule();
	let nextHandle = 1;
	const sampledWindow = {
		requestAnimationFrame: () => nextHandle++,
		cancelAnimationFrame: () => {},
	};
	const ownerRequested = [];
	const ownerCancelled = [];
	const ownerWindow = {
		requestAnimationFrame: (callback) => {
			ownerRequested.push(callback);
			return 41;
		},
		cancelAnimationFrame: (handle) => ownerCancelled.push(handle),
	};

	startRuntimeFrameProfiler({
		requestFrame: sampledWindow.requestAnimationFrame,
		cancelFrame: sampledWindow.cancelAnimationFrame,
	});
	try {
		const handle = requestRuntimeFrame(() => {}, ownerWindow);
		assert.equal(handle, 41);
		assert.equal(ownerRequested.length, 1);
		cancelRuntimeFrame(handle, ownerWindow);
		assert.deepEqual(ownerCancelled, [41]);
	} finally {
		stopRuntimeFrameProfiler();
	}
});
