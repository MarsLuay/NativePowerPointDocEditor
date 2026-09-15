import assert from "node:assert/strict";
import { test } from "node:test";
import { loadLoadTraceModule, loadLoggerModule } from "./helpers/load-plugin-modules.mjs";

test("monotonicNow prefers performance.now when available", async () => {
	const originalPerformance = globalThis.performance;
	let mockTime = 1000;

	try {
		globalThis.performance = {
			now: () => mockTime
		};

		const { monotonicNow } = await loadLoadTraceModule();
		assert.equal(monotonicNow(), 1000);
	} finally {
		globalThis.performance = originalPerformance;
	}
});

test("monotonicNow falls back to Date.now when performance is unavailable", async () => {
	const originalPerformance = globalThis.performance;
	const originalDateNow = Date.now;

	try {
		globalThis.performance = undefined;
		Date.now = () => 2000;

		const { monotonicNow } = await loadLoadTraceModule();

		assert.equal(monotonicNow(), 2000);
	} finally {
		globalThis.performance = originalPerformance;
		Date.now = originalDateNow;
	}
});

test("createLoadTrace records phases with relative timings", async () => {
	const { configureNativePowerPointDocEditorLogger } = await loadLoggerModule();
	globalThis.window = {
		nativePowerPointDocEditorDebugLogging: true,
		nativePowerPointDocEditorDebugLogs: []
	};
	configureNativePowerPointDocEditorLogger(true);

	const originalPerformanceNow = globalThis.performance?.now;
	let mockTime = 1000;
	if (!globalThis.performance) {
		globalThis.performance = {};
	}
	globalThis.performance.now = () => mockTime;

	const { createLoadTrace } = await loadLoadTraceModule();

	const originalWarn = console.warn;
	const originalDebug = console.debug;
	const warnings = [];
	const debugs = [];
	console.warn = (...args) => warnings.push(args);
	console.debug = (...args) => debugs.push(args);

	try {
		const trace = createLoadTrace('test-scope', { contextKey: 'contextValue' });

		// Phase 1 (fast phase)
		mockTime = 1050; // +50ms
		trace.mark('phase1', { dataKey: 'data1' });

		assert.equal(warnings.length, 1);
		assert.match(warnings[0][0], /load: test-scope: phase1/);
		assert.equal(warnings[0][1].sinceStartMs, 50);
		assert.equal(warnings[0][1].sincePreviousMs, 50);

		assert.equal(debugs.length, 1);
		assert.match(debugs[0][0], /load: test-scope: phase1/);
		assert.equal(debugs[0][1].sinceStartMs, 50);

		// Phase 2 (slow phase, triggers warnLog)
		mockTime = 1300; // +250ms
		trace.mark('phase2', { dataKey: 'data2' });

		assert.equal(warnings.length, 3); // Two marks + 1 logger warning output
		assert.match(warnings[1][0], /load: test-scope: phase2/);
		assert.equal(warnings[1][1].sinceStartMs, 300);
		assert.equal(warnings[1][1].sincePreviousMs, 250);

		assert.match(warnings[2][0], /load: test-scope: slow phase phase2/);
		assert.equal(warnings[2][1].sinceStartMs, 300);
		assert.equal(warnings[2][1].sincePreviousMs, 250);

		assert.equal(debugs.length, 2);
		assert.match(debugs[1][0], /load: test-scope: phase2/);

		// Finish
		mockTime = 1400; // +100ms
		trace.finish('test-finish', { dataKey: 'data3' });

		assert.equal(warnings.length, 4); // 3 marks + 1 logger warning output
		assert.match(warnings[3][0], /load: test-scope: test-finish/);
		assert.equal(warnings[3][1].totalMs, 400);
		assert.equal(warnings[3][1].phases.length, 2);
		assert.deepEqual(warnings[3][1].phases[0], { phase: 'phase1', sinceStartMs: 50, sincePreviousMs: 50 });

		assert.equal(debugs.length, 3);
		assert.match(debugs[2][0], /load: test-scope: test-finish/);

	} finally {
		if (originalPerformanceNow !== undefined) {
			globalThis.performance.now = originalPerformanceNow;
		} else {
			globalThis.performance = undefined;
		}
		console.warn = originalWarn;
		console.debug = originalDebug;
		globalThis.window = undefined;
	}
});

test("createLoadTrace logs warnings for slow loads", async () => {
	const originalPerformanceNow = globalThis.performance?.now;
	let mockTime = 1000;
	if (!globalThis.performance) {
		globalThis.performance = {};
	}
	globalThis.performance.now = () => mockTime;

	const { createLoadTrace } = await loadLoadTraceModule();

	const originalWarn = console.warn;
	const originalDebug = console.debug;
	const warnings = [];
	console.warn = (...args) => warnings.push(args);
	console.debug = () => {};

	try {
		const trace = createLoadTrace('slow-scope');

		// Slow load (>1500ms)
		mockTime = 2600; // +1600ms
		trace.finish('done');

		assert.equal(warnings.length, 2); // 1 mark + 1 logger warning output
		assert.match(warnings[0][0], /load: slow-scope: done/);

		assert.match(warnings[1][0], /load: slow-scope: slow load/);
		assert.equal(warnings[1][1].totalMs, 1600);

	} finally {
		if (originalPerformanceNow !== undefined) {
			globalThis.performance.now = originalPerformanceNow;
		} else {
			globalThis.performance = undefined;
		}
		console.warn = originalWarn;
		console.debug = originalDebug;
	}
});
