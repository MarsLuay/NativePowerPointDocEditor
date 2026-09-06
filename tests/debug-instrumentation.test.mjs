import assert from "node:assert/strict";
import { test, afterEach } from "node:test";
import { loadDebugInstrumentationModule } from "./helpers/load-plugin-modules.mjs";

afterEach(() => {
	globalThis.window = undefined;
	globalThis.performance = undefined;
	globalThis.MutationObserver = undefined;
	console.warn = originalWarn;
});

const originalWarn = console.warn;

test("logLifecycleStep logs to debugLog and console.warn", async () => {
	globalThis.window = {
		nativePowerPointDocEditorDebugLogging: true,
		nativePowerPointDocEditorDebugLogs: [],
	};
	const { logLifecycleStep } = await loadDebugInstrumentationModule();

	const warnings = [];
	console.warn = (...args) => warnings.push(args);

	logLifecycleStep("test-step", { foo: "bar" });

	const logs = globalThis.window.nativePowerPointDocEditorDebugLogs;
	const debugEntry = logs.find(log => log.message === "test-step" && log.area === "lifecycle");
	assert.ok(debugEntry);
	assert.deepEqual(debugEntry.data, { step: "test-step", foo: "bar" });

	assert.equal(warnings.length, 1);
	assert.equal(warnings[0][0], "[Native PowerPoint Doc Editor] lifecycle: test-step");
	assert.deepEqual(warnings[0][1], { step: "test-step", foo: "bar" });
});

test("traceSyncStep measures duration and logs start/done", async () => {
	globalThis.window = {
		nativePowerPointDocEditorDebugLogging: true,
		nativePowerPointDocEditorDebugLogs: [],
	};
	const { traceSyncStep } = await loadDebugInstrumentationModule();

	console.warn = () => {};

	let now = 1000;
	globalThis.performance = { now: () => now };

	traceSyncStep("sync-step", () => {
		now += 150.5; // Simulate 150.5ms work
		return "result";
	});

	const logs = globalThis.window.nativePowerPointDocEditorDebugLogs;

	const startEntry = logs.find(log => log.message === "sync-step:start");
	assert.ok(startEntry);
	assert.equal(startEntry.level, "debug");

	const doneEntry = logs.find(log => log.message === "sync-step:done");
	assert.ok(doneEntry);
	assert.equal(doneEntry.level, "debug");
	assert.equal(doneEntry.data.durationMs, 150.5);

	const warnEntry = logs.find(log => log.message === "Slow sync step: sync-step");
	assert.equal(warnEntry, undefined);
});

test("createObservedMutationObserver logs on high activity", async () => {
	globalThis.window = {
		nativePowerPointDocEditorDebugLogging: true,
		nativePowerPointDocEditorDebugLogs: [],
	};
    globalThis.MutationObserver = class MutationObserver {
        constructor(callback) {
            this.callback = callback;
        }
        observe() {}
        disconnect() {}
    };

	const { createObservedMutationObserver } = await loadDebugInstrumentationModule();

	const warnings = [];
	console.warn = (...args) => warnings.push(args);

	let now = 1000;
	globalThis.performance = { now: () => now };

    let callbackCalled = 0;
	const observer = createObservedMutationObserver("test-observer", () => {
        callbackCalled++;
    });

    // Fire 20 invocations within 1000ms
    for(let i=0; i<19; i++) {
        now += 10;
        observer.callback([], observer);
    }

    // 20th invocation triggers log check. Elapsed = 1000.
    now = 2000;
    observer.callback([], observer);

	const logs = globalThis.window.nativePowerPointDocEditorDebugLogs;
	const warnEntry = logs.find(log => log.message === "High mutation activity: test-observer");
	assert.ok(warnEntry);
    assert.equal(warnEntry.level, "warn");
	assert.equal(warnEntry.data.invocationCount, 20);

    const stormWarning = warnings.find(args => args[0] === "[Native PowerPoint Doc Editor] observer storm: test-observer");
	assert.ok(stormWarning);
	assert.deepEqual(stormWarning[1], { mutationCount: 0, invocationCount: 20 });
    assert.equal(callbackCalled, 20);
});


test("startOpenHeartbeat logs heartbeat repeatedly and can be cleared", async () => {
	globalThis.window = {
		nativePowerPointDocEditorDebugLogging: true,
		nativePowerPointDocEditorDebugLogs: [],
	};

	let now = 1000;
	globalThis.performance = { now: () => now };

	const timers = new Map();
	let timerId = 0;
	globalThis.window.setInterval = (cb, ms) => {
		timerId++;
		timers.set(timerId, cb);
		return timerId;
	};
	globalThis.window.clearInterval = (id) => {
		timers.delete(id);
	};

	const { startOpenHeartbeat } = await loadDebugInstrumentationModule();

	const warnings = [];
	console.warn = (...args) => warnings.push(args);

	const clear = startOpenHeartbeat("test-scope", () => ({ extra: "data" }));

	assert.equal(timers.size, 1);
	const cb = timers.get(timerId);

	now = 3000;
	cb();

	const logs = globalThis.window.nativePowerPointDocEditorDebugLogs;
	const warnEntry = logs.find(log => log.message === "test-scope: heartbeat");
	assert.ok(warnEntry);
	assert.equal(warnEntry.level, "warn");
	assert.deepEqual(warnEntry.data, { scope: "test-scope", elapsedMs: 2000, extra: "data" });

	const heartbeatWarning = warnings.find(args => args[0] === "[Native PowerPoint Doc Editor] load heartbeat: test-scope");
	assert.ok(heartbeatWarning);
	assert.deepEqual(heartbeatWarning[1], { scope: "test-scope", elapsedMs: 2000, extra: "data" });

	clear();
	assert.equal(timers.size, 0);
});

test("traceSyncStep logs warning if duration > 250", async () => {
	globalThis.window = {
		nativePowerPointDocEditorDebugLogging: true,
		nativePowerPointDocEditorDebugLogs: [],
	};
	const { traceSyncStep } = await loadDebugInstrumentationModule();

	console.warn = () => {};

	let now = 1000;
	globalThis.performance = { now: () => now };

	traceSyncStep("slow-sync-step", () => {
		now += 300; // Simulate 300ms work
		return "result";
	});

	const logs = globalThis.window.nativePowerPointDocEditorDebugLogs;
	const warnEntry = logs.find(log => log.message === "Slow sync step: slow-sync-step");
	assert.ok(warnEntry);
    assert.equal(warnEntry.level, "warn");
});
