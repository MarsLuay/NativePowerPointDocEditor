import assert from "node:assert/strict";
import { afterEach, test } from "node:test";
import { loadIdleScheduleModule } from "./helpers/load-plugin-modules.mjs";

const originalWindow = globalThis.window;

afterEach(() => {
	globalThis.window = originalWindow;
});

test("scheduleIdleWork uses requestIdleCallback when available", async () => {
	const { scheduleIdleWork } = await loadIdleScheduleModule();

	let callbackCalled = false;
	let requestedTimeout = -1;
	let cancelledId = -1;
	let callbackFn = null;

	globalThis.window = {
		requestIdleCallback: (cb, opts) => {
			callbackFn = cb;
			requestedTimeout = opts?.timeout;
			return 123;
		},
		cancelIdleCallback: (id) => {
			cancelledId = id;
		}
	};

	const cancel = scheduleIdleWork(() => {
		callbackCalled = true;
	});

	assert.equal(requestedTimeout, 2000, "uses default timeout of 2000");
	assert.equal(typeof callbackFn, "function", "passed a callback function");

	callbackFn();
	assert.equal(callbackCalled, true, "callback is invoked");

	cancel();
	assert.equal(cancelledId, 123, "cancel returns correctly");
});

test("scheduleIdleWork uses custom timeout for requestIdleCallback", async () => {
	const { scheduleIdleWork } = await loadIdleScheduleModule();

	let requestedTimeout = -1;

	globalThis.window = {
		requestIdleCallback: (cb, opts) => {
			requestedTimeout = opts?.timeout;
			return 456;
		},
		cancelIdleCallback: () => {}
	};

	scheduleIdleWork(() => {}, { timeout: 5000 });

	assert.equal(requestedTimeout, 5000, "uses custom timeout");
});

test("scheduleIdleWork falls back to setTimeout when requestIdleCallback is unavailable", async () => {
	const { scheduleIdleWork } = await loadIdleScheduleModule();

	let timeoutMs = -1;
	let clearedId = -1;
	let callbackFn = null;

	globalThis.window = {
		requestIdleCallback: undefined,
		setTimeout: (cb, timeout) => {
			callbackFn = cb;
			timeoutMs = timeout;
			return 789;
		},
		clearTimeout: (id) => {
			clearedId = id;
		}
	};

	let callbackCalled = false;
	const cancel = scheduleIdleWork(() => {
		callbackCalled = true;
	}, { timeout: 1000 });

	assert.equal(timeoutMs, 250, "uses Math.min(timeout, 250) for fallback timeout");
	assert.equal(typeof callbackFn, "function", "passed a callback function");

	callbackFn();
	assert.equal(callbackCalled, true, "callback is invoked");

	cancel();
	assert.equal(clearedId, 789, "cancel returns correctly via clearTimeout");
});

test("scheduleIdleWork falls back to setTimeout when requestIdleCallback is unavailable with smaller timeout", async () => {
	const { scheduleIdleWork } = await loadIdleScheduleModule();

	let timeoutMs = -1;

	globalThis.window = {
		requestIdleCallback: undefined,
		setTimeout: (cb, timeout) => {
			timeoutMs = timeout;
			return 101;
		},
		clearTimeout: () => {}
	};

	scheduleIdleWork(() => {}, { timeout: 100 });

	assert.equal(timeoutMs, 100, "uses timeout when it is smaller than 250");
});
