import assert from "node:assert/strict";
import { test } from "node:test";
import { build } from "esbuild";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createRequire } from "node:module";

const projectRoot = path.resolve(import.meta.dirname, "..");
const require = createRequire(import.meta.url);

let cachedModule;

async function loadCoalescedTimeoutModule() {
	if (cachedModule) return cachedModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), "npde-coalesced-timeout-"));
	const outfile = path.join(outputDirectory, "coalesced-timeout.cjs");
	await build({
		entryPoints: [path.join(projectRoot, "src/coalescedTimeout.ts")],
		bundle: true,
		format: "cjs",
		logLevel: "silent",
		outfile,
		platform: "node",
		target: "node22",
	});
	cachedModule = require(outfile);
	return cachedModule;
}

function createMockTimeoutView() {
	let nextId = 1;
	const timers = new Map();

	return {
		timers,
		setTimeout(handler, timeout) {
			const id = nextId++;
			timers.set(id, { handler, timeout });
			return id;
		},
		clearTimeout(id) {
			timers.delete(id);
		},
		fire(id) {
			const timer = timers.get(id);
			if (timer) {
				timers.delete(id);
				timer.handler();
			}
		},
	};
}

test("schedule registers a timeout and executes task when timer fires", async () => {
	const { CoalescedTimeout } = await loadCoalescedTimeoutModule();
	const mockView = createMockTimeoutView();
	let taskRunCount = 0;

	const timeout = new CoalescedTimeout(mockView, () => {
		taskRunCount += 1;
	});

	timeout.schedule(100);

	assert.equal(mockView.timers.size, 1);
	const [timerId, timerInfo] = Array.from(mockView.timers.entries())[0];
	assert.equal(timerInfo.timeout, 100);
	assert.equal(taskRunCount, 0);

	mockView.fire(timerId);
	assert.equal(taskRunCount, 1);
	assert.equal(mockView.timers.size, 0);
});

test("schedule coalesces duplicate calls when timer is already pending", async () => {
	const { CoalescedTimeout } = await loadCoalescedTimeoutModule();
	const mockView = createMockTimeoutView();
	let taskRunCount = 0;

	const timeout = new CoalescedTimeout(mockView, () => {
		taskRunCount += 1;
	});

	timeout.schedule(100);
	timeout.schedule(200);
	timeout.schedule(300);

	assert.equal(mockView.timers.size, 1);
	const [timerId, timerInfo] = Array.from(mockView.timers.entries())[0];
	assert.equal(timerInfo.timeout, 100);

	mockView.fire(timerId);
	assert.equal(taskRunCount, 1);
});

test("allows rescheduling after the previous timeout has executed", async () => {
	const { CoalescedTimeout } = await loadCoalescedTimeoutModule();
	const mockView = createMockTimeoutView();
	let taskRunCount = 0;

	const timeout = new CoalescedTimeout(mockView, () => {
		taskRunCount += 1;
	});

	timeout.schedule(100);
	const [firstId] = Array.from(mockView.timers.keys());
	mockView.fire(firstId);
	assert.equal(taskRunCount, 1);

	timeout.schedule(200);
	assert.equal(mockView.timers.size, 1);
	const [secondId, timerInfo] = Array.from(mockView.timers.entries())[0];
	assert.notEqual(firstId, secondId);
	assert.equal(timerInfo.timeout, 200);

	mockView.fire(secondId);
	assert.equal(taskRunCount, 2);
});

test("cancel clears the pending timeout and prevents task execution", async () => {
	const { CoalescedTimeout } = await loadCoalescedTimeoutModule();
	const mockView = createMockTimeoutView();
	let taskRunCount = 0;

	const timeout = new CoalescedTimeout(mockView, () => {
		taskRunCount += 1;
	});

	timeout.schedule(100);
	assert.equal(mockView.timers.size, 1);

	timeout.cancel();
	assert.equal(mockView.timers.size, 0);

	assert.equal(taskRunCount, 0);
});

test("cancel is a safe no-op when no timeout is pending", async () => {
	const { CoalescedTimeout } = await loadCoalescedTimeoutModule();
	const mockView = createMockTimeoutView();
	let taskRunCount = 0;

	const timeout = new CoalescedTimeout(mockView, () => {
		taskRunCount += 1;
	});

	assert.doesNotThrow(() => {
		timeout.cancel();
	});

	assert.equal(mockView.timers.size, 0);
	assert.equal(taskRunCount, 0);
});

test("allows scheduling again after a timeout has been canceled", async () => {
	const { CoalescedTimeout } = await loadCoalescedTimeoutModule();
	const mockView = createMockTimeoutView();
	let taskRunCount = 0;

	const timeout = new CoalescedTimeout(mockView, () => {
		taskRunCount += 1;
	});

	timeout.schedule(100);
	timeout.cancel();
	assert.equal(mockView.timers.size, 0);

	timeout.schedule(50);
	assert.equal(mockView.timers.size, 1);

	const [timerId] = Array.from(mockView.timers.keys());
	mockView.fire(timerId);
	assert.equal(taskRunCount, 1);
});
