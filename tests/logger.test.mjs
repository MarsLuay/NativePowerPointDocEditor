import assert from "node:assert/strict";
import { test } from "node:test";
import { loadLoggerModule } from "./helpers/load-plugin-modules.mjs";

// The retention probes intentionally emit thousands of debug events. Console
// output is not under test and would otherwise drown out the test summary.
console.debug = () => {};

test("logger retains diagnostic history and reports dropped entries", async () => {
  globalThis.window = {
    nativePowerPointDocEditorDebugLogging: false,
    nativePowerPointDocEditorDebugLogs: [],
  };

  const {
	configureNativePowerPointDocEditorLogger,
    debugLog,
    getNativePowerPointDocEditorLogSnapshot,
    getNativePowerPointDocEditorLogStats,
  } = await loadLoggerModule();

  configureNativePowerPointDocEditorLogger(true);

  debugLog("diagnostics", "nested error", {
    error: new Error("serialized failure"),
  });
  for (let index = 0; index < 2005; index += 1) {
    debugLog("feature", "retention probe", { index });
  }

  const logs = getNativePowerPointDocEditorLogSnapshot();
  const stats = getNativePowerPointDocEditorLogStats();
  assert.equal(logs.length, 2000);
  assert.equal(stats.maxRetainedEntries, 2000);
  assert.equal(stats.retainedEntries, 2000);
  assert.equal(stats.totalEntries, 2006);
  assert.equal(stats.droppedEntries, 6);
  assert.deepEqual(logs.at(-1)?.data, { index: 2004 });
});

test("logger serializes nested Error values", async () => {
  globalThis.window = {
    nativePowerPointDocEditorDebugLogging: false,
    nativePowerPointDocEditorDebugLogs: [],
  };

  const { configureNativePowerPointDocEditorLogger, debugLog, getNativePowerPointDocEditorLogSnapshot } = await loadLoggerModule();
  configureNativePowerPointDocEditorLogger(true);
  debugLog("diagnostics", "nested error probe", {
    error: new Error("serialized failure"),
  });

  const data = getNativePowerPointDocEditorLogSnapshot().at(-1)?.data;
  assert.equal(data.error.name, "Error");
  assert.equal(data.error.message, "serialized failure");
});

test("disabled debug logging does not evict diagnostic entries", async () => {
  const {
    configureNativePowerPointDocEditorLogger,
    debugLog,
    getNativePowerPointDocEditorLogSnapshot,
    getNativePowerPointDocEditorLogStats,
    warnLog,
  } = await loadLoggerModule();

  configureNativePowerPointDocEditorLogger(false);
  const before = getNativePowerPointDocEditorLogStats();
  for (let index = 0; index < 50; index += 1) {
    debugLog("hot-path", "should not retain", { index });
  }
  const afterDebug = getNativePowerPointDocEditorLogStats();
  assert.equal(afterDebug.totalEntries, before.totalEntries);
  assert.equal(afterDebug.retainedEntries, before.retainedEntries);

  warnLog("diagnostics", "still retained", { reason: "debug-disabled" });
  const logs = getNativePowerPointDocEditorLogSnapshot();
  assert.equal(getNativePowerPointDocEditorLogStats().totalEntries, before.totalEntries + 1);
  assert.equal(logs.at(-1)?.message, "still retained");
});

test("debug floods retain earlier warnings and errors", async () => {
  const {
    configureNativePowerPointDocEditorLogger,
    debugLog,
    errorLog,
    getNativePowerPointDocEditorLogSnapshot,
    warnLog,
  } = await loadLoggerModule();

  configureNativePowerPointDocEditorLogger(true);
  warnLog("recovery", "warning that must survive");
  errorLog("recovery", "error that must survive");
  for (let index = 0; index < 3000; index += 1) {
    debugLog("render", "noisy render event", { index });
  }

  const messages = getNativePowerPointDocEditorLogSnapshot().map((entry) => entry.message);
  assert.ok(messages.includes("warning that must survive"));
  assert.ok(messages.includes("error that must survive"));
});
