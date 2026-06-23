import assert from "node:assert/strict";
import { test } from "node:test";
import { loadLoggerModule } from "./helpers/load-plugin-modules.mjs";

test("logger retains diagnostic history and reports dropped entries", async () => {
  globalThis.window = {
    docxidianDebugLogging: false,
    docxidianDebugLogs: [],
  };

  const {
    debugLog,
    getDocxidianLogSnapshot,
    getDocxidianLogStats,
  } = await loadLoggerModule();

  debugLog("diagnostics", "nested error", {
    error: new Error("serialized failure"),
  });
  for (let index = 0; index < 2005; index += 1) {
    debugLog("feature", "retention probe", { index });
  }

  const logs = getDocxidianLogSnapshot();
  const stats = getDocxidianLogStats();
  assert.equal(logs.length, 2000);
  assert.equal(stats.maxRetainedEntries, 2000);
  assert.equal(stats.retainedEntries, 2000);
  assert.equal(stats.totalEntries, 2006);
  assert.equal(stats.droppedEntries, 6);
  assert.deepEqual(logs.at(-1)?.data, { index: 2004 });
});

test("logger serializes nested Error values", async () => {
  globalThis.window = {
    docxidianDebugLogging: false,
    docxidianDebugLogs: [],
  };

  const { debugLog, getDocxidianLogSnapshot } = await loadLoggerModule();
  debugLog("diagnostics", "nested error probe", {
    error: new Error("serialized failure"),
  });

  const data = getDocxidianLogSnapshot().at(-1)?.data;
  assert.equal(data.error.name, "Error");
  assert.equal(data.error.message, "serialized failure");
});
