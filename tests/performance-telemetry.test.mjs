import assert from "node:assert/strict";
import { test } from "node:test";
import { createRequire } from "node:module";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);
let modulePromise;
async function loadModule() {
  modulePromise ??= bundleSource(
    "src/performanceTelemetry.ts",
    "performance-telemetry.cjs",
  ).then((outfile) => require(outfile));
  return modulePromise;
}

test("slow performance telemetry emits only normalized above-threshold events", async () => {
  const originalWindow = globalThis.window;
  const originalDebug = console.debug;
  const debugCalls = [];
  globalThis.window = {
    nativePowerPointDocEditorDebugLogging: true,
    nativePowerPointDocEditorDebugLogs: [],
  };
  console.debug = (...args) => debugCalls.push(args);

  try {
    const { logSlowPerformance, SLOW_PERFORMANCE_THRESHOLD_MS } = await loadModule();
    assert.equal(SLOW_PERFORMANCE_THRESHOLD_MS, 100);
    for (const duration of [Number.NaN, Number.POSITIVE_INFINITY, -1, 99.9, 100]) {
      assert.equal(logSlowPerformance("save", "slow operation", duration), false);
    }

    assert.equal(
      logSlowPerformance("save", "slow operation", 150.26, { scope: "pptx" }),
      true,
    );
    assert.equal(debugCalls.length, 1);
    assert.equal(debugCalls[0][0], "[Native PowerPoint Doc Editor] save: slow operation");
    assert.deepEqual(debugCalls[0][1], {
      scope: "pptx",
      durationMs: 150.3,
      thresholdMs: 100,
    });
  } finally {
    console.debug = originalDebug;
    globalThis.window = originalWindow;
  }
});
