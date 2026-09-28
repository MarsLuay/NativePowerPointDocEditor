import assert from "node:assert/strict";
import { test } from "node:test";
import { createRequire } from "node:module";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);
let modulePromise;

async function loadDiagnostics() {
  modulePromise ??= bundleSource(
    "src/sessionMemoryDiagnostics.ts",
    "session-memory-diagnostics.cjs"
  ).then((outfile) => require(outfile));
  return modulePromise;
}

test("detectMemoryCapability reports explicit support or null when unsupported, never zero", async () => {
  const { detectMemoryCapability } = await loadDiagnostics();

  // 1. Node.js environment with process.memoryUsage
  const nodeReport = detectMemoryCapability({
    process: {
      memoryUsage: () => ({
        heapUsed: 1234567,
        heapTotal: 2345678,
        external: 50000,
        rss: 10000000,
        arrayBuffers: 20000,
      }),
    },
  });
  assert.equal(nodeReport.supported, true);
  assert.equal(nodeReport.provider, "v8-process");
  assert.equal(nodeReport.usedBytes, 1234567);
  assert.equal(nodeReport.totalBytes, 2345678);

  // 2. Chromium / Performance.memory environment
  const chromeReport = detectMemoryCapability({
    performance: {
      memory: {
        usedJSHeapSize: 9876543,
        totalJSHeapSize: 15000000,
        jsHeapSizeLimit: 30000000,
      },
    },
  });
  assert.equal(chromeReport.supported, true);
  assert.equal(chromeReport.provider, "performance-memory");
  assert.equal(chromeReport.usedBytes, 9876543);
  assert.equal(chromeReport.totalBytes, 15000000);
  assert.equal(chromeReport.limitBytes, 30000000);

  // 3. WKWebView / iOS / Safari where no memory API is exposed
  const unsupportedReport = detectMemoryCapability({
    process: {},
    performance: {},
  });
  assert.equal(unsupportedReport.supported, false);
  assert.equal(unsupportedReport.provider, "unsupported");
  assert.equal(unsupportedReport.usedBytes, null, "Must be null, NEVER 0");
  assert.equal(unsupportedReport.totalBytes, null);
  assert.equal(unsupportedReport.limitBytes, null);
  assert.ok(unsupportedReport.detail?.includes("WKWebView"));
});

test("sessionResourceRegistry tracks views using WeakRef and cleans up without retaining objects", async () => {
  const { sessionResourceRegistry } = await loadDiagnostics();
  sessionResourceRegistry.reset();

  let view1 = { id: "docx-view-1" };
  let view2 = { id: "docx-view-2" };
  const unmount1 = sessionResourceRegistry.registerView("docx", view1);
  const unmount2 = sessionResourceRegistry.registerView("docx", view2);

  let pptxView = { id: "pptx-view-1" };
  const unmountPptx = sessionResourceRegistry.registerView("pptx", pptxView);

  let counters = sessionResourceRegistry.getCounters();
  assert.equal(counters.mountedDocxViews, 2);
  assert.equal(counters.mountedPptxViews, 1);

  // Unregister view 1
  unmount1();
  counters = sessionResourceRegistry.getCounters();
  assert.equal(counters.mountedDocxViews, 1);

  // Unregister view 2 and pptxView
  unmount2();
  unmountPptx();
  counters = sessionResourceRegistry.getCounters();
  assert.equal(counters.mountedDocxViews, 0);
  assert.equal(counters.mountedPptxViews, 0);
});

test("resource counters track timers, animation frames, observers, and listeners with clean teardown", async () => {
  const { sessionResourceRegistry } = await loadDiagnostics();
  sessionResourceRegistry.reset();

  const unregisterTimer = sessionResourceRegistry.registerTimer();
  const unregisterFrame = sessionResourceRegistry.registerAnimationFrame();
  const unregisterMutation = sessionResourceRegistry.registerObserver("mutation");
  const unregisterResize = sessionResourceRegistry.registerObserver("resize");
  const unregisterListener = sessionResourceRegistry.registerListener();
  sessionResourceRegistry.setThumbnailCacheEntries(42);

  let counters = sessionResourceRegistry.getCounters();
  assert.equal(counters.activeTimers, 1);
  assert.equal(counters.activeAnimationFrames, 1);
  assert.equal(counters.activeMutationObservers, 1);
  assert.equal(counters.activeResizeObservers, 1);
  assert.equal(counters.registeredListeners, 1);
  assert.equal(counters.thumbnailCacheEntries, 42);

  // Clean teardown
  unregisterTimer();
  unregisterFrame();
  unregisterMutation();
  unregisterResize();
  unregisterListener();
  sessionResourceRegistry.setThumbnailCacheEntries(0);

  counters = sessionResourceRegistry.getCounters();
  assert.equal(counters.activeTimers, 0);
  assert.equal(counters.activeAnimationFrames, 0);
  assert.equal(counters.activeMutationObservers, 0);
  assert.equal(counters.activeResizeObservers, 0);
  assert.equal(counters.registeredListeners, 0);
  assert.equal(counters.thumbnailCacheEntries, 0);
});

test("compareSessionSnapshots flags retained resources and accurately reports clean teardown", async () => {
  const { captureSessionSnapshot, compareSessionSnapshots, sessionResourceRegistry } = await loadDiagnostics();
  sessionResourceRegistry.reset();

  const baseline = captureSessionSnapshot("baseline");

  // Simulate opening views and setting up observers
  const dummyView = { id: "test-view" };
  const unmount = sessionResourceRegistry.registerView("docx", dummyView);
  const unregTimer = sessionResourceRegistry.registerTimer();
  const unregObserver = sessionResourceRegistry.registerObserver("mutation");

  const midSession = captureSessionSnapshot("mid-session");
  const dirtyDelta = compareSessionSnapshots(baseline, midSession);
  assert.equal(dirtyDelta.isClean, false);
  assert.equal(dirtyDelta.retainedDocxViews, 1);
  assert.equal(dirtyDelta.retainedTimers, 1);
  assert.equal(dirtyDelta.retainedObservers, 1);
  assert.ok(dirtyDelta.reasons.length >= 3);

  // Teardown
  unmount();
  unregTimer();
  unregObserver();

  const final = captureSessionSnapshot("final");
  const cleanDelta = compareSessionSnapshots(baseline, final);
  assert.equal(cleanDelta.isClean, true);
  assert.equal(cleanDelta.retainedDocxViews, 0);
  assert.equal(cleanDelta.retainedTimers, 0);
  assert.equal(cleanDelta.retainedObservers, 0);
  assert.equal(cleanDelta.reasons.length, 0);
});

test("runSessionEnduranceTest runs repeated open/close and document switching cycles", async () => {
  const { runSessionEnduranceTest, sessionResourceRegistry } = await loadDiagnostics();
  sessionResourceRegistry.reset();

  let activeView = null;
  let activeCleanup = null;

  const report = await runSessionEnduranceTest({
    cycles: 5,
    onCycle: async (cycle) => {
      // Simulate document switch: alternate between docx and pptx
      const scope = cycle % 2 === 0 ? "docx" : "pptx";
      activeView = { name: `doc-${cycle}` };
      activeCleanup = sessionResourceRegistry.registerView(scope, activeView);

      // Simulate editor work & timer
      const unregTimer = sessionResourceRegistry.registerTimer();

      // Teardown before next cycle
      unregTimer();
      activeCleanup();
      activeCleanup = null;
      activeView = null;
    },
  });

  assert.equal(report.cycles, 5);
  assert.equal(report.passed, true);
  assert.equal(report.leakDelta.isClean, true);
  assert.equal(report.leakDelta.retainedDocxViews, 0);
  assert.equal(report.leakDelta.retainedPptxViews, 0);
});
