import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);

let profilerModulePromise;
function loadProfilerModule() {
  profilerModulePromise ??= bundleSource(
    "src/save/saveInterferenceProfiler.ts",
    "save-interference-profiler.cjs",
  ).then((outfile) => require(outfile));
  return profilerModulePromise;
}

let coordinatorModulePromise;
function loadCoordinatorModule() {
  coordinatorModulePromise ??= bundleSource(
    "src/save/DocumentSaveCoordinator.ts",
    "document-save-coordinator.cjs",
  ).then((outfile) => require(outfile));
  return coordinatorModulePromise;
}

test("autosave trace correlates save phases with overlapping typing and frame stalls", async () => {
  const { createAutosaveInterferenceProfiler } = await loadProfilerModule();
  const profiler = createAutosaveInterferenceProfiler();

  profiler.beginSave({
    scope: "docx",
    documentBytes: 1000,
    startedAt: 10,
  });
  profiler.recordTyping({
    startedAt: 12,
    completedAt: 72,
    inputToModelMs: 20,
    frameSchedulingDelayMs: 40,
    inputToVisibleMs: 60,
    slow: true,
  });
  profiler.recordFrame({ timestamp: 50, frameGapMs: 55, eventLoopDelayMs: 52 });

  const summary = profiler.completeSave({
    scope: "docx",
    startedAt: 10,
    endedAt: 80,
    outputBytes: 1040,
    changedContentBytes: 40,
    phases: [
      { name: "serialize", durationMs: 48, synchronousWorkMs: 40 },
      { name: "prepare", durationMs: 4, synchronousWorkMs: 0 },
      { name: "validate", durationMs: 8, synchronousWorkMs: 0 },
      { name: "persist", durationMs: 10, synchronousWorkMs: 0 },
    ],
  });

  assert.equal(summary.scope, "docx");
  assert.equal(summary.documentBytes, 1000);
  assert.equal(summary.changedContentBytes, 40);
  assert.equal(summary.outputBytes, 1040);
  assert.equal(summary.saveDurationMs, 70);
  assert.equal(summary.synchronousWorkMs, 40);
  assert.equal(summary.responsiveness.typingSamples, 1);
  assert.equal(summary.responsiveness.slowTypingSamples, 1);
  assert.equal(summary.responsiveness.frameSamples, 1);
  assert.equal(summary.blockingDetected, true);
});

test("nonblocking autosave path does not report a blocking save", async () => {
  const { createAutosaveInterferenceProfiler } = await loadProfilerModule();
  const profiler = createAutosaveInterferenceProfiler();
  profiler.beginSave({ scope: "pptx", documentBytes: 200, startedAt: 10 });
  profiler.recordFrame({ timestamp: 20, frameGapMs: 16, eventLoopDelayMs: 0 });
  const summary = profiler.completeSave({
    scope: "pptx",
    startedAt: 10,
    endedAt: 30,
    phases: [{ name: "serialize", durationMs: 1, synchronousWorkMs: 0 }],
  });

  assert.equal(summary.blockingDetected, false);
  assert.equal(summary.synchronousWorkMs, 0);
  assert.equal(summary.responsiveness.lateFrameCount, 0);
});

test("bounded traces compare representative mobile and desktop before/after measurements", async () => {
  const { createAutosaveInterferenceProfiler } = await loadProfilerModule();
  const cases = [
    { scope: "docx", appMode: "desktop", documentBytes: 4_000, label: "small" },
    { scope: "docx", appMode: "mobile", documentBytes: 2_000_000, label: "large" },
    { scope: "pptx", appMode: "desktop", documentBytes: 5_000_000, label: "image-heavy" },
  ];
  const profiler = createAutosaveInterferenceProfiler();
  for (const [index, sample] of cases.entries()) {
    const startedAt = index * 100;
    profiler.beginSave({ ...sample, startedAt });
    profiler.recordTyping({
      startedAt: startedAt + 1,
      completedAt: startedAt + 18,
      inputToModelMs: 2,
      frameSchedulingDelayMs: 3,
      inputToVisibleMs: 17,
      slow: false,
    });
    profiler.completeSave({
      scope: sample.scope,
      startedAt,
      endedAt: startedAt + 20,
      outputBytes: sample.documentBytes + 100,
      changedContentBytes: 100,
      phases: [{ name: "serialize", durationMs: 4, synchronousWorkMs: 0 }],
    });
  }
  const summaries = profiler.getRecentSummaries();
  assert.deepEqual(summaries.map((summary) => summary.appMode), ["desktop", "mobile", "desktop"]);
  assert.deepEqual(summaries.map((summary) => summary.scope), ["docx", "docx", "pptx"]);
  assert.deepEqual(summaries.map((summary) => summary.documentBytes), [4_000, 2_000_000, 5_000_000]);
  assert.ok(summaries.every((summary) => summary.responsiveness.typingSamples === 1));
});

async function runCoordinatorSave(serialize) {
  const { DocumentSaveCoordinator } = await loadCoordinatorModule();
  let clock = 0;
  let measurement;
  const coordinator = new DocumentSaveCoordinator({
    adapter: {
      serialize,
      prepareForWrite: async (buffer) => buffer,
      validate: async (buffer) => buffer,
      persist: async () => {},
    },
    getContext: () => ({ sourceBytes: 10 }),
    getSaveMetrics: (context, output) => ({
      documentBytes: context.sourceBytes,
      outputBytes: output?.byteLength ?? null,
      changedContentBytes: output ? Math.abs(output.byteLength - context.sourceBytes) : null,
    }),
    now: () => clock,
    onSaveCompleted: (next) => { measurement = next; },
    autosave: { enabled: () => false, delayMs: () => 0, source: "autosave" },
  });
  coordinator.markDirty();
  assert.equal(await coordinator.save("autosave"), true);
  return measurement;
}

test("coordinator distinguishes deliberately blocking serialization from an async/nonblocking path", async () => {
  const blocking = await runCoordinatorSave(async () => {
    // The deterministic clock models main-thread work without sleeping the test.
    // Returning a promise keeps the adapter contract identical to production.
    return new ArrayBuffer(20);
  });
  assert.equal(blocking.phases[0].synchronousWorkMs, 0);

  const { DocumentSaveCoordinator } = await loadCoordinatorModule();
  let clock = 0;
  let measurement;
  const coordinator = new DocumentSaveCoordinator({
    adapter: {
      serialize: () => {
        clock = 40;
        return Promise.resolve(new ArrayBuffer(20));
      },
      prepareForWrite: async (buffer) => buffer,
      validate: async (buffer) => buffer,
      persist: async () => {},
    },
    getContext: () => ({ sourceBytes: 10 }),
    now: () => clock,
    onSaveCompleted: (next) => { measurement = next; },
    autosave: { enabled: () => false, delayMs: () => 0, source: "autosave" },
  });
  coordinator.markDirty();
  assert.equal(await coordinator.save("autosave"), true);

  assert.equal(measurement.phases[0].synchronousWorkMs, 40);
  assert.equal(measurement.phases[0].durationMs, 40);
});
