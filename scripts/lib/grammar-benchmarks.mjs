import { performance } from "node:perf_hooks";
import { createRequire } from "node:module";
import { bundleSource } from "../../tests/helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);
let serviceModulePromise;

async function loadGrammarService() {
  serviceModulePromise ??= bundleSource(
    "src/harper/harperGrammarService.ts",
    "grammar-benchmark-service.cjs",
  ).then((outfile) => require(outfile));
  return serviceModulePromise;
}

function nextImmediate() {
  return new Promise((resolve) => setImmediate(resolve));
}

function createFixtureText(kind, units) {
  const sentence = kind === "docx"
    ? "This document section contains a repeated sentence for grammar checking and search."
    : "This presentation slide contains a repeated sentence for grammar checking and review.";
  return Array.from({ length: units }, (_, index) => `${sentence} Unit ${index}.`).join(" ");
}

function startFrameProbe() {
  let stopped = false;
  let last = performance.now();
  const gaps = [];
  const tick = () => {
    const now = performance.now();
    gaps.push(now - last);
    last = now;
    if (!stopped) setImmediate(tick);
  };
  setImmediate(tick);
  return {
    async stop() {
      stopped = true;
      await nextImmediate();
      return gaps;
    },
  };
}

function percentile(values, fraction) {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.min(sorted.length - 1, Math.floor((sorted.length - 1) * fraction))] ?? 0;
}

function summarize(values) {
  return {
    count: values.length,
    p50Ms: Math.round(percentile(values, 0.5) * 10) / 10,
    p95Ms: Math.round(percentile(values, 0.95) * 10) / 10,
    maxMs: Math.round(Math.max(0, ...values) * 10) / 10,
  };
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export async function runGrammarBenchmarkScenario({
  kind = "docx",
  units = 100,
  enabled = true,
  editCount = 12,
  lintDelayMs = 2,
  debounceMs = 4,
} = {}) {
  const { createHarperGrammarService } = await loadGrammarService();
  const fixtureText = createFixtureText(kind, units);
  const worker = {
    lintCalls: 0,
    active: 0,
    maxOutstanding: 0,
    async setup() {},
    async getDefaultLintConfig() { return {}; },
    async lint() {
      this.lintCalls += 1;
      this.active += 1;
      this.maxOutstanding = Math.max(this.maxOutstanding, this.active);
      await delay(lintDelayMs);
      this.active -= 1;
      return [];
    },
    async dispose() {},
  };
  const logEntries = [];
  const timers = new Set();
  const scheduleTimes = [];
  const service = createHarperGrammarService({
    debounceMs,
    now: () => performance.now(),
    schedule(callback, delayMs) {
      const scheduledAt = performance.now();
      const timer = setTimeout(() => {
        timers.delete(timer);
        scheduleTimes.push({ scheduledAt, firedAt: performance.now(), delayMs });
        callback();
      }, delayMs);
      timers.add(timer);
      return { cancel: () => { clearTimeout(timer); timers.delete(timer); } };
    },
    createLinter: () => worker,
    log: (entry) => logEntries.push(entry),
  });
  if (!enabled) service.disable();

  const requestDurations = [];
  const frameProbe = startFrameProbe();
  const promises = [];
  for (let index = 0; index < editCount; index += 1) {
    const started = performance.now();
    promises.push(service.requestLint(`${fixtureText} Edit ${index}`));
    requestDurations.push(performance.now() - started);
  }
  const results = await Promise.all(promises);
  const frameGaps = await frameProbe.stop();
  await service.dispose();

  const lintLogs = logEntries.filter((entry) => entry.data?.phase === "lint");
  const slowLintLogs = lintLogs.filter((entry) => entry.data?.durationMs >= 16.7);
  const debounceWaits = scheduleTimes.map((entry) => entry.firedAt - entry.scheduledAt);
  return {
    kind,
    enabled,
    units,
    textLength: fixtureText.length,
    editCount,
    lintRequests: worker.lintCalls,
    staleOrCanceledResults: results.filter((result) => result === null).length,
    completedResults: results.filter((result) => result !== null).length,
    lintDuration: summarize(lintLogs.map((entry) => Number(entry.data?.durationMs) || 0)),
    debounceWait: summarize(debounceWaits),
    typingLatency: summarize(requestDurations),
    diagnosticsProcessing: summarize(lintLogs.map((entry) => Number(entry.data?.durationMs) || 0)),
    outstandingLintRequests: worker.maxOutstanding,
    frameImpact: {
      maxGapMs: Math.round(Math.max(0, ...frameGaps) * 10) / 10,
      gapsOver16Ms: frameGaps.filter((gap) => gap > 16.7).length,
    },
    slowDiagnosticsRetained: slowLintLogs.length,
    logEntryCount: logEntries.length,
  };
}

export async function runGrammarBenchmarks(options = {}) {
  const scenarios = [];
  for (const kind of ["docx", "pptx"]) {
    for (const units of options.units ?? [10, 100, 500]) {
      for (const enabled of [false, true]) {
        options.onProgress?.(`Running ${kind} grammar ${enabled ? "enabled" : "disabled"} tier ${units}...`);
        scenarios.push(await runGrammarBenchmarkScenario({
          kind,
          units,
          enabled,
          editCount: options.editCount ?? 12,
          lintDelayMs: options.lintDelayMs ?? 2,
          debounceMs: options.debounceMs ?? 4,
        }));
      }
    }
  }
  return {
    schemaVersion: 1,
    timestamp: new Date().toISOString(),
    environment: { platform: process.platform, nodeVersion: process.version },
    scenarios,
  };
}
