import { performance } from "node:perf_hooks";

export const ENDURANCE_OPERATIONS = [
  "type", "delete", "paste", "format", "scroll", "zoom", "slide-switch",
  "drag", "resize", "undo", "redo", "search", "autosave", "manual-save", "open-close",
];
export const DEFAULT_ENDURANCE_DURATION_MS = 10 * 60 * 1000;
export const DEFAULT_SUMMARY_INTERVAL_MS = 10 * 1000;

function round(value) { return Math.round(value * 10) / 10; }
function createRng(seed) {
  let state = seed >>> 0;
  return () => {
    state = (state * 1664525 + 1013904223) >>> 0;
    return state / 0x100000000;
  };
}

export function createDeterministicWorkload({ seed = 224, actionCount = 1000 } = {}) {
  const random = createRng(seed);
  return Array.from({ length: actionCount }, (_, index) => ({
    index,
    kind: ENDURANCE_OPERATIONS[Math.floor(random() * ENDURANCE_OPERATIONS.length)] ?? "type",
    document: random() < 0.5 ? "docx" : "pptx",
    complexity: 1 + Math.floor(random() * 500),
  }));
}

function percentile(values, fraction) {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.min(sorted.length - 1, Math.floor((sorted.length - 1) * fraction))] ?? 0;
}

function summarize(values) {
  return {
    count: values.length,
    p95Ms: round(percentile(values, 0.95)),
    maxMs: round(Math.max(0, ...values)),
  };
}

function defaultActionMeasurement(action) {
  const base = action.kind === "type" ? 1.5 : action.kind === "drag" || action.kind === "resize" ? 3 : 2;
  const complexityPenalty = action.complexity / 500;
  return {
    typingMs: action.kind === "type" ? base + complexityPenalty : 0,
    frameGapMs: base + complexityPenalty,
    saveMs: action.kind === "autosave" || action.kind === "manual-save" ? 4 + complexityPenalty * 4 : 0,
    resources: {
      domNodes: action.kind === "open-close" ? 0 : 1,
      mutationObservers: action.kind === "open-close" ? -1 : 0,
      timers: action.kind === "open-close" ? -1 : 0,
      cacheEntries: action.kind === "slide-switch" ? 1 : 0,
      editorViews: action.kind === "open-close" ? -1 : 0,
    },
  };
}

function aggregate(actions, measurements, resources) {
  return {
    actions: actions.length,
    actionKinds: Object.fromEntries(ENDURANCE_OPERATIONS.map((kind) => [kind, actions.filter((action) => action.kind === kind).length])),
    typing: summarize(measurements.filter((measurement) => measurement.typingMs > 0).map((measurement) => measurement.typingMs)),
    frame: summarize(measurements.map((measurement) => measurement.frameGapMs)),
    save: summarize(measurements.filter((measurement) => measurement.saveMs > 0).map((measurement) => measurement.saveMs)),
    resources: { ...resources },
  };
}

export async function runEnduranceHarness(options = {}) {
  const actionCount = options.actionCount ?? 1000;
  const workload = options.workload ?? createDeterministicWorkload({ seed: options.seed ?? 224, actionCount });
  const summaryInterval = Math.max(1, options.summaryEveryActions ?? 100);
  const summaries = [];
  const measurements = [];
  const resources = { domNodes: 0, mutationObservers: 0, timers: 0, cacheEntries: 0, editorViews: 0 };
  const startedAt = performance.now();

  for (let index = 0; index < workload.length; index += 1) {
    const action = workload[index];
    const measurement = await (options.measureAction?.(action) ?? defaultActionMeasurement(action));
    measurements.push(measurement);
    for (const [key, delta] of Object.entries(measurement.resources ?? {})) {
      resources[key] = Math.max(0, (resources[key] ?? 0) + delta);
    }
    if ((index + 1) % summaryInterval === 0 || index === workload.length - 1) {
      summaries.push({
        actionIndex: index,
        elapsedMs: round(performance.now() - startedAt),
        summary: aggregate(workload.slice(0, index + 1), measurements, resources),
        heapUsedBytes: process.memoryUsage?.().heapUsed ?? null,
      });
    }
  }

  const first = summaries[0]?.summary ?? aggregate([], [], resources);
  const final = summaries.at(-1)?.summary ?? first;
  return {
    schemaVersion: 1,
    seed: options.seed ?? 224,
    actionCount: workload.length,
    summaries,
    comparison: {
      early: first,
      late: final,
      typingP95DeltaMs: round(final.typing.p95Ms - first.typing.p95Ms),
      frameMaxDeltaMs: round(final.frame.maxMs - first.frame.maxMs),
      saveP95DeltaMs: round(final.save.p95Ms - first.save.p95Ms),
      cacheEntryGrowth: final.resources.cacheEntries - first.resources.cacheEntries,
      domNodeGrowth: final.resources.domNodes - first.resources.domNodes,
      timerGrowth: final.resources.timers - first.resources.timers,
      observerGrowth: final.resources.mutationObservers - first.resources.mutationObservers,
      editorViewGrowth: final.resources.editorViews - first.resources.editorViews,
    },
  };
}
