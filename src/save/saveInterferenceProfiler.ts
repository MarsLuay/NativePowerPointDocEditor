import {
	getRuntimeFrameProfile,
	subscribeRuntimeFrameObserver,
	type RuntimeFrameAppMode,
} from '../runtimeFrameProfiler';

export type SaveTraceScope = 'docx' | 'pptx';

export interface SaveTraceTypingSample {
	startedAt: number;
	completedAt: number;
	inputToModelMs: number;
	frameSchedulingDelayMs: number;
	inputToVisibleMs: number;
	slow: boolean;
}

export interface SaveTraceFrameSample {
	timestamp: number;
	frameGapMs: number;
	eventLoopDelayMs: number;
}

export interface SaveTracePhase {
	name: 'serialize' | 'prepare' | 'validate' | 'persist';
	durationMs: number;
	synchronousWorkMs: number;
}

export interface SaveTraceLatencySummary {
	count: number;
	p50: number | null;
	p95: number | null;
	max: number | null;
}

export interface SaveTraceResponsiveness {
	typingSamples: number;
	slowTypingSamples: number;
	inputToVisibleMs: SaveTraceLatencySummary;
	frameSchedulingDelayMs: SaveTraceLatencySummary;
	frameSamples: number;
	lateFrameCount: number;
	longestFrameGapMs: number | null;
	maxEventLoopDelayMs: number | null;
}

export interface AutosaveInterferenceSummary {
	id: number;
	scope: SaveTraceScope;
	appMode: RuntimeFrameAppMode;
	documentBytes: number | null;
	changedContentBytes: number | null;
	outputBytes: number | null;
	startedAt: number;
	endedAt: number;
	saveDurationMs: number;
	synchronousWorkMs: number;
	blockingDetected: boolean;	phases: SaveTracePhase[];	
	responsiveness: SaveTraceResponsiveness;
}

export interface SaveTraceStart {
	scope: SaveTraceScope;
	appMode?: RuntimeFrameAppMode;
	documentBytes?: number | null;
	changedContentBytes?: number | null;
	startedAt?: number;
}

export interface SaveTraceCompletion {
	id?: number;
	scope: SaveTraceScope;
	documentBytes?: number | null;
	changedContentBytes?: number | null;
	outputBytes?: number | null;
	startedAt: number;
	endedAt: number;
	saveDurationMs?: number;
	synchronousWorkMs?: number;
	phases?: readonly SaveTracePhase[];
}

export interface AutosaveInterferenceProfiler {
	beginSave(start: SaveTraceStart): number;
	recordTyping(sample: SaveTraceTypingSample): void;
	recordFrame(sample: SaveTraceFrameSample): void;
	completeSave(completion: SaveTraceCompletion): AutosaveInterferenceSummary | null;
	getRecentSummaries(): AutosaveInterferenceSummary[];
	reset(): void;
	dispose(): void;
}

const MAX_RECENT_SUMMARIES = 24;
const FRAME_BUDGET_FALLBACK_MS = 16.7;
const LATE_FRAME_MULTIPLIER = 1.5;
const FRAME_GRACE_MS = 100;

function finite(value: number | null | undefined): number | null {
	return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function nonNegative(value: number | null | undefined): number {
	return Math.max(0, finite(value) ?? 0);
}

function round(value: number): number {
	return Math.round(value * 10) / 10;
}

function percentile(values: readonly number[], ratio: number): number | null {
	if (values.length === 0) return null;
	const sorted = [...values].sort((left, right) => left - right);
	return round(sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * ratio) - 1))]!);
}

function summarize(values: readonly number[]): SaveTraceLatencySummary {
	return {
		count: values.length,
		p50: percentile(values, 0.5),
		p95: percentile(values, 0.95),
		max: percentile(values, 1),
	};
}

function appMode(): RuntimeFrameAppMode {
	return getRuntimeFrameProfile().appMode;
}

interface MutableTrace {
	id: number;
	scope: SaveTraceScope;
	appMode: RuntimeFrameAppMode;
	documentBytes: number | null;
	changedContentBytes: number | null;
	startedAt: number;
	endedAt: number | null;
	outputBytes: number | null;
	startedTyping: SaveTraceTypingSample[];
	frames: SaveTraceFrameSample[];
}

function overlapsTrace(trace: MutableTrace, startedAt: number, completedAt = startedAt): boolean {
	const end = trace.endedAt === null ? Number.POSITIVE_INFINITY : trace.endedAt + FRAME_GRACE_MS;
	return completedAt >= trace.startedAt && startedAt <= end;
}

function traceResponsiveness(trace: MutableTrace): SaveTraceResponsiveness {
	const frameBudgetMs = getRuntimeFrameProfile().resolvedFrameBudgetMs || FRAME_BUDGET_FALLBACK_MS;
	const typingVisible = trace.startedTyping.map((sample) => nonNegative(sample.inputToVisibleMs));
	const typingFrameDelay = trace.startedTyping.map((sample) => nonNegative(sample.frameSchedulingDelayMs));
	const frameGaps = trace.frames.map((sample) => nonNegative(sample.frameGapMs));
	const eventLoopDelays = trace.frames.map((sample) => nonNegative(sample.eventLoopDelayMs));
	return {
		typingSamples: trace.startedTyping.length,
		slowTypingSamples: trace.startedTyping.filter((sample) => sample.slow).length,
		inputToVisibleMs: summarize(typingVisible),
		frameSchedulingDelayMs: summarize(typingFrameDelay),
		frameSamples: trace.frames.length,
		lateFrameCount: frameGaps.filter((gap) => gap >= frameBudgetMs * LATE_FRAME_MULTIPLIER).length,
		longestFrameGapMs: frameGaps.length > 0 ? round(Math.max(...frameGaps)) : null,
		maxEventLoopDelayMs: eventLoopDelays.length > 0 ? round(Math.max(...eventLoopDelays)) : null,
	};
}

function toSummary(trace: MutableTrace, completion: SaveTraceCompletion): AutosaveInterferenceSummary {
	const phases = (completion.phases ?? []).map((phase) => ({
		name: phase.name,
		durationMs: round(nonNegative(phase.durationMs)),
		synchronousWorkMs: round(nonNegative(phase.synchronousWorkMs)),
	}));
	const synchronousWorkMs = round(completion.synchronousWorkMs ?? phases.reduce((total, phase) => total + phase.synchronousWorkMs, 0));
	const endedAt = finite(completion.endedAt) ?? trace.startedAt;
	const saveDurationMs = round(completion.saveDurationMs ?? Math.max(0, endedAt - trace.startedAt));
	const responsiveness = traceResponsiveness(trace);
	const frameBudgetMs = getRuntimeFrameProfile().resolvedFrameBudgetMs || FRAME_BUDGET_FALLBACK_MS;
	const blockingDetected = synchronousWorkMs >= Math.max(8, frameBudgetMs)
		|| (responsiveness.maxEventLoopDelayMs ?? 0) >= frameBudgetMs * LATE_FRAME_MULTIPLIER
		|| responsiveness.slowTypingSamples > 0;
	return {
		id: trace.id,
		scope: trace.scope,
		appMode: trace.appMode,
		documentBytes: finite(completion.documentBytes) ?? trace.documentBytes,
		changedContentBytes: finite(completion.changedContentBytes) ?? trace.changedContentBytes,
		outputBytes: finite(completion.outputBytes) ?? trace.outputBytes,
		startedAt: round(trace.startedAt),
		endedAt: round(endedAt),
		saveDurationMs,
		synchronousWorkMs,
		blockingDetected,
		phases,
		responsiveness,
	};
}

/**
 * Aggregates save windows with bounded typing/frame evidence. It deliberately
 * retains no text, paths, or individual keystroke records.
 */
export function createAutosaveInterferenceProfiler(): AutosaveInterferenceProfiler {
	let nextId = 0;
	let active: MutableTrace | null = null;
	let recent: AutosaveInterferenceSummary[] = [];
	let previousFrameTimestamp: number | null = null;

	const recordFrame = (sample: SaveTraceFrameSample): void => {
		if (!active && recent.length === 0) return;
		const timestamp = finite(sample.timestamp);
		if (timestamp === null) return;
		const resolved = {
			timestamp,
			frameGapMs: nonNegative(sample.frameGapMs),
			eventLoopDelayMs: nonNegative(sample.eventLoopDelayMs),
		};
		if (active && overlapsTrace(active, timestamp)) {
			active.frames.push(resolved);
			return;
		}
		const last = recent.at(-1);
		if (last && timestamp >= last.startedAt && timestamp <= last.endedAt + FRAME_GRACE_MS) {
			// A blocking save can release its queued rAF after the save promise
			// resolves. Keep that one bounded frame attached to the just-closed save.
			const frameGaps = last.responsiveness.longestFrameGapMs === null
				? [resolved.frameGapMs]
				: [last.responsiveness.longestFrameGapMs, resolved.frameGapMs];
			last.responsiveness.frameSamples += 1;
			last.responsiveness.longestFrameGapMs = round(Math.max(...frameGaps));
			last.responsiveness.maxEventLoopDelayMs = round(Math.max(
				last.responsiveness.maxEventLoopDelayMs ?? 0,
				resolved.eventLoopDelayMs,
			));
			const frameBudgetMs = getRuntimeFrameProfile().resolvedFrameBudgetMs || FRAME_BUDGET_FALLBACK_MS;
			if (resolved.frameGapMs >= frameBudgetMs * LATE_FRAME_MULTIPLIER) last.responsiveness.lateFrameCount += 1;
			if (resolved.eventLoopDelayMs >= frameBudgetMs * LATE_FRAME_MULTIPLIER) last.blockingDetected = true;
		}
	};

	const unsubscribeFrameObserver = subscribeRuntimeFrameObserver((timestamp) => {
		const now = typeof performance !== 'undefined' && typeof performance.now === 'function' ? performance.now() : Date.now();
		const gap = previousFrameTimestamp === null ? 0 : Math.max(0, timestamp - previousFrameTimestamp);
		previousFrameTimestamp = timestamp;
		recordFrame({
			timestamp,
			frameGapMs: gap,
			eventLoopDelayMs: Math.max(0, now - timestamp),
		});
	});

	return {
		beginSave(start) {
			if (active) {
				// A coordinator never overlaps saves, but preserve a bounded record if
				// an adapter violates that contract in a host integration.
				active.endedAt = start.startedAt ?? active.startedAt;
				recent.push(toSummary(active, {
					scope: active.scope,
					startedAt: active.startedAt,
					endedAt: active.endedAt,
				}));
			}
			active = {
				id: ++nextId,
				scope: start.scope,
				appMode: start.appMode ?? appMode(),
				documentBytes: finite(start.documentBytes),
				changedContentBytes: finite(start.changedContentBytes),
				startedAt: finite(start.startedAt) ?? (typeof performance !== 'undefined' && typeof performance.now === 'function' ? performance.now() : Date.now()),
				endedAt: null,
				outputBytes: null,
				startedTyping: [],
				frames: [],
			};
			return active.id;
		},
		recordTyping(sample) {
			const startedAt = finite(sample.startedAt);
			const completedAt = finite(sample.completedAt);
			if (startedAt === null || completedAt === null) return;
			const target = active && overlapsTrace(active, startedAt, completedAt)
				? active
				: null;
			if (target) {
				target.startedTyping.push({ ...sample, startedAt, completedAt });
				return;
			}
			// Late typing frames cannot mutate a complete latency distribution without
			// retaining every sample. Update only the bounded count and maxima/p95.
			const last = recent.at(-1);
			if (!last || completedAt < last.startedAt || startedAt > last.endedAt + FRAME_GRACE_MS) return;
			last.responsiveness.typingSamples += 1;
			if (sample.slow) {
				last.responsiveness.slowTypingSamples += 1;
				last.blockingDetected = true;
			}
			const value = round(nonNegative(sample.inputToVisibleMs));
			const current = last.responsiveness.inputToVisibleMs.max ?? 0;
			last.responsiveness.inputToVisibleMs.max = Math.max(current, value);
			last.responsiveness.inputToVisibleMs.count += 1;
		},
		recordFrame,
		completeSave(completion) {
			if (!active) return null;
			active.endedAt = completion.endedAt;
			active.outputBytes = finite(completion.outputBytes);
			const summary = toSummary(active, { ...completion, id: active.id });
			active = null;
			recent = [...recent, summary].slice(-MAX_RECENT_SUMMARIES);
			return summary;
		},
		getRecentSummaries() {
			return recent.map((summary) => ({
				...summary,
				phases: summary.phases.map((phase) => ({ ...phase })),
				responsiveness: {
					...summary.responsiveness,
					inputToVisibleMs: { ...summary.responsiveness.inputToVisibleMs },
					frameSchedulingDelayMs: { ...summary.responsiveness.frameSchedulingDelayMs },
				},
			}));
		},
		reset() {
			active = null;
			recent = [];
			previousFrameTimestamp = null;
		},
		dispose() {
			unsubscribeFrameObserver();
			active = null;
			recent = [];
			previousFrameTimestamp = null;
		},
	};

}

let sharedProfiler: AutosaveInterferenceProfiler | null = null;

export function getSharedAutosaveInterferenceProfiler(): AutosaveInterferenceProfiler {
	sharedProfiler ??= createAutosaveInterferenceProfiler();
	return sharedProfiler;
}

export function resetSharedAutosaveInterferenceProfiler(): void {
	sharedProfiler?.dispose();
	sharedProfiler = null;
}
