import { getRuntimeFrameProfile, requestRuntimeFrame } from './runtimeFrameProfiler';

export type TextInputLatencySource = 'keydown' | 'beforeinput' | 'paste';

export interface TextInputLatencySummary {
	scope: 'docx' | 'pptx';
	sampleCount: number;
	slowInteractionCount: number;
	inputToModelMs: LatencyPercentiles;
	frameSchedulingDelayMs: LatencyPercentiles;
	inputToVisibleMs: LatencyPercentiles;
	lastCorrelationId: string | null;
}

export interface LatencyPercentiles {
	p50: number | null;
	p95: number | null;
	max: number | null;
}

export interface TextInputLatencyEventLike {
	isTrusted?: unknown;
	key?: unknown;
	inputType?: unknown;
	ctrlKey?: unknown;
	metaKey?: unknown;
}

export interface TextInputLatencyTracker {
	begin(event: TextInputLatencyEventLike, source: TextInputLatencySource): string | null;
	markModelUpdated(correlationId?: string | null): void;
	getSummary(): TextInputLatencySummary;
	dispose(): void;
}

interface TextInputLatencyOptions {
	scope: 'docx' | 'pptx';
	now?: () => number;
	requestFrame?: (callback: (timestamp: number) => void) => number | null;
	onSummary?: (summary: TextInputLatencySummary) => void;
	onInteractionComplete?: (data: {
		startedAt: number;
		completedAt: number;
		inputToModelMs: number;
		frameSchedulingDelayMs: number;
		inputToVisibleMs: number;
		slow: boolean;
	}) => void;
	onSlowInteraction?: (data: {
		correlationId: string;
		source: TextInputLatencySource;
		inputToModelMs: number;
		frameSchedulingDelayMs: number;
		inputToVisibleMs: number;
		slowThresholdMs: number;
	}) => void;
}

interface PendingInteraction {
	correlationId: string;
	source: TextInputLatencySource;
	startedAt: number;
	modelUpdatedAt?: number;
}

const MAX_SAMPLES = 128;
const MAX_PENDING = 16;
const SUMMARY_INTERVAL = 32;
const INPUT_REUSE_WINDOW_MS = 100;

function monotonicNow(): number {
	return typeof performance !== 'undefined' && typeof performance.now === 'function'
		? performance.now()
		: Date.now();
}

function round(value: number): number {
	return Math.round(value * 10) / 10;
}

function percentiles(values: readonly number[]): LatencyPercentiles {
	if (values.length === 0) return { p50: null, p95: null, max: null };
	const sorted = [...values].sort((left, right) => left - right);
	const percentile = (ratio: number) => sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * ratio) - 1))]!;
	return {
		p50: round(percentile(0.5)),
		p95: round(percentile(0.95)),
		max: round(sorted[sorted.length - 1]!),
	};
}

function isRelevant(event: TextInputLatencyEventLike, source: TextInputLatencySource): boolean {
	if (event.isTrusted !== true) return false;
	if (source === 'paste') return true;
	const inputType = typeof event.inputType === 'string' ? event.inputType : '';
	if (source === 'beforeinput') return inputType.startsWith('insert') || inputType.startsWith('delete') || inputType.startsWith('format');
	const key = typeof event.key === 'string' ? event.key : '';
	return key.length === 1 || key === 'Backspace' || key === 'Delete' || key === 'Enter'
		|| ((event.ctrlKey === true || event.metaKey === true) && key.toLowerCase() === 'v');
}

export function createTextInputLatencyTracker(options: TextInputLatencyOptions): TextInputLatencyTracker {
	const now = options.now ?? monotonicNow;
	const requestFrame = options.requestFrame ?? requestRuntimeFrame;
	let nextId = 0;
	let lastCorrelationId: string | null = null;
	let slowInteractionCount = 0;
	let disposed = false;
	const pending = new Map<string, PendingInteraction>();
	const modelTimes: number[] = [];
	const frameDelays: number[] = [];
	const visibleTimes: number[] = [];

	const trim = (values: number[]) => {
		if (values.length > MAX_SAMPLES) values.splice(0, values.length - MAX_SAMPLES);
	};
	const summary = (): TextInputLatencySummary => ({
		scope: options.scope,
		sampleCount: visibleTimes.length,
		slowInteractionCount,
		inputToModelMs: percentiles(modelTimes),
		frameSchedulingDelayMs: percentiles(frameDelays),
		inputToVisibleMs: percentiles(visibleTimes),
		lastCorrelationId,
	});
	const emitSummaryIfNeeded = () => {
		if (visibleTimes.length > 0 && visibleTimes.length % SUMMARY_INTERVAL === 0) options.onSummary?.(summary());
	};

	return {
		begin(event, source) {
			if (disposed || !isRelevant(event, source)) return null;
			const startedAt = now();
			const recent = lastCorrelationId ? pending.get(lastCorrelationId) : undefined;
			if (source === 'beforeinput' && recent && recent.modelUpdatedAt === undefined && startedAt - recent.startedAt <= INPUT_REUSE_WINDOW_MS) {
				return recent.correlationId;
			}
			const correlationId = `${options.scope}-typing-${++nextId}`;
			pending.set(correlationId, { correlationId, source, startedAt });
			lastCorrelationId = correlationId;
			if (pending.size > MAX_PENDING) pending.delete(pending.keys().next().value as string);
			return correlationId;
		},
		markModelUpdated(correlationId) {
			if (disposed) return;
			const resolvedId = correlationId ?? lastCorrelationId;
			const interaction = resolvedId ? pending.get(resolvedId) : undefined;
			if (!interaction || interaction.modelUpdatedAt !== undefined) return;
			const modelUpdatedAt = now();
			interaction.modelUpdatedAt = modelUpdatedAt;
			const inputToModelMs = Math.max(0, modelUpdatedAt - interaction.startedAt);
			requestFrame((frameTimestamp) => {
				if (disposed || pending.get(interaction.correlationId) !== interaction) return;
				const visibleAt = Number.isFinite(frameTimestamp) ? frameTimestamp : now();
				const frameSchedulingDelayMs = Math.max(0, visibleAt - modelUpdatedAt);
				const inputToVisibleMs = Math.max(0, visibleAt - interaction.startedAt);
				modelTimes.push(inputToModelMs);
				frameDelays.push(frameSchedulingDelayMs);
				visibleTimes.push(inputToVisibleMs);
				trim(modelTimes);
				trim(frameDelays);
				trim(visibleTimes);
				const slowThresholdMs = getRuntimeFrameProfile().thresholds.substantialStallMs;
				const slow = inputToVisibleMs >= slowThresholdMs;
				options.onInteractionComplete?.({
					startedAt: interaction.startedAt,
					completedAt: visibleAt,
					inputToModelMs: round(inputToModelMs),
					frameSchedulingDelayMs: round(frameSchedulingDelayMs),
					inputToVisibleMs: round(inputToVisibleMs),
					slow,
				});
				if (slow) {
					slowInteractionCount += 1;
					options.onSlowInteraction?.({
						correlationId: interaction.correlationId,
						source: interaction.source,
						inputToModelMs: round(inputToModelMs),
						frameSchedulingDelayMs: round(frameSchedulingDelayMs),
						inputToVisibleMs: round(inputToVisibleMs),
						slowThresholdMs,
					});
				}
				pending.delete(interaction.correlationId);
				emitSummaryIfNeeded();
			});
		},
		getSummary: summary,
		dispose() {
			if (disposed) return;
			disposed = true;
			pending.clear();
			if (visibleTimes.length > 0) options.onSummary?.(summary());
		},
	};
}
