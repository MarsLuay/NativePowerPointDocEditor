import { getRuntimeFrameProfile, type RuntimeFrameProfile } from './runtimeFrameProfiler';
import { debugLog, warnLog } from './logger';

export type ContinuousInteractionType =
	| 'docx-scroll'
	| 'docx-zoom'
	| 'pptx-canvas-scroll'
	| 'pptx-filmstrip-scroll'
	| 'pptx-zoom'
	| 'pptx-shape-drag'
	| 'pptx-multi-selection-drag'
	| 'pptx-resize'
	| 'pptx-multi-selection-resize'
	| 'pptx-rotate'
	| 'pptx-multi-selection-rotate'
	| 'pptx-image-crop'
	| 'pptx-text-box-resize';

export interface ContinuousInteractionEventDetails {
	/** Pointer samples reported by PointerEvent.getCoalescedEvents(). */
	coalescedInputCount?: number;
	/** An optional platform-provided estimate; ordinary browser events cannot observe lost input. */
	droppedInputEstimate?: number;
}

export interface ContinuousInteractionMetric {
	p50: number | null;
	p95: number | null;
	max: number | null;
}

export type FrameAnomalyClassification = 'normal' | 'late' | 'missed' | 'stall';

export interface WorstFrameRecord {
	timestamp: number;
	intervalMs: number;
	rafDelayMs: number;
	synchronousWorkMs: number;
	eventLoopDelayMs: number;
	classification: FrameAnomalyClassification;
}

export interface ContinuousInteractionSummary {
	scope: 'docx' | 'pptx';
	interactionType: ContinuousInteractionType;
	frameCount: number;
	totalDurationMs: number;
	lateFrameCount: number;
	missedFrameEstimate: number;
	longestFrameGapMs: number;
	frameIntervalsMs: ContinuousInteractionMetric;
	rafSchedulingDelayMs: ContinuousInteractionMetric;
	synchronousWorkMs: ContinuousInteractionMetric;
	eventLoopDelayMs: ContinuousInteractionMetric;
	inputEventCount: number;
	coalescedInputCount: number;
	droppedInputEstimate: number;
	domMutationCount: number;
	worstFrames: WorstFrameRecord[];
	frameBudgetMs: number;
	resolvedRefreshHz: number;
}

export interface ContinuousInteractionProfilerOptions {
	now?: () => number;
	requestFrame?: (callback: (timestamp: number) => void) => number | null;
	cancelFrame?: (handle: number) => void;
	scheduleTimeout?: (callback: () => void, ms: number) => number | null;
	cancelTimeout?: (handle: number) => void;
	getFrameProfile?: () => RuntimeFrameProfile;
	settleTimeoutMs?: number;
	maxHistorySummaries?: number;
	maxWorstFrames?: number;
	onSummary?: (summary: ContinuousInteractionSummary) => void;
	onSlowInteraction?: (summary: ContinuousInteractionSummary) => void;
}

const DEFAULT_SETTLE_TIMEOUT_MS = 150;
const MAX_HISTORY_SUMMARIES = 16;
const MAX_WORST_FRAMES = 5;

function defaultNow(): number {
	return typeof performance !== 'undefined' && typeof performance.now === 'function'
		? performance.now()
		: Date.now();
}

function defaultRequestFrame(callback: (timestamp: number) => void): number | null {
	if (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
		return window.requestAnimationFrame(callback);
	}
	return null;
}

function defaultCancelFrame(handle: number): void {
	if (typeof window !== 'undefined' && typeof window.cancelAnimationFrame === 'function') {
		window.cancelAnimationFrame(handle);
	}
}

function defaultScheduleTimeout(callback: () => void, ms: number): number | null {
	if (typeof window !== 'undefined' && typeof window.setTimeout === 'function') {
		return window.setTimeout(callback, ms);
	}
	return null;
}

function defaultCancelTimeout(handle: number): void {
	if (typeof window !== 'undefined' && typeof window.clearTimeout === 'function') {
		window.clearTimeout(handle);
	}
}

function round(value: number): number {
	return Math.round(value * 10) / 10;
}

function percentiles(values: readonly number[]): ContinuousInteractionMetric {
	if (values.length === 0) return { p50: null, p95: null, max: null };
	const sorted = [...values].sort((left, right) => left - right);
	const p = (ratio: number) => sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * ratio) - 1))]!;
	return {
		p50: round(p(0.5)),
		p95: round(p(0.95)),
		max: round(sorted[sorted.length - 1]!),
	};
}

class InteractionSession {
	readonly type: ContinuousInteractionType;
	readonly scope: 'docx' | 'pptx';
	readonly startedAt: number;
	private lastFrameTimestamp: number | null = null;
	private lastRequestTime: number | null = null;
	private frameHandle: number | null = null;
	private settleTimer: number | null = null;
	private syncWorkAccMs = 0;
	inputEventCount = 0;
	coalescedInputCount = 0;
	droppedInputEstimate = 0;
	domMutationCount = 0;

	readonly frameIntervals: number[] = [];
	readonly rafDelays: number[] = [];
	readonly syncWorks: number[] = [];
	readonly eventLoopDelays: number[] = [];
	readonly worstFrames: WorstFrameRecord[] = [];
	lateFrameCount = 0;
	missedFrameEstimate = 0;
	longestFrameGapMs = 0;

	constructor(type: ContinuousInteractionType, now: number) {
		this.type = type;
		this.scope = type.startsWith('docx') ? 'docx' : 'pptx';
		this.startedAt = now;
	}

	addSynchronousWork(durationMs: number): void {
		this.syncWorkAccMs += Math.max(0, durationMs);
	}

	recordInput(details: ContinuousInteractionEventDetails = {}): void {
		this.inputEventCount += 1;
		this.coalescedInputCount += Math.max(0, Math.floor(details.coalescedInputCount ?? 0));
		this.droppedInputEstimate += Math.max(0, Math.floor(details.droppedInputEstimate ?? 0));
	}

	recordDomMutations(count: number): void {
		this.domMutationCount += Math.max(0, Math.floor(count));
	}

	scheduleFrame(
		requestFrame: (cb: (ts: number) => void) => number | null,
		onFrame: (session: InteractionSession, timestamp: number) => void,
		now: () => number,
	): void {
		if (this.frameHandle !== null) return;
		this.lastRequestTime = now();
		this.frameHandle = requestFrame((timestamp) => {
			this.frameHandle = null;
			onFrame(this, timestamp);
		});
	}

	recordFrame(
		timestamp: number,
		now: () => number,
		profile: RuntimeFrameProfile,
		maxWorstFrames: number,
	): void {
		const currentNow = now();
		const resolvedTimestamp = Number.isFinite(timestamp) ? timestamp : currentNow;
		const rafDelay = this.lastRequestTime !== null ? Math.max(0, resolvedTimestamp - this.lastRequestTime) : 0;
		const eventLoopDelay = Math.max(0, currentNow - resolvedTimestamp);
		const syncWork = this.syncWorkAccMs;
		this.syncWorkAccMs = 0;

		let interval = profile.resolvedFrameBudgetMs;
		if (this.lastFrameTimestamp !== null && resolvedTimestamp > this.lastFrameTimestamp) {
			interval = resolvedTimestamp - this.lastFrameTimestamp;
		}
		this.lastFrameTimestamp = resolvedTimestamp;

		this.frameIntervals.push(interval);
		this.rafDelays.push(rafDelay);
		this.syncWorks.push(syncWork);
		this.eventLoopDelays.push(eventLoopDelay);

		if (interval > this.longestFrameGapMs) {
			this.longestFrameGapMs = interval;
		}

		let classification: FrameAnomalyClassification = 'normal';
		if (interval >= profile.thresholds.substantialStallMs) {
			classification = 'stall';
			this.lateFrameCount += 1;
			const missed = Math.max(1, Math.floor(interval / profile.resolvedFrameBudgetMs) - 1);
			this.missedFrameEstimate += missed;
		} else if (interval >= profile.thresholds.missedFrameGapMs) {
			classification = 'missed';
			this.lateFrameCount += 1;
			const missed = Math.max(1, Math.floor(interval / profile.resolvedFrameBudgetMs) - 1);
			this.missedFrameEstimate += missed;
		} else if (interval >= profile.thresholds.lateFrameGapMs) {
			classification = 'late';
			this.lateFrameCount += 1;
		}

		if (classification !== 'normal' || this.worstFrames.length < maxWorstFrames) {
			this.worstFrames.push({
				timestamp: round(resolvedTimestamp),
				intervalMs: round(interval),
				rafDelayMs: round(rafDelay),
				synchronousWorkMs: round(syncWork),
				eventLoopDelayMs: round(eventLoopDelay),
				classification,
			});
			this.worstFrames.sort((left, right) => right.intervalMs - left.intervalMs);
			if (this.worstFrames.length > maxWorstFrames) {
				this.worstFrames.length = maxWorstFrames;
			}
		}
	}

	resetSettleTimer(
		scheduleTimeout: (cb: () => void, ms: number) => number | null,
		cancelTimeout: (handle: number) => void,
		settleMs: number,
		onSettle: (session: InteractionSession) => void,
	): void {
		if (this.settleTimer !== null) {
			cancelTimeout(this.settleTimer);
		}
		this.settleTimer = scheduleTimeout(() => {
			this.settleTimer = null;
			onSettle(this);
		}, settleMs);
	}

	cancel(cancelFrame: (handle: number) => void, cancelTimeout: (handle: number) => void): void {
		if (this.frameHandle !== null) {
			cancelFrame(this.frameHandle);
			this.frameHandle = null;
		}
		if (this.settleTimer !== null) {
			cancelTimeout(this.settleTimer);
			this.settleTimer = null;
		}
	}

	buildSummary(now: number, profile: RuntimeFrameProfile): ContinuousInteractionSummary {
		return {
			scope: this.scope,
			interactionType: this.type,
			frameCount: this.frameIntervals.length,
			totalDurationMs: round(now - this.startedAt),
			lateFrameCount: this.lateFrameCount,
			missedFrameEstimate: this.missedFrameEstimate,
			longestFrameGapMs: round(this.longestFrameGapMs),
			frameIntervalsMs: percentiles(this.frameIntervals),
			rafSchedulingDelayMs: percentiles(this.rafDelays),
			synchronousWorkMs: percentiles(this.syncWorks),
			eventLoopDelayMs: percentiles(this.eventLoopDelays),
			inputEventCount: this.inputEventCount,
			coalescedInputCount: this.coalescedInputCount,
			droppedInputEstimate: this.droppedInputEstimate,
			domMutationCount: this.domMutationCount,
			worstFrames: [...this.worstFrames],
			frameBudgetMs: profile.resolvedFrameBudgetMs,
			resolvedRefreshHz: profile.resolvedRefreshHz,
		};
	}
}

export class ContinuousInteractionProfiler {
	private readonly now: () => number;
	private readonly requestFrame: (callback: (timestamp: number) => void) => number | null;
	private readonly cancelFrame: (handle: number) => void;
	private readonly scheduleTimeout: (callback: () => void, ms: number) => number | null;
	private readonly cancelTimeout: (handle: number) => void;
	private readonly getFrameProfile: () => RuntimeFrameProfile;
	private readonly settleTimeoutMs: number;
	private readonly maxHistorySummaries: number;
	private readonly maxWorstFrames: number;
	private readonly onSummary?: (summary: ContinuousInteractionSummary) => void;
	private readonly onSlowInteraction?: (summary: ContinuousInteractionSummary) => void;

	private activeSessions = new Map<ContinuousInteractionType, InteractionSession>();
	private recentSummaries: ContinuousInteractionSummary[] = [];
	private disposed = false;

	constructor(options: ContinuousInteractionProfilerOptions = {}) {
		this.now = options.now ?? defaultNow;
		this.requestFrame = options.requestFrame ?? defaultRequestFrame;
		this.cancelFrame = options.cancelFrame ?? defaultCancelFrame;
		this.scheduleTimeout = options.scheduleTimeout ?? defaultScheduleTimeout;
		this.cancelTimeout = options.cancelTimeout ?? defaultCancelTimeout;
		this.getFrameProfile = options.getFrameProfile ?? getRuntimeFrameProfile;
		this.settleTimeoutMs = options.settleTimeoutMs ?? DEFAULT_SETTLE_TIMEOUT_MS;
		this.maxHistorySummaries = options.maxHistorySummaries ?? MAX_HISTORY_SUMMARIES;
		this.maxWorstFrames = options.maxWorstFrames ?? MAX_WORST_FRAMES;
		this.onSummary = options.onSummary;
		this.onSlowInteraction = options.onSlowInteraction;
	}

	recordInteractionEvent(
		type: ContinuousInteractionType,
		details: ContinuousInteractionEventDetails = {},
	): void {
		if (this.disposed) return;
		let session = this.activeSessions.get(type);
		if (!session) {
			session = new InteractionSession(type, this.now());
			this.activeSessions.set(type, session);
			this.scheduleSessionFrame(session);
		}
		session.recordInput(details);
		session.resetSettleTimer(
			this.scheduleTimeout,
			this.cancelTimeout,
			this.settleTimeoutMs,
			(settledSession) => this.settleSession(settledSession),
		);
	}

	recordDomMutations(type: ContinuousInteractionType, count = 1): void {
		if (this.disposed) return;
		this.activeSessions.get(type)?.recordDomMutations(count);
	}

	measureSynchronousWork<T>(type: ContinuousInteractionType, work: () => T): T {
		if (this.disposed) return work();
		const session = this.activeSessions.get(type);
		if (!session) return work();
		const start = this.now();
		try {
			return work();
		} finally {
			session.addSynchronousWork(this.now() - start);
		}
	}

	private scheduleSessionFrame(session: InteractionSession): void {
		if (this.disposed || !this.activeSessions.has(session.type)) return;
		session.scheduleFrame(
			this.requestFrame,
			(currentSession, timestamp) => {
				if (this.disposed || !this.activeSessions.has(currentSession.type)) return;
				const profile = this.getFrameProfile();
				currentSession.recordFrame(timestamp, this.now, profile, this.maxWorstFrames);
				this.scheduleSessionFrame(currentSession);
			},
			this.now,
		);
	}

	private settleSession(session: InteractionSession): void {
		if (!this.activeSessions.has(session.type)) return;
		this.activeSessions.delete(session.type);
		session.cancel(this.cancelFrame, this.cancelTimeout);

		if (session.frameIntervals.length === 0) return;

		const profile = this.getFrameProfile();
		const summary = session.buildSummary(this.now(), profile);

		this.recentSummaries.push(summary);
		if (this.recentSummaries.length > this.maxHistorySummaries) {
			this.recentSummaries.shift();
		}

		if (summary.lateFrameCount > 0 || summary.missedFrameEstimate > 0) {
			this.onSlowInteraction?.(summary);
		}
		this.onSummary?.(summary);
	}

	getRecentSummaries(): ContinuousInteractionSummary[] {
		return [...this.recentSummaries];
	}

	getLatestSummary(type?: ContinuousInteractionType): ContinuousInteractionSummary | null {
		if (!type) {
			return this.recentSummaries.length > 0 ? this.recentSummaries[this.recentSummaries.length - 1]! : null;
		}
		for (let i = this.recentSummaries.length - 1; i >= 0; i--) {
			if (this.recentSummaries[i]!.interactionType === type) {
				return this.recentSummaries[i]!;
			}
		}
		return null;
	}

	dispose(): void {
		if (this.disposed) return;
		this.disposed = true;
		for (const session of this.activeSessions.values()) {
			session.cancel(this.cancelFrame, this.cancelTimeout);
		}
		this.activeSessions.clear();
	}
}

let activeContinuousProfiler: ContinuousInteractionProfiler | null = null;

export function getSharedContinuousInteractionProfiler(): ContinuousInteractionProfiler {
	if (!activeContinuousProfiler) {
		activeContinuousProfiler = new ContinuousInteractionProfiler({
			onSummary: (summary) => {
				debugLog('performance', `Continuous interaction summary: ${summary.interactionType}`, summary);
			},
			onSlowInteraction: (summary) => {
				warnLog('performance', `Slow continuous interaction detected: ${summary.interactionType}`, {
					scope: summary.scope,
					interactionType: summary.interactionType,
					frameCount: summary.frameCount,
					lateFrames: summary.lateFrameCount,
					missedEstimate: summary.missedFrameEstimate,
					longestGapMs: summary.longestFrameGapMs,
					p95IntervalMs: summary.frameIntervalsMs.p95,
					maxIntervalMs: summary.frameIntervalsMs.max,
					maxSyncWorkMs: summary.synchronousWorkMs.max,
					maxRafDelayMs: summary.rafSchedulingDelayMs.max,
					worstFrames: summary.worstFrames,
				});
			},
		});
	}
	return activeContinuousProfiler;
}

export function resetSharedContinuousInteractionProfiler(): void {
	activeContinuousProfiler?.dispose();
	activeContinuousProfiler = null;
}
