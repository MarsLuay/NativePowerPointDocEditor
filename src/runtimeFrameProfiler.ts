export type RuntimeFrameConfidence = 'insufficient' | 'low' | 'stable';
export type RuntimeFrameThresholdSource = 'measured-raf' | 'platform-fallback';
export type RuntimeFrameAppMode = 'mobile' | 'desktop';

export interface RuntimeFrameThresholds {
	/** Fixed synchronous plugin-work budget; independent of display cadence. */
	synchronousWorkMs: number;
	lateFrameGapMs: number;
	missedFrameGapMs: number;
	substantialStallMs: number;
}

export interface RuntimeFrameProfile {
	measuredRefreshHz: number | null;
	measuredFrameBudgetMs: number | null;
	resolvedRefreshHz: number;
	resolvedFrameBudgetMs: number;
	sampleCount: number;
	confidence: RuntimeFrameConfidence;
	thresholdSource: RuntimeFrameThresholdSource;
	platform: string;
	appMode: RuntimeFrameAppMode;
	thresholds: RuntimeFrameThresholds;
}

export interface RuntimeFrameProfilerOptions {
	platform?: string;
	appMode?: RuntimeFrameAppMode;
	fallbackRefreshHz?: number;
	requestFrame?: (callback: (timestamp: number) => void) => number | null;
	cancelFrame?: (handle: number) => void;
	isVisible?: () => boolean;
	onProfileChange?: (profile: RuntimeFrameProfile) => void;
	minimumSamples?: number;
	stableSamples?: number;
	maxSamples?: number;
}

export const DEFAULT_RUNTIME_REFRESH_HZ = 60;
export const SYNCHRONOUS_WORK_THRESHOLD_MS = 8;
export const MINIMUM_RUNTIME_FRAME_SAMPLES = 8;
export const STABLE_RUNTIME_FRAME_SAMPLES = 30;
export const MAX_RUNTIME_FRAME_SAMPLES = 120;

const SAMPLE_WINDOW_FRAMES = 48;
const RESAMPLE_INTERVAL_MS = 15_000;
const HIDDEN_RETRY_INTERVAL_MS = 1_000;
const MIN_CLEAN_FRAME_INTERVAL_MS = 1.5;
const MAX_CLEAN_FRAME_INTERVAL_MS = 50;
const FRAME_CLUSTER_TOLERANCE_RATIO = 0.12;
const MIN_FRAME_CLUSTER_TOLERANCE_MS = 0.65;

function round(value: number): number {
	return Math.round(value * 100) / 100;
}

function median(values: readonly number[]): number {
	if (values.length === 0) return 0;
	const sorted = [...values].sort((left, right) => left - right);
	const middle = Math.floor(sorted.length / 2);
	return sorted.length % 2 === 0
		? (sorted[middle - 1]! + sorted[middle]!) / 2
		: sorted[middle]!;
}

/** Select the largest tight cadence cluster; ties favor the lower (base) cadence. */
function stableCluster(samples: readonly number[]): number[] {
	const sorted = [...samples].sort((left, right) => left - right);
	let best: number[] = [];
	for (const center of sorted) {
		const tolerance = Math.max(MIN_FRAME_CLUSTER_TOLERANCE_MS, center * FRAME_CLUSTER_TOLERANCE_RATIO);
		const cluster = sorted.filter((sample) => Math.abs(sample - center) <= tolerance);
		if (cluster.length > best.length || (cluster.length === best.length && center < (median(best) || Number.POSITIVE_INFINITY))) {
			best = cluster;
		}
	}
	return best;
}

function medianDeviationRatio(samples: readonly number[], center: number): number {
	if (samples.length === 0 || center <= 0) return Number.POSITIVE_INFINITY;
	return median(samples.map((sample) => Math.abs(sample - center))) / center;
}

function thresholds(frameBudgetMs: number): RuntimeFrameThresholds {
	return {
		synchronousWorkMs: SYNCHRONOUS_WORK_THRESHOLD_MS,
		lateFrameGapMs: round(frameBudgetMs * 1.5),
		missedFrameGapMs: round(frameBudgetMs * 2),
		substantialStallMs: round(frameBudgetMs * 3),
	};
}

function defaultRequestFrame(callback: (timestamp: number) => void): number | null {
	if (typeof window === 'undefined' || typeof window.requestAnimationFrame !== 'function') return null;
	return window.requestAnimationFrame(callback);
}

function defaultCancelFrame(handle: number): void {
	if (typeof window !== 'undefined' && typeof window.cancelAnimationFrame === 'function') {
		window.cancelAnimationFrame(handle);
	}
}

function defaultVisibility(): boolean {
	return typeof activeDocument === 'undefined' || activeDocument.visibilityState !== 'hidden';
}

function scheduleTimer(callback: () => void, delayMs: number): number | null {
	return typeof window === 'undefined' ? null : window.setTimeout(callback, delayMs);
}

function cancelTimer(handle: number): void {
	if (typeof window !== 'undefined') window.clearTimeout(handle);
}

export class RuntimeFrameProfiler {
	private readonly platform: string;
	private readonly appMode: RuntimeFrameAppMode;
	private readonly fallbackRefreshHz: number;
	private readonly requestFrame: (callback: (timestamp: number) => void) => number | null;
	private readonly cancelFrame: (handle: number) => void;
	private readonly isVisible: () => boolean;
	private readonly onProfileChange?: (profile: RuntimeFrameProfile) => void;
	private readonly minimumSamples: number;
	private readonly stableSamples: number;
	private readonly maxSamples: number;
	private readonly samples: number[] = [];
	private lastTimestamp: number | null = null;
	private samplesSinceProfileUpdate = 0;
	private profile: RuntimeFrameProfile;
	private profileKey: string;
	private running = false;
	private sampleWindowRemaining = 0;
	private frameHandle: number | null = null;
	private resampleTimer: number | null = null;

	constructor(options: RuntimeFrameProfilerOptions = {}) {
		this.platform = options.platform ?? 'unknown';
		this.appMode = options.appMode ?? 'desktop';
		this.fallbackRefreshHz = positiveFinite(options.fallbackRefreshHz)
			? options.fallbackRefreshHz
			: DEFAULT_RUNTIME_REFRESH_HZ;
		this.requestFrame = options.requestFrame ?? defaultRequestFrame;
		this.cancelFrame = options.cancelFrame ?? defaultCancelFrame;
		this.isVisible = options.isVisible ?? defaultVisibility;
		this.onProfileChange = options.onProfileChange;
		this.minimumSamples = Math.max(2, Math.floor(options.minimumSamples ?? MINIMUM_RUNTIME_FRAME_SAMPLES));
		this.stableSamples = Math.max(this.minimumSamples, Math.floor(options.stableSamples ?? STABLE_RUNTIME_FRAME_SAMPLES));
		this.maxSamples = Math.max(this.stableSamples, Math.floor(options.maxSamples ?? MAX_RUNTIME_FRAME_SAMPLES));
		this.profile = this.resolveProfile();
		this.profileKey = profileKey(this.profile);
	}

	getProfile(): RuntimeFrameProfile {
		return this.profile;
	}

	getThresholds(): RuntimeFrameThresholds {
		return this.profile.thresholds;
	}

	/** Direct deterministic input for tests and explicitly scheduled sample windows. */
	observeFrame(timestamp: number): RuntimeFrameProfile {
		if (!this.isVisible()) {
			if (this.samples.length > 0 || this.lastTimestamp !== null) this.reset();
			return this.profile;
		}
		if (!Number.isFinite(timestamp)) {
			this.lastTimestamp = null;
			return this.profile;
		}
		if (this.lastTimestamp !== null && timestamp > this.lastTimestamp) {
			const interval = timestamp - this.lastTimestamp;
			if (interval >= MIN_CLEAN_FRAME_INTERVAL_MS && interval <= MAX_CLEAN_FRAME_INTERVAL_MS) {
				this.samples.push(interval);
				if (this.samples.length > this.maxSamples) this.samples.shift();
				this.samplesSinceProfileUpdate += 1;
				if (this.samples.length <= this.minimumSamples || this.samplesSinceProfileUpdate >= 8) {
					this.recomputeProfile();
				}
			}
		}
		this.lastTimestamp = timestamp;
		return this.profile;
	}

	/** Avoid work in normal DOCX/PPTX rAF callbacks outside a bounded sample window. */
	observeScheduledFrame(timestamp: number): void {
		if (this.sampleWindowRemaining > 0) this.observeFrame(timestamp);
	}

	requestFrameForConsumer(callback: (timestamp: number) => void): number | null {
		return this.requestFrame(callback);
	}

	cancelFrameForConsumer(handle: number): void {
		this.cancelFrame(handle);
	}

	reset(): void {
		this.samples.length = 0;
		this.lastTimestamp = null;
		this.samplesSinceProfileUpdate = 0;
		this.recomputeProfile();
	}

	start(): void {
		if (this.running) return;
		this.running = true;
		this.startSampleWindow();
	}

	stop(): void {
		this.running = false;
		if (this.frameHandle !== null) this.cancelFrame(this.frameHandle);
		if (this.resampleTimer !== null) cancelTimer(this.resampleTimer);
		this.frameHandle = null;
		this.resampleTimer = null;
		this.sampleWindowRemaining = 0;
		this.lastTimestamp = null;
	}

	private startSampleWindow(): void {
		if (!this.running) return;
		if (!this.isVisible()) {
			this.reset();
			this.scheduleNextWindow(HIDDEN_RETRY_INTERVAL_MS);
			return;
		}
		this.lastTimestamp = null;
		this.sampleWindowRemaining = Math.min(SAMPLE_WINDOW_FRAMES, this.maxSamples);
		const sample = (timestamp: number): void => {
			this.frameHandle = null;
			if (!this.running) return;
			if (!this.isVisible()) {
				this.reset();
				this.scheduleNextWindow(HIDDEN_RETRY_INTERVAL_MS);
				return;
			}
			this.observeFrame(timestamp);
			this.sampleWindowRemaining -= 1;
			if (this.sampleWindowRemaining > 0) {
				this.frameHandle = this.requestFrame(sample);
			} else {
				this.lastTimestamp = null;
				this.scheduleNextWindow(RESAMPLE_INTERVAL_MS);
			}
		};
		this.frameHandle = this.requestFrame(sample);
	}

	private scheduleNextWindow(delayMs: number): void {
		if (!this.running) return;
		this.resampleTimer = scheduleTimer(() => {
			this.resampleTimer = null;
			this.startSampleWindow();
		}, delayMs);
	}

	private resolveProfile(): RuntimeFrameProfile {
		const cluster = stableCluster(this.samples);
		const measured = cluster.length >= this.minimumSamples;
		const measuredFrameBudgetMs = measured ? median(cluster) : null;
		const measuredRefreshHz = measuredFrameBudgetMs === null ? null : 1000 / measuredFrameBudgetMs;
		const resolvedRefreshHz = measuredRefreshHz ?? this.fallbackRefreshHz;
		const resolvedFrameBudgetMs = 1000 / resolvedRefreshHz;
		const confidence: RuntimeFrameConfidence = !measured
			? 'insufficient'
			: cluster.length >= this.stableSamples && medianDeviationRatio(cluster, measuredFrameBudgetMs!) <= 0.1
				? 'stable'
				: 'low';
		return {
			measuredRefreshHz: measuredRefreshHz === null ? null : round(measuredRefreshHz),
			measuredFrameBudgetMs: measuredFrameBudgetMs === null ? null : round(measuredFrameBudgetMs),
			resolvedRefreshHz: round(resolvedRefreshHz),
			resolvedFrameBudgetMs: round(resolvedFrameBudgetMs),
			sampleCount: cluster.length,
			confidence,
			thresholdSource: measured ? 'measured-raf' : 'platform-fallback',
			platform: this.platform,
			appMode: this.appMode,
			thresholds: thresholds(resolvedFrameBudgetMs),
		};
	}

	private recomputeProfile(): void {
		this.samplesSinceProfileUpdate = 0;
		const nextProfile = this.resolveProfile();
		const nextKey = profileKey(nextProfile);
		this.profile = nextProfile;
		if (nextKey === this.profileKey) return;
		this.profileKey = nextKey;
		this.onProfileChange?.(nextProfile);
	}
}

let activeProfiler: RuntimeFrameProfiler | null = null;

export function startRuntimeFrameProfiler(options: RuntimeFrameProfilerOptions = {}): RuntimeFrameProfiler {
	activeProfiler?.stop();
	activeProfiler = new RuntimeFrameProfiler(options);
	activeProfiler.start();
	return activeProfiler;
}

export function stopRuntimeFrameProfiler(): void {
	activeProfiler?.stop();
	activeProfiler = null;
}

export function getRuntimeFrameProfile(): RuntimeFrameProfile {
	return activeProfiler?.getProfile() ?? new RuntimeFrameProfiler().getProfile();
}

export function getRuntimeFrameThresholds(): RuntimeFrameThresholds {
	return getRuntimeFrameProfile().thresholds;
}

export type RuntimeFrameWindow = Pick<Window, 'requestAnimationFrame' | 'cancelAnimationFrame'>;
export type RuntimeFrameObserver = (timestamp: number, scheduledAt: number) => void;

const runtimeFrameObservers = new Set<RuntimeFrameObserver>();

/** Observe the shared editor rAF stream without adding per-frame log entries. */
export function subscribeRuntimeFrameObserver(observer: RuntimeFrameObserver): () => void {
	runtimeFrameObservers.add(observer);
	return () => runtimeFrameObservers.delete(observer);
}

function monotonicNow(): number {
	return typeof performance !== 'undefined' && typeof performance.now === 'function' ? performance.now() : Date.now();
}

/** Shared rAF entry point for DOCX and PPTX visual work. */
export function requestRuntimeFrame(
	callback: (timestamp: number) => void,
	frameWindow?: RuntimeFrameWindow,
): number | null {
	const request = frameWindow
		? (handler: (timestamp: number) => void) => frameWindow.requestAnimationFrame(handler)
		: (handler: (timestamp: number) => void) => activeProfiler?.requestFrameForConsumer(handler) ?? defaultRequestFrame(handler);
	const scheduledAt = monotonicNow();
	return request((timestamp) => {
		activeProfiler?.observeScheduledFrame(timestamp);
		for (const observer of runtimeFrameObservers) observer(timestamp, scheduledAt);
		callback(timestamp);
	});
}

export function cancelRuntimeFrame(handle: number | null, frameWindow?: RuntimeFrameWindow): void {
	if (handle === null) return;
	if (frameWindow) frameWindow.cancelAnimationFrame(handle);
	else if (activeProfiler) activeProfiler.cancelFrameForConsumer(handle);
	else defaultCancelFrame(handle);
}

export function detectRuntimePlatform(platform: {
	isMobile?: boolean;
	isMobileApp?: boolean;
	isIosApp?: boolean;
	isAndroidApp?: boolean;
	isMacOS?: boolean;
	isWin?: boolean;
	isLinux?: boolean;
} = {}): string {
	if (platform.isIosApp) return 'ios';
	if (platform.isAndroidApp) return 'android';
	if (platform.isMobileApp || platform.isMobile) return 'mobile';
	if (platform.isMacOS) return 'macos';
	if (platform.isWin) return 'windows';
	if (platform.isLinux) return 'linux';
	return 'desktop';
}

function positiveFinite(value: number | undefined): value is number {
	return typeof value === 'number' && Number.isFinite(value) && value > 0;
}

function profileKey(profile: RuntimeFrameProfile): string {
	return [Math.round(profile.resolvedRefreshHz / 2) * 2, profile.confidence, profile.thresholdSource].join(':');
}
