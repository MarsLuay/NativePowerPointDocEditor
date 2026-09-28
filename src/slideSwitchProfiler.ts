export type SlideSwitchStage =
	| 'selection'
	| 'cache-lookup'
	| 'package-xml'
	| 'render'
	| 'svg-sanitize'
	| 'svg-parse'
	| 'font-readiness'
	| 'dom-swap'
	| 'postprocess';

export type SlideSwitchCacheState = 'warm' | 'cold' | 'unknown';
export type SlideSwitchDistance = 'adjacent' | 'distant';
export type SlideSwitchStatus = 'stable' | 'failed' | 'superseded';

export interface SlideSwitchStartOptions {
	fromSlide: number;
	toSlide: number;
	reason: string;
	slideCount: number;
	cacheState?: SlideSwitchCacheState;
	now?: number;
}

export interface SlideSwitchStageTimings {
	selection: number;
	cacheLookup: number | null;
	packageXml: number | null;
	render: number | null;
	svgSanitize: number | null;
	svgParse: number | null;
	fontReadiness: number | null;
	domSwap: number | null;
	postprocess: number | null;
}

export interface SlideSwitchMeasurement {
	id: number;
	fromSlide: number;
	toSlide: number;
	distanceSlides: number;
	distance: SlideSwitchDistance;
	reason: string;
	slideCount: number;
	cacheState: SlideSwitchCacheState;
	startedAt: number;
	visibleAt: number;
	inputToVisibleMs: number;
	frameSchedulingDelayMs: number;
	stageTimingsMs: SlideSwitchStageTimings;
	status: SlideSwitchStatus;
}

export interface SlideSwitchTimingStats {
	count: number;
	p50: number;
	p95: number;
	max: number;
}

export interface SlideSwitchProfilerSummary {
	sampleCount: number;
	retainedSampleCount: number;
	timing: SlideSwitchTimingStats;
	byCacheAndDistance: Record<string, SlideSwitchTimingStats>;
	worstSwitches: SlideSwitchMeasurement[];
}

interface ActiveSlideSwitch {
	id: number;
	fromSlide: number;
	toSlide: number;
	distanceSlides: number;
	distance: SlideSwitchDistance;
	reason: string;
	slideCount: number;
	cacheState: SlideSwitchCacheState;
	startedAt: number;
	frameAt: number | null;
	finished: boolean;
	status: SlideSwitchStatus;
	stageTimingsMs: SlideSwitchStageTimings;
}

export interface SlideSwitchProfilerOptions {
	maxSamples?: number;
	maxWorstSwitches?: number;
	now?: () => number;
	onComplete?: (measurement: SlideSwitchMeasurement) => void;
}

const DEFAULT_MAX_SAMPLES = 80;
const DEFAULT_MAX_WORST_SWITCHES = 10;

function finiteDuration(value: number): number {
	return Number.isFinite(value) && value >= 0 ? value : 0;
}

function percentile(values: number[], fraction: number): number {
	if (values.length === 0) return 0;
	const sorted = [...values].sort((a, b) => a - b);
	return sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * fraction) - 1))] ?? 0;
}

function stats(values: number[]): SlideSwitchTimingStats {
	return {
		count: values.length,
		p50: percentile(values, 0.5),
		p95: percentile(values, 0.95),
		max: values.length > 0 ? Math.max(...values) : 0,
	};
}

function stageKey(stage: SlideSwitchStage): keyof SlideSwitchStageTimings {
	if (stage === 'cache-lookup') return 'cacheLookup';
	if (stage === 'svg-sanitize') return 'svgSanitize';
	if (stage === 'svg-parse') return 'svgParse';
	if (stage === 'font-readiness') return 'fontReadiness';
	if (stage === 'dom-swap') return 'domSwap';
	if (stage === 'postprocess') return 'postprocess';
	if (stage === 'package-xml') return 'packageXml';
	return stage;
}

function createStageTimings(): SlideSwitchStageTimings {
	return {
		selection: 0,
		cacheLookup: null,
		packageXml: null,
		render: null,
		svgSanitize: null,
		svgParse: null,
		fontReadiness: null,
		domSwap: null,
		postprocess: null,
	};
}

function cacheDistanceKey(sample: Pick<SlideSwitchMeasurement, 'cacheState' | 'distance'>): string {
	return `${sample.cacheState}-${sample.distance}`;
}

export class SlideSwitchProfiler {
	private readonly maxSamples: number;
	private readonly maxWorstSwitches: number;
	private readonly now: () => number;
	private readonly onComplete?: (measurement: SlideSwitchMeasurement) => void;
	private measurements: SlideSwitchMeasurement[] = [];
	private active: ActiveSlideSwitch | null = null;
	private nextId = 1;
	private totalCompleted = 0;

	constructor(options: SlideSwitchProfilerOptions = {}) {
		this.maxSamples = Math.max(1, Math.floor(options.maxSamples ?? DEFAULT_MAX_SAMPLES));
		this.maxWorstSwitches = Math.max(1, Math.floor(options.maxWorstSwitches ?? DEFAULT_MAX_WORST_SWITCHES));
		this.now = options.now ?? (() => performance.now());
		this.onComplete = options.onComplete;
	}

	begin(options: SlideSwitchStartOptions): number {
		this.active = {
			id: this.nextId++,
			fromSlide: Math.max(0, Math.floor(options.fromSlide)),
			toSlide: Math.max(0, Math.floor(options.toSlide)),
			distanceSlides: Math.abs(Math.floor(options.toSlide) - Math.floor(options.fromSlide)),
			distance: Math.abs(options.toSlide - options.fromSlide) === 1 ? 'adjacent' : 'distant',
			reason: options.reason,
			slideCount: Math.max(0, Math.floor(options.slideCount)),
			cacheState: options.cacheState ?? 'unknown',
			startedAt: options.now ?? this.now(),
			frameAt: null,
			finished: false,
			status: 'stable',
			stageTimingsMs: createStageTimings(),
		};
		return this.active.id;
	}

	setCacheState(cacheState: SlideSwitchCacheState): void {
		if (this.active) this.active.cacheState = cacheState;
	}

	recordStage(stage: SlideSwitchStage, durationMs: number): void {
		if (!this.active) return;
		const key = stageKey(stage);
		const duration = finiteDuration(durationMs);
		const previous = this.active.stageTimingsMs[key];
		this.active.stageTimingsMs[key] = previous === null ? duration : previous + duration;
	}

	markFrame(at = this.now()): SlideSwitchMeasurement | null {
		if (!this.active) return null;
		this.active.frameAt = at;
		return this.tryComplete(at);
	}

	finish(status: SlideSwitchStatus = 'stable', at = this.now()): SlideSwitchMeasurement | null {
		if (!this.active) return null;
		this.active.finished = true;
		this.active.status = status;
		if (status !== 'stable') {
			return this.tryComplete(at);
		}
		return this.tryComplete(this.active.frameAt ?? at);
	}

	private tryComplete(visibleAt: number): SlideSwitchMeasurement | null {
		if (!this.active || !this.active.finished || (this.active.status === 'stable' && this.active.frameAt === null)) return null;
		const active = this.active;
		this.active = null;
		const measurement: SlideSwitchMeasurement = {
			id: active.id,
			fromSlide: active.fromSlide,
			toSlide: active.toSlide,
			distanceSlides: active.distanceSlides,
			distance: active.distance,
			reason: active.reason,
			slideCount: active.slideCount,
			cacheState: active.cacheState,
			startedAt: active.startedAt,
			visibleAt,
			inputToVisibleMs: finiteDuration(visibleAt - active.startedAt),
			frameSchedulingDelayMs: active.frameAt === null
				? 0
				: finiteDuration(active.frameAt - (active.startedAt + Object.values(active.stageTimingsMs).filter((value): value is number => value !== null).reduce((sum, value) => sum + value, 0))),
			stageTimingsMs: { ...active.stageTimingsMs },
			status: active.status,
		};
		this.totalCompleted += 1;
		this.measurements.push(measurement);
		if (this.measurements.length > this.maxSamples) {
			this.measurements.splice(0, this.measurements.length - this.maxSamples);
		}
		this.onComplete?.(measurement);
		return measurement;
	}

	getSummary(): SlideSwitchProfilerSummary {
		const byCacheAndDistance: Record<string, SlideSwitchTimingStats> = {};
		const groups = new Map<string, number[]>();
		for (const measurement of this.measurements) {
			const key = cacheDistanceKey(measurement);
			const values = groups.get(key) ?? [];
			values.push(measurement.inputToVisibleMs);
			groups.set(key, values);
		}
		for (const [key, values] of groups) byCacheAndDistance[key] = stats(values);

		return {
			sampleCount: this.totalCompleted,
			retainedSampleCount: this.measurements.length,
			timing: stats(this.measurements.map((sample) => sample.inputToVisibleMs)),
			byCacheAndDistance,
			worstSwitches: [...this.measurements]
				.sort((left, right) => right.inputToVisibleMs - left.inputToVisibleMs)
				.slice(0, this.maxWorstSwitches)
				.map((sample) => ({ ...sample, stageTimingsMs: { ...sample.stageTimingsMs } })),
		};
	}

	reset(): void {
		this.measurements = [];
		this.active = null;
		this.nextId = 1;
		this.totalCompleted = 0;
	}
}

export function createSlideSwitchProfiler(options: SlideSwitchProfilerOptions = {}): SlideSwitchProfiler {
	return new SlideSwitchProfiler(options);
}

export function summarizeSlideSwitchForLog(measurement: SlideSwitchMeasurement): Omit<SlideSwitchMeasurement, 'startedAt' | 'visibleAt'> {
	return {
		id: measurement.id,
		fromSlide: measurement.fromSlide,
		toSlide: measurement.toSlide,
		distanceSlides: measurement.distanceSlides,
		distance: measurement.distance,
		reason: measurement.reason,
		slideCount: measurement.slideCount,
		cacheState: measurement.cacheState,
		inputToVisibleMs: measurement.inputToVisibleMs,
		frameSchedulingDelayMs: measurement.frameSchedulingDelayMs,
		stageTimingsMs: measurement.stageTimingsMs,
		status: measurement.status,
	};
}
