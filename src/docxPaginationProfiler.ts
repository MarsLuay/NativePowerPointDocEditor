export type DocxReflowEditKind =
	| 'paragraph-insert'
	| 'paragraph-delete'
	| 'font-size'
	| 'table-edit'
	| 'image-insert'
	| 'image-resize'
	| 'page-break-change'
	| 'transaction';

export type DocxReflowEditLocation = 'top' | 'middle' | 'end';

export interface DocxReflowAffectedRange {
	from: number;
	to: number;
	documentSize: number;
	pageStart: number | null;
	pageEnd: number | null;
}

export interface DocxReflowMeasurement {
	id: number;
	editKind: DocxReflowEditKind;
	location: DocxReflowEditLocation;
	affectedRange: DocxReflowAffectedRange;
	startedAt: number;
	transactionAt: number;
	stableAt: number;
	totalMs: number;
	synchronousWorkMs: number;
	frameSchedulingDelayMs: number;
	paginationPasses: number;
	layoutCallbackCount: number;
	pagesRecalculated: number;
	initialPageCount: number | null;
	stablePageCount: number | null;
	status: 'stable' | 'timeout' | 'superseded';
}

export interface DocxReflowTimingStats {
	count: number;
	p50: number;
	p95: number;
	max: number;
}

export interface DocxPaginationProfilerSummary {
	sampleCount: number;
	retainedSampleCount: number;
	timing: DocxReflowTimingStats;
	synchronousWork: DocxReflowTimingStats;
	frameSchedulingDelay: DocxReflowTimingStats;
	paginationPasses: DocxReflowTimingStats;
	layoutCallbacks: DocxReflowTimingStats;
	pagesRecalculated: DocxReflowTimingStats;
	recentMeasurements: DocxReflowMeasurement[];
}

interface ActiveMeasurement {
	id: number;
	editKind: DocxReflowEditKind;
	location: DocxReflowEditLocation;
	affectedRange: DocxReflowAffectedRange;
	startedAt: number;
	transactionAt: number;
	synchronousWorkMs: number | null;
	paginationPasses: number;
	layoutCallbackCount: number;
	initialPageCount: number | null;
	lastPageCount: number | null;
	pagesRecalculated: number;
}

export interface DocxPaginationProfilerOptions {
	maxSamples?: number;
	now?: () => number;
}

export interface BeginDocxReflowOptions {
	editKind: DocxReflowEditKind;
	location: DocxReflowEditLocation;
	affectedRange: DocxReflowAffectedRange;
	initialPageCount?: number | null;
}

export interface CompleteDocxReflowOptions {
	stablePageCount?: number | null;
	pagesRecalculated?: number;
	status?: 'stable' | 'timeout';
	now?: number;
}

const DEFAULT_MAX_SAMPLES = 60;

function finiteOrZero(value: number): number {
	return Number.isFinite(value) && value >= 0 ? value : 0;
}

function percentile(values: number[], fraction: number): number {
	if (values.length === 0) return 0;
	const sorted = [...values].sort((a, b) => a - b);
	const index = Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * fraction) - 1));
	return sorted[index] ?? 0;
}

function stats(values: number[]): DocxReflowTimingStats {
	return {
		count: values.length,
		p50: percentile(values, 0.5),
		p95: percentile(values, 0.95),
		max: values.length > 0 ? Math.max(...values) : 0,
	};
}

function locationForPosition(position: number, documentSize: number): DocxReflowEditLocation {
	if (documentSize <= 0 || position <= documentSize / 3) return 'top';
	if (position >= (documentSize * 2) / 3) return 'end';
	return 'middle';
}

export function inferDocxReflowLocation(position: number, documentSize: number): DocxReflowEditLocation {
	return locationForPosition(position, documentSize);
}

export class DocxPaginationProfiler {
	private readonly maxSamples: number;
	private readonly now: () => number;
	private measurements: DocxReflowMeasurement[] = [];
	private active: ActiveMeasurement | null = null;
	private nextId = 1;
	private totalCompleted = 0;

	constructor(options: DocxPaginationProfilerOptions = {}) {
		this.maxSamples = Math.max(1, Math.floor(options.maxSamples ?? DEFAULT_MAX_SAMPLES));
		this.now = options.now ?? (() => performance.now());
	}

	beginEdit(options: BeginDocxReflowOptions, startedAt = this.now()): number {
		this.active = {
			id: this.nextId++,
			editKind: options.editKind,
			location: options.location,
			affectedRange: {
				from: Math.max(0, Math.floor(options.affectedRange.from)),
				to: Math.max(0, Math.floor(options.affectedRange.to)),
				documentSize: Math.max(0, Math.floor(options.affectedRange.documentSize)),
				pageStart: options.affectedRange.pageStart,
				pageEnd: options.affectedRange.pageEnd,
			},
			startedAt,
			transactionAt: startedAt,
			synchronousWorkMs: null,
			paginationPasses: 0,
			layoutCallbackCount: 0,
			initialPageCount: options.initialPageCount ?? null,
			lastPageCount: options.initialPageCount ?? null,
			pagesRecalculated: 0,
		};
		return this.active.id;
	}

	markSynchronousWorkComplete(at = this.now()): void {
		if (!this.active) return;
		this.active.synchronousWorkMs = finiteOrZero(at - this.active.transactionAt);
	}

	recordPaginationPass(pageCount?: number | null): void {
		if (!this.active) return;
		this.active.paginationPasses += 1;
		if (typeof pageCount === 'number' && Number.isFinite(pageCount)) {
			if (this.active.lastPageCount !== null && pageCount !== this.active.lastPageCount) {
				this.active.pagesRecalculated += Math.abs(pageCount - this.active.lastPageCount);
			}
			this.active.lastPageCount = pageCount;
		}
	}

	recordLayoutCallback(): void {
		if (!this.active) return;
		this.active.layoutCallbackCount += 1;
	}

	complete(options: CompleteDocxReflowOptions = {}): DocxReflowMeasurement | null {
		if (!this.active) return null;
		const active = this.active;
		this.active = null;
		const stableAt = options.now ?? this.now();
		const totalMs = finiteOrZero(stableAt - active.startedAt);
		const synchronousWorkMs = active.synchronousWorkMs ?? totalMs;
		const frameSchedulingDelayMs = finiteOrZero(totalMs - synchronousWorkMs);
		const pagesRecalculated = Math.max(0, Math.floor(options.pagesRecalculated ?? active.pagesRecalculated));
		const measurement: DocxReflowMeasurement = {
			id: active.id,
			editKind: active.editKind,
			location: active.location,
			affectedRange: active.affectedRange,
			startedAt: active.startedAt,
			transactionAt: active.transactionAt,
			stableAt,
			totalMs,
			synchronousWorkMs,
			frameSchedulingDelayMs,
			paginationPasses: active.paginationPasses,
			layoutCallbackCount: active.layoutCallbackCount,
			pagesRecalculated,
			initialPageCount: active.initialPageCount,
			stablePageCount: options.stablePageCount ?? active.lastPageCount,
			status: options.status ?? 'stable',
		};
		this.totalCompleted += 1;
		this.measurements.push(measurement);
		if (this.measurements.length > this.maxSamples) {
			this.measurements.splice(0, this.measurements.length - this.maxSamples);
		}
		return measurement;
	}

	cancel(): void {
		this.active = null;
	}

	getSummary(): DocxPaginationProfilerSummary {
		const measurements = this.measurements;
		return {
			sampleCount: this.totalCompleted,
			retainedSampleCount: measurements.length,
			timing: stats(measurements.map((sample) => sample.totalMs)),
			synchronousWork: stats(measurements.map((sample) => sample.synchronousWorkMs)),
			frameSchedulingDelay: stats(measurements.map((sample) => sample.frameSchedulingDelayMs)),
			paginationPasses: stats(measurements.map((sample) => sample.paginationPasses)),
			layoutCallbacks: stats(measurements.map((sample) => sample.layoutCallbackCount)),
			pagesRecalculated: stats(measurements.map((sample) => sample.pagesRecalculated)),
			recentMeasurements: measurements.map((sample) => ({
				...sample,
				affectedRange: { ...sample.affectedRange },
			})),
		};
	}

	reset(): void {
		this.measurements = [];
		this.active = null;
		this.totalCompleted = 0;
		this.nextId = 1;
	}
}

export function createDocxPaginationProfiler(options: DocxPaginationProfilerOptions = {}): DocxPaginationProfiler {
	return new DocxPaginationProfiler(options);
}

export function createAffectedDocxRange(
	from: number,
	to: number,
	documentSize: number,
	totalPages: number | null,
): DocxReflowAffectedRange {
	const safeDocumentSize = Math.max(0, Math.floor(documentSize));
	const safeFrom = Math.min(safeDocumentSize, Math.max(0, Math.floor(from)));
	const safeTo = Math.min(safeDocumentSize, Math.max(safeFrom, Math.floor(to)));
	const pageCount = totalPages !== null && Number.isFinite(totalPages) ? Math.max(0, Math.floor(totalPages)) : 0;
	const pageStart = pageCount > 0 ? Math.min(pageCount, Math.floor((safeFrom / Math.max(1, safeDocumentSize)) * pageCount) + 1) : null;
	const pageEnd = pageCount > 0 ? Math.min(pageCount, Math.floor((safeTo / Math.max(1, safeDocumentSize)) * pageCount) + 1) : null;
	return {
		from: safeFrom,
		to: safeTo,
		documentSize: safeDocumentSize,
		pageStart,
		pageEnd,
	};
}

export function summarizeDocxReflowForLog(summary: DocxPaginationProfilerSummary): Pick<DocxPaginationProfilerSummary, 'sampleCount' | 'retainedSampleCount' | 'timing' | 'synchronousWork' | 'frameSchedulingDelay' | 'paginationPasses' | 'layoutCallbacks' | 'pagesRecalculated'> {
	return {
		sampleCount: summary.sampleCount,
		retainedSampleCount: summary.retainedSampleCount,
		timing: summary.timing,
		synchronousWork: summary.synchronousWork,
		frameSchedulingDelay: summary.frameSchedulingDelay,
		paginationPasses: summary.paginationPasses,
		layoutCallbacks: summary.layoutCallbacks,
		pagesRecalculated: summary.pagesRecalculated,
	};
}
