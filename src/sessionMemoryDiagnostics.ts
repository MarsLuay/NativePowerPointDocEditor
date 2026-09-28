export type MemoryProviderType = 'v8-process' | 'performance-memory' | 'unsupported';

export interface MemoryCapabilityReport {
	supported: boolean;
	provider: MemoryProviderType;
	usedBytes: number | null;
	totalBytes: number | null;
	limitBytes: number | null;
	detail?: string;
}

export interface SessionResourceCounters {
	mountedDocxViews: number;
	mountedPptxViews: number;
	activeTimers: number;
	activeAnimationFrames: number;
	activeMutationObservers: number;
	activeResizeObservers: number;
	thumbnailCacheEntries: number;
	domNodeCount: number | null;
	registeredListeners: number;
}

export interface SessionResourceSnapshot {
	timestamp: number;
	label: string;
	memory: MemoryCapabilityReport;
	resources: SessionResourceCounters;
}

export interface SessionLeakDelta {
	baselineLabel: string;
	targetLabel: string;
	retainedDocxViews: number;
	retainedPptxViews: number;
	retainedTimers: number;
	retainedFrames: number;
	retainedObservers: number;
	retainedDomNodes: number | null;
	heapGrowthBytes: number | null;
	isClean: boolean;
	reasons: string[];
}

export function detectMemoryCapability(env?: {
	process?: { memoryUsage?: () => NodeJS.MemoryUsage };
	performance?: { memory?: { usedJSHeapSize: number; totalJSHeapSize: number; jsHeapSizeLimit: number } };
}): MemoryCapabilityReport {
	const p = env?.performance ?? (typeof performance !== 'undefined' ? (performance as unknown as { memory?: { usedJSHeapSize: number; totalJSHeapSize: number; jsHeapSizeLimit: number } }) : undefined);
	if (p && p.memory && typeof p.memory.usedJSHeapSize === 'number' && p.memory.usedJSHeapSize > 0) {
		return {
			supported: true,
			provider: 'performance-memory',
			usedBytes: p.memory.usedJSHeapSize,
			totalBytes: p.memory.totalJSHeapSize,
			limitBytes: p.memory.jsHeapSizeLimit,
		};
	}

	const proc = env?.process ?? (typeof process !== 'undefined' ? process : undefined);
	if (proc && typeof proc.memoryUsage === 'function') {
		try {
			const mem = proc.memoryUsage();
			if (typeof mem.heapUsed === 'number' && mem.heapUsed > 0) {
				return {
					supported: true,
					provider: 'v8-process',
					usedBytes: mem.heapUsed,
					totalBytes: mem.heapTotal,
					limitBytes: null,
				};
			}
		} catch {
			// Fall through to unsupported
		}
	}

	return {
		supported: false,
		provider: 'unsupported',
		usedBytes: null,
		totalBytes: null,
		limitBytes: null,
		detail: 'Platform does not expose JavaScript heap metrics (e.g. WKWebView on iOS/macOS).',
	};
}

class SessionResourceRegistry {
	private docxViews = new Set<WeakRef<object>>();
	private pptxViews = new Set<WeakRef<object>>();
	private activeTimersCount = 0;
	private activeAnimationFramesCount = 0;
	private activeMutationObserversCount = 0;
	private activeResizeObserversCount = 0;
	private thumbnailCacheSize = 0;
	private activeListenersCount = 0;

	private prune(set: Set<WeakRef<object>>): number {
		let live = 0;
		for (const ref of Array.from(set)) {
			if (ref.deref() === undefined) {
				set.delete(ref);
			} else {
				live += 1;
			}
		}
		return live;
	}

	registerView(scope: 'docx' | 'pptx', instance: object): () => void {
		const targetSet = scope === 'docx' ? this.docxViews : this.pptxViews;
		const ref = new WeakRef(instance);
		targetSet.add(ref);
		return () => {
			targetSet.delete(ref);
		};
	}

	registerTimer(): () => void {
		this.activeTimersCount += 1;
		let disposed = false;
		return () => {
			if (!disposed) {
				disposed = true;
				this.activeTimersCount = Math.max(0, this.activeTimersCount - 1);
			}
		};
	}

	registerAnimationFrame(): () => void {
		this.activeAnimationFramesCount += 1;
		let disposed = false;
		return () => {
			if (!disposed) {
				disposed = true;
				this.activeAnimationFramesCount = Math.max(0, this.activeAnimationFramesCount - 1);
			}
		};
	}

	registerObserver(type: 'mutation' | 'resize'): () => void {
		if (type === 'mutation') {
			this.activeMutationObserversCount += 1;
		} else {
			this.activeResizeObserversCount += 1;
		}
		let disposed = false;
		return () => {
			if (!disposed) {
				disposed = true;
				if (type === 'mutation') {
					this.activeMutationObserversCount = Math.max(0, this.activeMutationObserversCount - 1);
				} else {
					this.activeResizeObserversCount = Math.max(0, this.activeResizeObserversCount - 1);
				}
			}
		};
	}

	setThumbnailCacheEntries(count: number): void {
		this.thumbnailCacheSize = Math.max(0, count);
	}

	registerListener(): () => void {
		this.activeListenersCount += 1;
		let disposed = false;
		return () => {
			if (!disposed) {
				disposed = true;
				this.activeListenersCount = Math.max(0, this.activeListenersCount - 1);
			}
		};
	}

	getCounters(): SessionResourceCounters {
		let domNodeCount: number | null = null;
		if (typeof document !== 'undefined' && typeof document.getElementsByTagName === 'function') {
			domNodeCount = document.getElementsByTagName('*').length;
		}

		return {
			mountedDocxViews: this.prune(this.docxViews),
			mountedPptxViews: this.prune(this.pptxViews),
			activeTimers: this.activeTimersCount,
			activeAnimationFrames: this.activeAnimationFramesCount,
			activeMutationObservers: this.activeMutationObserversCount,
			activeResizeObservers: this.activeResizeObserversCount,
			thumbnailCacheEntries: this.thumbnailCacheSize,
			domNodeCount,
			registeredListeners: this.activeListenersCount,
		};
	}

	reset(): void {
		this.docxViews.clear();
		this.pptxViews.clear();
		this.activeTimersCount = 0;
		this.activeAnimationFramesCount = 0;
		this.activeMutationObserversCount = 0;
		this.activeResizeObserversCount = 0;
		this.thumbnailCacheSize = 0;
		this.activeListenersCount = 0;
	}
}

export const sessionResourceRegistry = new SessionResourceRegistry();

export function captureSessionSnapshot(
	label: string,
	options?: {
		env?: Parameters<typeof detectMemoryCapability>[0];
		customCounters?: Partial<SessionResourceCounters>;
	},
): SessionResourceSnapshot {
	const memory = detectMemoryCapability(options?.env);
	const baseCounters = sessionResourceRegistry.getCounters();
	const resources: SessionResourceCounters = {
		...baseCounters,
		...options?.customCounters,
	};

	return {
		timestamp: Date.now(),
		label,
		memory,
		resources,
	};
}

export function compareSessionSnapshots(
	baseline: SessionResourceSnapshot,
	target: SessionResourceSnapshot,
	options: {
		allowedDocxLeak?: number;
		allowedPptxLeak?: number;
		allowedTimerLeak?: number;
		allowedObserverLeak?: number;
		allowedDomGrowth?: number;
	} = {},
): SessionLeakDelta {
	const retainedDocxViews = Math.max(0, target.resources.mountedDocxViews - baseline.resources.mountedDocxViews);
	const retainedPptxViews = Math.max(0, target.resources.mountedPptxViews - baseline.resources.mountedPptxViews);
	const retainedTimers = Math.max(0, target.resources.activeTimers - baseline.resources.activeTimers);
	const retainedFrames = Math.max(0, target.resources.activeAnimationFrames - baseline.resources.activeAnimationFrames);
	const retainedObservers = Math.max(
		0,
		(target.resources.activeMutationObservers + target.resources.activeResizeObservers) -
		(baseline.resources.activeMutationObservers + baseline.resources.activeResizeObservers),
	);

	let retainedDomNodes: number | null = null;
	if (baseline.resources.domNodeCount !== null && target.resources.domNodeCount !== null) {
		retainedDomNodes = Math.max(0, target.resources.domNodeCount - baseline.resources.domNodeCount);
	}

	let heapGrowthBytes: number | null = null;
	if (baseline.memory.supported && target.memory.supported && baseline.memory.usedBytes !== null && target.memory.usedBytes !== null) {
		heapGrowthBytes = target.memory.usedBytes - baseline.memory.usedBytes;
	}

	const reasons: string[] = [];
	if (retainedDocxViews > (options.allowedDocxLeak ?? 0)) {
		reasons.push(`Retained DOCX views: ${retainedDocxViews}`);
	}
	if (retainedPptxViews > (options.allowedPptxLeak ?? 0)) {
		reasons.push(`Retained PPTX views: ${retainedPptxViews}`);
	}
	if (retainedTimers > (options.allowedTimerLeak ?? 0)) {
		reasons.push(`Retained timers: ${retainedTimers}`);
	}
	if (retainedFrames > 0) {
		reasons.push(`Retained animation frames: ${retainedFrames}`);
	}
	if (retainedObservers > (options.allowedObserverLeak ?? 0)) {
		reasons.push(`Retained observers: ${retainedObservers}`);
	}
	if (retainedDomNodes !== null && retainedDomNodes > (options.allowedDomGrowth ?? 20)) {
		reasons.push(`Retained DOM nodes: ${retainedDomNodes}`);
	}

	const isClean = reasons.length === 0;

	return {
		baselineLabel: baseline.label,
		targetLabel: target.label,
		retainedDocxViews,
		retainedPptxViews,
		retainedTimers,
		retainedFrames,
		retainedObservers,
		retainedDomNodes,
		heapGrowthBytes,
		isClean,
		reasons,
	};
}

export interface EnduranceTestOptions {
	cycles?: number;
	onCycle?: (cycle: number) => Promise<void> | void;
}

export interface EnduranceTestReport {
	cycles: number;
	baseline: SessionResourceSnapshot;
	final: SessionResourceSnapshot;
	leakDelta: SessionLeakDelta;
	passed: boolean;
}

export async function runSessionEnduranceTest(options: EnduranceTestOptions = {}): Promise<EnduranceTestReport> {
	const cycles = Math.max(1, options.cycles ?? 5);
	const baseline = captureSessionSnapshot('baseline');

	for (let i = 1; i <= cycles; i++) {
		if (options.onCycle) {
			await options.onCycle(i);
		}
	}

	const final = captureSessionSnapshot(`after-${cycles}-cycles`);
	const leakDelta = compareSessionSnapshots(baseline, final);

	return {
		cycles,
		baseline,
		final,
		leakDelta,
		passed: leakDelta.isClean,
	};
}
