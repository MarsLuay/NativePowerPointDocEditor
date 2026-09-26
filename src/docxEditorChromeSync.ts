/**
 * DOCX editor chrome is reconciled from a MutationObserver on title and
 * aria-label. Normalization writes those same attributes, so the observer
 * must ignore plugin-authored mutations or it schedules another animation
 * frame forever.
 */

export const EDITOR_CHROME_OBSERVED_ATTRIBUTES = ['title', 'aria-label'] as const;

const STORM_WINDOW_MS = 1000;
const STORM_THRESHOLD = 8;

export interface ChromeObservation {
	suspend(): void;
	resume(): void;
	dispose(): void;
}

export interface EditorChromeStormSummary {
	observerCallbacks: number;
	syncCount: number;
	pluginAuthoredMutations: number;
	elapsedMs: number;
	attributeNames: readonly string[];
	lastSyncChangedDom: boolean;
}

export interface EditorChromeReconciler {
	schedule(): void;
	run(): void;
	dispose(): void;
}

export interface EditorChromeReconcilerOptions {
	observe(listener: () => void): ChromeObservation;
	sync(): boolean | void;
	requestFrame(callback: () => void): number;
	cancelFrame(handle: number): void;
	syncTarget?: Node | null;
	now?: () => number;
	onStorm?: (summary: EditorChromeStormSummary) => void;
	stormThreshold?: number;
	stormWindowMs?: number;
}

export function setAttributeIfChanged(element: Element, name: string, value: string): boolean {
	if (element.getAttribute(name) === value) {
		return false;
	}
	element.setAttribute(name, value);
	return true;
}

export function removeAttributeIfPresent(element: Element, name: string): boolean {
	if (!element.hasAttribute(name)) {
		return false;
	}
	element.removeAttribute(name);
	return true;
}

export function observeEditorChromeAttributes(
	target: Node,
	listener: (records: MutationRecord[]) => void,
): ChromeObservation {
	const Observer = target.ownerDocument?.defaultView?.MutationObserver;
	if (typeof Observer !== 'function') {
		return { suspend() {}, resume() {}, dispose() {} };
	}

	const observer = new Observer(listener);
	const options: MutationObserverInit = {
		childList: true,
		subtree: true,
		attributes: true,
		attributeFilter: [...EDITOR_CHROME_OBSERVED_ATTRIBUTES],
	};
	let suspended = false;
	let disposed = false;
	observer.observe(target, options);

	return {
		suspend() {
			suspended = true;
			observer.disconnect();
		},
		resume() {
			if (disposed || !suspended) {
				return;
			}
			suspended = false;
			observer.observe(target, options);
		},
		dispose() {
			disposed = true;
			suspended = true;
			observer.disconnect();
		},
	};
}

export function createEditorChromeReconciler(options: EditorChromeReconcilerOptions): EditorChromeReconciler {
	const now = options.now ?? (() => Date.now());
	const stormThreshold = options.stormThreshold ?? STORM_THRESHOLD;
	const stormWindowMs = options.stormWindowMs ?? STORM_WINDOW_MS;
	let disposed = false;
	let queued = false;
	let syncing = false;
	let frame: number | null = null;
	let windowStartedAt = now();
	let observerCallbacks = 0;
	let syncCount = 0;
	let pluginAuthoredMutations = 0;
	let lastSyncChangedDom = false;
	let stormLogged = false;

	const resetStormWindow = (timestamp: number) => {
		windowStartedAt = timestamp;
		observerCallbacks = 0;
		syncCount = 0;
		pluginAuthoredMutations = 0;
		stormLogged = false;
	};

	const noteStormActivity = () => {
		const timestamp = now();
		if (timestamp - windowStartedAt > stormWindowMs) {
			resetStormWindow(timestamp);
		}
		if (stormLogged || !options.onStorm) {
			return;
		}
		if (observerCallbacks < stormThreshold && syncCount < stormThreshold) {
			return;
		}
		stormLogged = true;
		options.onStorm({
			observerCallbacks,
			syncCount,
			pluginAuthoredMutations,
			elapsedMs: Math.max(0, timestamp - windowStartedAt),
			attributeNames: EDITOR_CHROME_OBSERVED_ATTRIBUTES,
			lastSyncChangedDom,
		});
	};

	const observation = options.observe(() => {
		observerCallbacks += 1;
		if (syncing || disposed) {
			if (syncing) {
				pluginAuthoredMutations += 1;
			}
			noteStormActivity();
			return;
		}
		noteStormActivity();
		schedule();
	});

	const measureSyncChange = (): boolean => {
		if (!options.syncTarget) {
			return options.sync() === true;
		}
		const Observer = options.syncTarget.ownerDocument?.defaultView?.MutationObserver;
		if (typeof Observer !== 'function') {
			return options.sync() === true;
		}
		let changed = false;
		const probe = new Observer(() => {
			changed = true;
		});
		probe.observe(options.syncTarget, {
			childList: true,
			subtree: true,
			attributes: true,
			attributeFilter: [...EDITOR_CHROME_OBSERVED_ATTRIBUTES],
		});
		try {
			const explicit = options.sync();
			return explicit === true || changed;
		} finally {
			probe.disconnect();
		}
	};

	const run = () => {
		if (disposed || syncing) {
			return;
		}
		syncing = true;
		observation.suspend();
		try {
			lastSyncChangedDom = measureSyncChange();
		} finally {
			syncing = false;
			observation.resume();
		}
		syncCount += 1;
		noteStormActivity();
	};

	const schedule = () => {
		if (disposed || queued || syncing) {
			return;
		}
		queued = true;
		frame = options.requestFrame(() => {
			frame = null;
			queued = false;
			if (disposed) {
				return;
			}
			run();
		});
	};

	return {
		schedule,
		run,
		dispose() {
			disposed = true;
			queued = false;
			if (frame !== null) {
				options.cancelFrame(frame);
				frame = null;
			}
			observation.dispose();
		},
	};
}
