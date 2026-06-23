export function scheduleIdleWork(callback: () => void, options: { timeout?: number } = {}): () => void {
	const timeout = options.timeout ?? 2000;
	const win = window as Window & {
		requestIdleCallback?: (cb: IdleRequestCallback, opts?: IdleRequestOptions) => number;
		cancelIdleCallback?: (id: number) => void;
	};

	if (typeof win.requestIdleCallback === 'function') {
		const idleId = win.requestIdleCallback(() => callback(), { timeout });
		return () => win.cancelIdleCallback?.(idleId);
	}

	const timeoutId = window.setTimeout(callback, Math.min(timeout, 250));
	return () => window.clearTimeout(timeoutId);
}
