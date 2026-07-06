import { debugLog, warnLog } from './logger';
import { monotonicNow } from './loadTrace';

const LOG_PREFIX = '[Native PowerPoint Doc Editor]';

export function logLifecycleStep(step: string, data?: Record<string, unknown>) {
	const payload = { step, ...data };
	debugLog('lifecycle', step, payload);
	// Always mirror to the devtools console so logs survive main-thread stalls.
	console.warn(`${LOG_PREFIX} lifecycle: ${step}`, payload);
}

export function traceSyncStep<T>(step: string, run: () => T, data?: Record<string, unknown>): T {
	const startedAt = monotonicNow();
	logLifecycleStep(`${step}:start`, data);
	try {
		return run();
	} finally {
		const durationMs = Math.round((monotonicNow() - startedAt) * 10) / 10;
		logLifecycleStep(`${step}:done`, { ...data, durationMs });
		if (durationMs >= 250) {
			warnLog('lifecycle', `Slow sync step: ${step}`, { ...data, durationMs });
		}
	}
}

export function createObservedMutationObserver(
	name: string,
	callback: MutationCallback,
): MutationObserver {
	let mutationCount = 0;
	let windowStart = monotonicNow();
	let invocationCount = 0;

	return new MutationObserver((records, observer) => {
		invocationCount += 1;
		mutationCount += records.length;
		const now = monotonicNow();
		if (now - windowStart >= 1000) {
			if (mutationCount >= 40 || invocationCount >= 20) {
				warnLog('observer', `High mutation activity: ${name}`, {
					mutationCount,
					invocationCount,
					windowMs: Math.round(now - windowStart),
				});
				console.warn(`${LOG_PREFIX} observer storm: ${name}`, {
					mutationCount,
					invocationCount,
				});
			}
			mutationCount = 0;
			invocationCount = 0;
			windowStart = now;
		}

		callback(records, observer);
	});
}

export function startOpenHeartbeat(scope: string, context: () => Record<string, unknown>): () => void {
	const startedAt = monotonicNow();
	const intervalId = window.setInterval(() => {
		const elapsedMs = Math.round((monotonicNow() - startedAt) * 10) / 10;
		const payload = { scope, elapsedMs, ...context() };
		warnLog('load', `${scope}: heartbeat`, payload);
		console.warn(`${LOG_PREFIX} load heartbeat: ${scope}`, payload);
	}, 2000);

	return () => window.clearInterval(intervalId);
}
