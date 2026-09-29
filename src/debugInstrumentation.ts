import { debugLog, warnLog } from './logger';
import { monotonicNow } from './loadTrace';

const LOG_PREFIX = '[Native PowerPoint Doc Editor]';

export interface ObserverAmplificationRecord {
	action: string;
	durationMs: number;
	callbackCount: number;
	mutationCount: number;
	callbackWorkMs: number;
	sources: Record<string, { callbacks: number; mutations: number; workMs: number }>;
}

export interface ObserverAmplificationTrace {
	begin(action: string): void;
	record(source: string, mutationCount: number, callbackWorkMs: number): void;
	end(): ObserverAmplificationRecord | null;
}

export function createObserverAmplificationTrace(): ObserverAmplificationTrace {
	let active: {
		action: string;
		startedAt: number;
		callbackCount: number;
		mutationCount: number;
		callbackWorkMs: number;
		sources: Map<string, { callbacks: number; mutations: number; workMs: number }>;
	} | null = null;

	return {
		begin(action) {
			active = {
				action,
				startedAt: monotonicNow(),
				callbackCount: 0,
				mutationCount: 0,
				callbackWorkMs: 0,
				sources: new Map(),
			};
		},
		record(source, mutationCount, callbackWorkMs) {
			if (!active) return;
			active.callbackCount += 1;
			active.mutationCount += Math.max(0, mutationCount);
			active.callbackWorkMs += Math.max(0, callbackWorkMs);
			const current = active.sources.get(source) ?? { callbacks: 0, mutations: 0, workMs: 0 };
			if (active.sources.size < 8 || active.sources.has(source)) {
				current.callbacks += 1;
				current.mutations += Math.max(0, mutationCount);
				current.workMs += Math.max(0, callbackWorkMs);
				active.sources.set(source, current);
			}
		},
		end() {
			if (!active) return null;
			const record: ObserverAmplificationRecord = {
				action: active.action,
				durationMs: Math.round((monotonicNow() - active.startedAt) * 10) / 10,
				callbackCount: active.callbackCount,
				mutationCount: active.mutationCount,
				callbackWorkMs: Math.round(active.callbackWorkMs * 10) / 10,
				sources: Object.fromEntries([...active.sources.entries()].map(([source, value]) => [source, {
					callbacks: value.callbacks,
					mutations: value.mutations,
					workMs: Math.round(value.workMs * 10) / 10,
				}]))
			};
			active = null;
			if (record.callbackCount >= 20 || record.callbackWorkMs >= 16.7) {
				warnLog('observer', `Observer amplification: ${record.action}`, record);
			}
			return record;
		},
	};
}

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
	trace?: ObserverAmplificationTrace,
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

		const callbackStartedAt = monotonicNow();
		callback(records, observer);
		trace?.record(name, records.length, monotonicNow() - callbackStartedAt);
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
