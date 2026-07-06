import { debugLog, infoLog, warnLog } from './logger';

export function monotonicNow(): number {
	if (typeof performance !== 'undefined' && typeof performance.now === 'function') {
		return performance.now();
	}

	return Date.now();
}

function roundMs(value: number): number {
	return Math.round(value * 10) / 10;
}

export interface LoadTrace {
	mark(phase: string, data?: Record<string, unknown>): void;
	finish(message: string, data?: Record<string, unknown>): void;
}

export function createLoadTrace(scope: string, context?: Record<string, unknown>): LoadTrace {
	const startedAt = monotonicNow();
	let previousAt = startedAt;
	const phases: Array<{ phase: string; sinceStartMs: number; sincePreviousMs: number }> = [];

	const mark = (phase: string, data?: Record<string, unknown>) => {
		const now = monotonicNow();
		const sinceStartMs = roundMs(now - startedAt);
		const sincePreviousMs = roundMs(now - previousAt);
		previousAt = now;
		phases.push({ phase, sinceStartMs, sincePreviousMs });
		const payload = {
			...context,
			...data,
			sinceStartMs,
			sincePreviousMs,
		};
		debugLog('load', `${scope}: ${phase}`, payload);
		console.warn(`[Native PowerPoint Doc Editor] load: ${scope}: ${phase}`, payload);
		if (sincePreviousMs >= 200 || sinceStartMs >= 1000) {
			warnLog('load', `${scope}: slow phase ${phase}`, payload);
		}
	};

	const finish = (message: string, data?: Record<string, unknown>) => {
		const totalMs = roundMs(monotonicNow() - startedAt);
		const payload = {
			...context,
			...data,
			totalMs,
			phases,
		};
		debugLog('load', `${scope}: ${message}`, payload);
		console.warn(`[Native PowerPoint Doc Editor] load: ${scope}: ${message}`, payload);
		if (totalMs >= 1500) {
			infoLog('load', `${scope}: slow load (${totalMs}ms)`, payload);
			warnLog('load', `${scope}: slow load (${totalMs}ms)`, payload);
		}
	};

	return { mark, finish };
}
