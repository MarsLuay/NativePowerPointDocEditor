import { debugLog } from './logger';
import { monotonicNow } from './loadTrace';

/** Shared opt-in threshold for bounded local slow-operation diagnostics. */
export const SLOW_PERFORMANCE_THRESHOLD_MS = 100;

function roundMs(value: number): number {
	return Math.round(value * 10) / 10;
}

/**
 * Emit one normalized slow-operation event without collecting normal hot-path
 * traffic. Payloads are deliberately caller-supplied metadata so document
 * contents, buffers, paths, and credentials never enter performance logs.
 */
export function logSlowPerformance(
	area: string,
	message: string,
	durationMs: number,
	details: Record<string, unknown> = {},
): boolean {
	if (!Number.isFinite(durationMs) || durationMs <= SLOW_PERFORMANCE_THRESHOLD_MS) return false;
	debugLog(area, message, {
		...details,
		durationMs: roundMs(durationMs),
		thresholdMs: SLOW_PERFORMANCE_THRESHOLD_MS,
	});
	return true;
}

export function startPerformanceTimer(): () => number {
	const startedAt = monotonicNow();
	return () => Math.max(0, monotonicNow() - startedAt);
}
