export type NativePowerPointDocEditorLogLevel = 'debug' | 'info' | 'warn' | 'error';

export interface NativePowerPointDocEditorLogEntry {
	time: string;
	level: NativePowerPointDocEditorLogLevel;
	area: string;
	message: string;
	data?: unknown;
}

const MAX_LOG_ENTRIES = 2000;
const LOG_PREFIX = '[Native PowerPoint Doc Editor]';

interface NativePowerPointDocEditorLogState {
	debugLoggingEnabled: boolean;
	entries: NativePowerPointDocEditorLogEntry[];
	totalEntries: number;
	droppedEntries: number;
}

let logState: NativePowerPointDocEditorLogState | null = null;

export interface NativePowerPointDocEditorLogStats {
	debugLoggingEnabled: boolean;
	maxRetainedEntries: number;
	retainedEntries: number;
	totalEntries: number;
	droppedEntries: number;
}

function getWindowWithLogs() {
	if (typeof window === 'undefined') {
		return null;
	}

	return window as Window & {
		nativePowerPointDocEditorDebugLogs?: NativePowerPointDocEditorLogEntry[];
		nativePowerPointDocEditorDebugLogging?: boolean;
	};
}

function getLogState() {
	if (!logState) {
		const logsWindow = getWindowWithLogs();
		const existingEntries = Array.isArray(logsWindow?.nativePowerPointDocEditorDebugLogs)
			? logsWindow.nativePowerPointDocEditorDebugLogs.slice(-MAX_LOG_ENTRIES)
			: [];
		logState = {
			debugLoggingEnabled: logsWindow?.nativePowerPointDocEditorDebugLogging === true,
			entries: existingEntries,
			totalEntries: existingEntries.length,
			droppedEntries: 0,
		};
	}

	return logState;
}

type NativePowerPointDocEditorLogSink = (entry: NativePowerPointDocEditorLogEntry) => void;

let logSink: NativePowerPointDocEditorLogSink | null = null;

export function setNativePowerPointDocEditorLogSink(sink: NativePowerPointDocEditorLogSink | null) {
	logSink = sink;
}

function syncWindowLogState() {
	const logsWindow = getWindowWithLogs();
	if (!logsWindow) {
		return;
	}

	const state = getLogState();
	logsWindow.nativePowerPointDocEditorDebugLogging = state.debugLoggingEnabled;
	logsWindow.nativePowerPointDocEditorDebugLogs = state.entries;
}

function shouldPrint(level: NativePowerPointDocEditorLogLevel) {
	return getLogState().debugLoggingEnabled || level === 'warn' || level === 'error';
}

function retentionPriority(level: NativePowerPointDocEditorLogLevel) {
	return level === 'debug' ? 0
		: level === 'info' ? 1
			: level === 'warn' ? 2
				: 3;
}

function writeConsole(level: NativePowerPointDocEditorLogLevel, area: string, message: string, data?: unknown) {
	if (!shouldPrint(level)) {
		return;
	}

	const consoleMethod = level === 'debug' ? console.debug
		: level === 'info' ? console.debug
			: level === 'warn' ? console.warn
				: console.error;

	if (data === undefined) {
		consoleMethod.call(console, `${LOG_PREFIX} ${area}: ${message}`);
	} else {
		consoleMethod.call(console, `${LOG_PREFIX} ${area}: ${message}`, data);
	}
}

function normalizeLogData(data: unknown, seen = new WeakSet<object>(), depth = 0): unknown {
	if (data instanceof Error) {
		return {
			name: data.name,
			message: data.message,
			stack: data.stack,
		};
	}

	if (data === null || typeof data !== 'object') {
		return data;
	}

	if (seen.has(data)) {
		return '[Circular]';
	}
	if (depth >= 6) {
		return '[Max depth]';
	}
	seen.add(data);

	if (Array.isArray(data)) {
		return data.map((value) => normalizeLogData(value, seen, depth + 1));
	}

	if (data instanceof Date) {
		return data.toISOString();
	}

	const prototype = Object.getPrototypeOf(data) as object | null;
	if (prototype === Object.prototype || prototype === null) {
		const normalized: Record<string, unknown> = {};
		for (const [key, value] of Object.entries(data)) {
			normalized[key] = normalizeLogData(value, seen, depth + 1);
		}
		return normalized;
	}

	return data;
}

export function configureNativePowerPointDocEditorLogger(enabled: boolean) {
	getLogState().debugLoggingEnabled = enabled;
	syncWindowLogState();
}

export function logNativePowerPointDocEditor(level: NativePowerPointDocEditorLogLevel, area: string, message: string, data?: unknown) {
	const state = getLogState();
	// Debug events are intentionally opt-in. Retaining them while debug logging is
	// disabled both spends work on hot paths and evicts the warnings/errors needed
	// to diagnose a real failure.
	if (level === 'debug' && !state.debugLoggingEnabled) {
		return;
	}
	const entry: NativePowerPointDocEditorLogEntry = {
		time: new Date().toISOString(),
		level,
		area,
		message,
		data: normalizeLogData(data),
	};

	state.totalEntries += 1;
	if (state.entries.length >= MAX_LOG_ENTRIES) {
		// A long debug session must not erase the warning/error that explains the
		// failure. Evict the oldest entry from the lowest available severity tier.
		const incomingPriority = retentionPriority(entry.level);
		let evictedIndex = -1;
		for (let priority = 0; priority <= incomingPriority && evictedIndex < 0; priority += 1) {
			evictedIndex = state.entries.findIndex((candidate) => retentionPriority(candidate.level) === priority);
		}
		if (evictedIndex >= 0) {
			state.entries.splice(evictedIndex, 1);
		}
		state.droppedEntries += 1;
	}
	if (state.entries.length < MAX_LOG_ENTRIES) {
		state.entries.push(entry);
	}
	syncWindowLogState();

	if (logSink) {
		try {
			logSink(entry);
		} catch {
			// Never let a log sink failure break the operation being logged.
		}
	}

	writeConsole(level, area, message, data);
}

export function debugLog(area: string, message: string, data?: unknown) {
	logNativePowerPointDocEditor('debug', area, message, data);
}

export function logPptxAction(area: string, op: string, data?: Record<string, unknown>) {
	debugLog(area, 'PowerPoint action started', { ...data, op });
}

export function infoLog(area: string, message: string, data?: unknown) {
	logNativePowerPointDocEditor('info', area, message, data);
}

export function warnLog(area: string, message: string, data?: unknown) {
	logNativePowerPointDocEditor('warn', area, message, data);
}

export function errorLog(area: string, message: string, data?: unknown) {
	logNativePowerPointDocEditor('error', area, message, data);
}

export function getNativePowerPointDocEditorLogSnapshot() {
	return getLogState().entries.slice();
}

export function getNativePowerPointDocEditorLogStats(): NativePowerPointDocEditorLogStats {
	const state = getLogState();
	return {
		debugLoggingEnabled: state.debugLoggingEnabled,
		maxRetainedEntries: MAX_LOG_ENTRIES,
		retainedEntries: state.entries.length,
		totalEntries: state.totalEntries,
		droppedEntries: state.droppedEntries,
	};
}
