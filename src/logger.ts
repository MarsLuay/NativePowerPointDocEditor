export type DocxidianLogLevel = 'debug' | 'info' | 'warn' | 'error';

export interface DocxidianLogEntry {
	time: string;
	level: DocxidianLogLevel;
	area: string;
	message: string;
	data?: unknown;
}

const MAX_LOG_ENTRIES = 2000;
const LOG_PREFIX = '[Native PowerPoint Doc Editor]';

interface DocxidianLogState {
	debugLoggingEnabled: boolean;
	entries: DocxidianLogEntry[];
	totalEntries: number;
	droppedEntries: number;
}

let logState: DocxidianLogState | null = null;

export interface DocxidianLogStats {
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
		docxidianDebugLogs?: DocxidianLogEntry[];
		docxidianDebugLogging?: boolean;
	};
}

function getLogState() {
	if (!logState) {
		const logsWindow = getWindowWithLogs();
		const existingEntries = Array.isArray(logsWindow?.docxidianDebugLogs)
			? logsWindow.docxidianDebugLogs.slice(-MAX_LOG_ENTRIES)
			: [];
		logState = {
			debugLoggingEnabled: logsWindow?.docxidianDebugLogging === true,
			entries: existingEntries,
			totalEntries: existingEntries.length,
			droppedEntries: 0,
		};
	}

	return logState;
}

type DocxidianLogSink = (entry: DocxidianLogEntry) => void;

let logSink: DocxidianLogSink | null = null;

export function setDocxidianLogSink(sink: DocxidianLogSink | null) {
	logSink = sink;
}

function syncWindowLogState() {
	const logsWindow = getWindowWithLogs();
	if (!logsWindow) {
		return;
	}

	const state = getLogState();
	logsWindow.docxidianDebugLogging = state.debugLoggingEnabled;
	logsWindow.docxidianDebugLogs = state.entries;
}

function shouldPrint(level: DocxidianLogLevel) {
	return getLogState().debugLoggingEnabled || level === 'warn' || level === 'error';
}

function writeConsole(level: DocxidianLogLevel, area: string, message: string, data?: unknown) {
	if (!shouldPrint(level)) {
		return;
	}

	const consoleMethod = level === 'debug' ? console.debug
		: level === 'info' ? console.info
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

	const prototype = Object.getPrototypeOf(data);
	if (prototype === Object.prototype || prototype === null) {
		const normalized: Record<string, unknown> = {};
		for (const [key, value] of Object.entries(data)) {
			normalized[key] = normalizeLogData(value, seen, depth + 1);
		}
		return normalized;
	}

	return data;
}

export function configureDocxidianLogger(enabled: boolean) {
	getLogState().debugLoggingEnabled = enabled;
	syncWindowLogState();
}

export function logDocxidian(level: DocxidianLogLevel, area: string, message: string, data?: unknown) {
	const state = getLogState();
	const entry: DocxidianLogEntry = {
		time: new Date().toISOString(),
		level,
		area,
		message,
		data: normalizeLogData(data),
	};

	state.entries.push(entry);
	state.totalEntries += 1;
	if (state.entries.length > MAX_LOG_ENTRIES) {
		const overflow = state.entries.length - MAX_LOG_ENTRIES;
		state.entries.splice(0, overflow);
		state.droppedEntries += overflow;
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
	logDocxidian('debug', area, message, data);
}

export function infoLog(area: string, message: string, data?: unknown) {
	logDocxidian('info', area, message, data);
}

export function warnLog(area: string, message: string, data?: unknown) {
	logDocxidian('warn', area, message, data);
}

export function errorLog(area: string, message: string, data?: unknown) {
	logDocxidian('error', area, message, data);
}

export function getDocxidianLogSnapshot() {
	return getLogState().entries.slice();
}

export function getDocxidianLogStats(): DocxidianLogStats {
	const state = getLogState();
	return {
		debugLoggingEnabled: state.debugLoggingEnabled,
		maxRetainedEntries: MAX_LOG_ENTRIES,
		retainedEntries: state.entries.length,
		totalEntries: state.totalEntries,
		droppedEntries: state.droppedEntries,
	};
}
