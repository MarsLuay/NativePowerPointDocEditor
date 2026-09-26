import type {
	NativePowerPointDocEditorLogEntry,
	NativePowerPointDocEditorLogStats,
} from './logger';

export const MAX_COPIED_LOG_CHARACTERS = 32_000;
const MAX_METADATA_STRING_CHARACTERS = 2_048;

export type CopiedLogScope = 'all' | 'docx' | 'pptx';

export interface CopiedLogDiagnostics {
	obsidianVersion: string | null;
	obsidianApiVersion: string | null;
	platform: string | null;
	appMode: 'desktop' | 'mobile' | 'unknown';
	runtime: {
		electron: string | null;
		chromium: string | null;
		node: string | null;
	};
	userAgent: string | null;
	devicePixelRatio: number | null;
}

export interface CopiedLogPayloadInput {
	generatedAt: string;
	scope: CopiedLogScope;
	activeDocxPath?: string;
	plugin: {
		id: string;
		version: string;
		dir?: string;
	};
	settings: Record<string, unknown>;
	docxEditorBundle: string;
	logStats: NativePowerPointDocEditorLogStats;
	diagnostics: CopiedLogDiagnostics;
	logs: NativePowerPointDocEditorLogEntry[];
}

export interface CopiedLogPayload extends Omit<CopiedLogPayloadInput, 'logs'> {
	logs: NativePowerPointDocEditorLogEntry[];
	logRetention: {
		maxCharacters: number;
		truncated: boolean;
		retainedEntries: number;
	};
}

function truncateMetadataString(value: string): string;
function truncateMetadataString(value: undefined): undefined;
function truncateMetadataString(value: string | undefined): string | undefined;
function truncateMetadataString(value: string | undefined): string | undefined {
	if (value === undefined || value.length <= MAX_METADATA_STRING_CHARACTERS) return value;
	return `${value.slice(0, MAX_METADATA_STRING_CHARACTERS - 1)}…`;
}

function normalizeDiagnostics(diagnostics: CopiedLogDiagnostics): CopiedLogDiagnostics {
	return {
		obsidianVersion: diagnostics.obsidianVersion === null ? null : truncateMetadataString(diagnostics.obsidianVersion),
		obsidianApiVersion: diagnostics.obsidianApiVersion === null ? null : truncateMetadataString(diagnostics.obsidianApiVersion),
		platform: diagnostics.platform === null ? null : truncateMetadataString(diagnostics.platform),
		appMode: diagnostics.appMode,
		runtime: {
			electron: diagnostics.runtime.electron === null ? null : truncateMetadataString(diagnostics.runtime.electron),
			chromium: diagnostics.runtime.chromium === null ? null : truncateMetadataString(diagnostics.runtime.chromium),
			node: diagnostics.runtime.node === null ? null : truncateMetadataString(diagnostics.runtime.node),
		},
		userAgent: diagnostics.userAgent === null ? null : truncateMetadataString(diagnostics.userAgent),
		devicePixelRatio: Number.isFinite(diagnostics.devicePixelRatio) ? diagnostics.devicePixelRatio : null,
	};
}

function normalizeMetadata(input: CopiedLogPayloadInput): Omit<CopiedLogPayload, 'logs'> {
	return {
		generatedAt: input.generatedAt,
		scope: input.scope,
		...(input.activeDocxPath === undefined
			? {}
			: { activeDocxPath: truncateMetadataString(input.activeDocxPath) }),
		plugin: {
			id: truncateMetadataString(input.plugin.id),
			version: truncateMetadataString(input.plugin.version),
			dir: truncateMetadataString(input.plugin.dir),
		},
		settings: input.settings,
		docxEditorBundle: truncateMetadataString(input.docxEditorBundle),
		logStats: input.logStats,
		diagnostics: normalizeDiagnostics(input.diagnostics),
		logRetention: {
			maxCharacters: MAX_COPIED_LOG_CHARACTERS,
			truncated: false,
			retainedEntries: 0,
		},
	};
}

function serialize(payload: CopiedLogPayload): string {
	return JSON.stringify(payload, null, 2);
}

function withLogs(
	metadata: Omit<CopiedLogPayload, 'logs'>,
	logs: NativePowerPointDocEditorLogEntry[],
	sourceLogCount: number,
): CopiedLogPayload {
	return {
		...metadata,
		logRetention: {
			...metadata.logRetention,
			truncated: logs.length < sourceLogCount,
			retainedEntries: logs.length,
		},
		logs,
	};
}

function fitNewestLogEntry(
	metadata: Omit<CopiedLogPayload, 'logs'>,
	entry: NativePowerPointDocEditorLogEntry,
	sourceLogCount: number,
): CopiedLogPayload | null {
	const baseEntry = {
		time: entry.time,
		level: entry.level,
		area: entry.area,
		data: '[truncated]',
	};
	let low = 0;
	let high = entry.message.length;
	let best: NativePowerPointDocEditorLogEntry | null = null;

	while (low <= high) {
		const length = Math.floor((low + high) / 2);
		const candidate = {
			...baseEntry,
			message: length === 0 ? '' : entry.message.slice(-length),
		} satisfies NativePowerPointDocEditorLogEntry;
		const payload = withLogs(metadata, [candidate], sourceLogCount);
		if (serialize(payload).length <= MAX_COPIED_LOG_CHARACTERS) {
			best = candidate;
			low = length + 1;
		} else {
			high = length - 1;
		}
	}

	return best ? withLogs(metadata, [best], sourceLogCount) : null;
}

function minimalPayload(metadata: Omit<CopiedLogPayload, 'logs'>): CopiedLogPayload {
	return {
		generatedAt: metadata.generatedAt,
		scope: metadata.scope,
		plugin: metadata.plugin,
		settings: {},
		docxEditorBundle: metadata.docxEditorBundle,
		logStats: metadata.logStats,
		diagnostics: metadata.diagnostics,
		logRetention: {
			...metadata.logRetention,
			truncated: true,
			retainedEntries: 0,
		},
		logs: [],
	};
}

/**
 * Build a valid JSON payload while reserving space for copy-time diagnostics and
 * metadata. Logs are retained as the newest suffix, never the oldest prefix.
 */
export function buildCopiedLogPayload(input: CopiedLogPayloadInput): CopiedLogPayload {
	const metadata = normalizeMetadata(input);
	const logs = input.logs;
	let low = 0;
	let high = logs.length;
	let best = withLogs(metadata, [], logs.length);

	while (low <= high) {
		const count = Math.floor((low + high) / 2);
		const candidate = withLogs(metadata, count === 0 ? [] : logs.slice(-count), logs.length);
		if (serialize(candidate).length <= MAX_COPIED_LOG_CHARACTERS) {
			best = candidate;
			low = count + 1;
		} else {
			high = count - 1;
		}
	}

	if (serialize(best).length <= MAX_COPIED_LOG_CHARACTERS) {
		if (logs.length > 0 && best.logs.length === 0) {
			return fitNewestLogEntry(metadata, logs[logs.length - 1]!, logs.length) ?? best;
		}
		return best;
	}

	// Plugin-controlled metadata should fit comfortably. Keep the contract
	// machine-readable even if a future setting or runtime adds an unexpectedly
	// large value by omitting only the least essential settings object.
	const fallback = minimalPayload(metadata);
	if (serialize(fallback).length <= MAX_COPIED_LOG_CHARACTERS) {
		return fallback;
	}

	// The remaining fields are bounded above, so this is only a defensive final
	// shape for hostile host-provided metadata. It still preserves diagnostics and
	// the newest-log guarantee (there are no logs left to retain).
	return {
		generatedAt: truncateMetadataString(input.generatedAt),
		scope: input.scope,
		plugin: {
			id: truncateMetadataString(input.plugin.id),
			version: truncateMetadataString(input.plugin.version),
			dir: truncateMetadataString(input.plugin.dir),
		},
		settings: {},
		docxEditorBundle: truncateMetadataString(input.docxEditorBundle),
		logStats: {
			debugLoggingEnabled: input.logStats.debugLoggingEnabled,
			maxRetainedEntries: input.logStats.maxRetainedEntries,
			retainedEntries: input.logStats.retainedEntries,
			totalEntries: input.logStats.totalEntries,
			droppedEntries: input.logStats.droppedEntries,
		},
		diagnostics: normalizeDiagnostics(input.diagnostics),
		logRetention: {
			maxCharacters: MAX_COPIED_LOG_CHARACTERS,
			truncated: true,
			retainedEntries: 0,
		},
		logs: [],
	};
}

export function serializeCopiedLogPayload(input: CopiedLogPayloadInput): string {
	return serialize(buildCopiedLogPayload(input));
}
