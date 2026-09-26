import { isPowerPointExtension } from '../powerpoint/extensions';
import type { DocumentOp } from './types';

export interface DescribeClipboardRequest {
	path?: string;
}

export interface ApplyClipboardRequest {
	path?: string;
	ops?: DocumentOp[];
	dryRun?: boolean;
}

export interface SaveClipboardRequest {
	path?: string;
}

export type UndoClipboardRequest = SaveClipboardRequest;

export type RedoClipboardRequest = SaveClipboardRequest;

export interface ValidateClipboardRequest {
	ops?: DocumentOp[];
}

export function parseClipboardJson(text: string): unknown {
	const trimmed = text.trim();
	if (!trimmed) {
		throw new Error('Clipboard is empty.');
	}
	return JSON.parse(trimmed) as unknown;
}

export function isRecord(value: unknown): value is Record<string, unknown> {
	return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

export function resolveDocumentPath(
	requestPath: string | undefined,
	activePath: string | null,
): string | null {
	if (typeof requestPath === 'string' && requestPath.trim().length > 0) {
		return requestPath.trim();
	}
	return activePath;
}

export function getActiveDocumentPath(activeFile: { path: string; extension: string } | null): string | null {
	if (!activeFile) return null;
	const extension = activeFile.extension.toLowerCase();
	if (isPowerPointExtension(extension) || extension === 'docx') {
		return activeFile.path;
	}
	return null;
}

export function parseDescribeRequest(value: unknown): DescribeClipboardRequest {
	if (!isRecord(value)) {
		return {};
	}
	return {
		path: typeof value.path === 'string' ? value.path : undefined,
	};
}

export function parseApplyRequest(value: unknown): ApplyClipboardRequest {
	if (!isRecord(value)) {
		throw new Error('Apply payload must be a JSON object.');
	}
	if (!Array.isArray(value.ops)) {
		throw new Error('Apply payload must include an ops array.');
	}
	return {
		path: typeof value.path === 'string' ? value.path : undefined,
		ops: value.ops as DocumentOp[],
		dryRun: value.dryRun === true,
	};
}

export function parseSaveRequest(value: unknown): SaveClipboardRequest {
	return parsePathClipboardRequest(value);
}

export function parseUndoRequest(value: unknown): UndoClipboardRequest {
	return parsePathClipboardRequest(value);
}

export function parseRedoRequest(value: unknown): RedoClipboardRequest {
	return parsePathClipboardRequest(value);
}

function parsePathClipboardRequest(value: unknown): SaveClipboardRequest {
	if (!isRecord(value)) {
		return {};
	}
	return {
		path: typeof value.path === 'string' ? value.path : undefined,
	};
}

export function parseValidateRequest(value: unknown): ValidateClipboardRequest {
	if (!isRecord(value)) {
		throw new Error('Validate payload must be a JSON object.');
	}
	if (!Array.isArray(value.ops)) {
		throw new Error('Validate payload must include an ops array.');
	}
	return {
		ops: value.ops as DocumentOp[],
	};
}

export async function copyJsonToClipboard(value: unknown): Promise<void> {
	await navigator.clipboard.writeText(JSON.stringify(value, null, 2));
}

export async function readClipboardJson(): Promise<unknown> {
	return parseClipboardJson(await navigator.clipboard.readText());
}
