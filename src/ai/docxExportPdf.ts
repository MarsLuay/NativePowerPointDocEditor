import { TFile, normalizePath, type Vault } from 'obsidian';

import {
	getAvailableNumberedPath,
	getVaultFolderPrefix,
	sanitizeArtifactBaseName,
	writeVaultBinaryArtifact,
	type ArtifactWriteTarget,
} from '../export/artifactPaths';
import { debugLog } from '../logger';
import { AI_ERROR_CODES, createAiError, isAiErrorDetail, type AiErrorDetail } from './errors';

export type DocxExportPdfConflict = 'replace' | 'keep-both';

export interface DocxExportPdfOptions {
	/** Vault-relative output path. Defaults to <sourceBasename>.pdf beside the source. */
	outputPath?: string;
	/** Conflict policy when outputPath already exists. Default: keep-both. */
	conflict?: DocxExportPdfConflict;
}

export interface DocxExportPdfResult {
	ok: boolean;
	path?: string;
	bytes?: number;
	errors: AiErrorDetail[];
}

export function resolveDocxExportPdfOutputTarget(
	vault: Vault,
	sourceFile: TFile,
	options: DocxExportPdfOptions = {},
): ArtifactWriteTarget {
	const conflict = options.conflict ?? 'keep-both';
	const folderPrefix = getVaultFolderPrefix(sourceFile.parent?.path);
	const defaultBase = sanitizeArtifactBaseName(sourceFile.basename, 'document');
	const requestedPath = normalizePath(
		options.outputPath?.trim() ? options.outputPath.trim() : `${folderPrefix}${defaultBase}.pdf`,
	);
	if (!requestedPath.toLowerCase().endsWith('.pdf')) {
		throw createAiError(
			AI_ERROR_CODES.SCHEMA_INVALID,
			`outputPath must end with .pdf (got ${requestedPath}).`,
			{ path: requestedPath, field: 'outputPath' },
		);
	}

	const exists = (candidate: string): boolean => Boolean(vault.getAbstractFileByPath(candidate));
	const existing = vault.getAbstractFileByPath(requestedPath);
	if (!existing) return { path: requestedPath, existingFile: null, replace: false };
	if (!(existing instanceof TFile)) {
		throw createAiError(
			AI_ERROR_CODES.FILE_EXISTS,
			`Output path exists and is not a file: ${requestedPath}.`,
			{ path: requestedPath },
		);
	}
	if (conflict === 'replace') return { path: requestedPath, existingFile: existing, replace: true };
	return { path: getAvailableNumberedPath(requestedPath, exists), existingFile: null, replace: false };
}

export async function writeDocxExportPdfArtifact(
	vault: Vault,
	sourceFile: TFile,
	bytes: ArrayBuffer,
	options: DocxExportPdfOptions = {},
): Promise<{ path: string; bytes: number }> {
	const target = resolveDocxExportPdfOutputTarget(vault, sourceFile, options);
	const written = await writeVaultBinaryArtifact(vault, target, bytes);
	debugLog('export', 'AI NPDE DOCX PDF export written', {
		op: 'ai-export-docx-pdf',
		sourcePath: sourceFile.path,
		targetPath: written.path,
		bytes: bytes.byteLength,
		replaced: target.replace,
	});
	return { path: written.path, bytes: bytes.byteLength };
}

export function toDocxExportPdfFailure(error: unknown, path?: string): DocxExportPdfResult {
	if (isAiErrorDetail(error)) return { ok: false, errors: [error] };
	return {
		ok: false,
		errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, String(error), path ? { path } : {})],
	};
}
