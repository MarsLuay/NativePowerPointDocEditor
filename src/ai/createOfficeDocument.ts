import type { Vault } from 'obsidian';
import { normalizePath } from 'obsidian';

import { buildBlankDocxArrayBuffer, buildBlankPptxArrayBuffer } from '../vault/blankOfficePackages';
import { applyReplaceBodyParagraphs } from './docxBodyParagraphs';
import { DocxPatchSession } from './docxPatchSession';
import { AI_ERROR_CODES, createAiError, type AiErrorDetail } from './errors';

export type CreateOfficeDocumentKind = 'docx' | 'pptx';

export interface CreateOfficeDocumentOptions {
	path: string;
	kind: CreateOfficeDocumentKind;
	/** DOCX only: replace body with these paragraphs after creating the blank package. */
	paragraphs?: string[];
	overwrite?: boolean;
}

export interface CreateOfficeDocumentResult {
	ok: boolean;
	path?: string;
	errors: AiErrorDetail[];
}

async function ensureParentFolders(vault: Vault, filePath: string): Promise<void> {
	const parts = normalizePath(filePath).split('/').filter(Boolean);
	parts.pop();
	let current = '';
	for (const part of parts) {
		current = current ? `${current}/${part}` : part;
		if (vault.getAbstractFileByPath(current) == null) {
			await vault.createFolder(current);
		}
	}
}

/**
 * Create a blank vault DOCX/PPTX for agent workflows.
 * Reuses the same blank packages as the New DOCX/PPTX file-menu actions.
 */
export async function createOfficeDocument(
	vault: Vault,
	options: CreateOfficeDocumentOptions,
): Promise<CreateOfficeDocumentResult> {
	const path = normalizePath(options.path);
	const kind = options.kind;
	if (kind !== 'docx' && kind !== 'pptx') {
		return {
			ok: false,
			errors: [createAiError(AI_ERROR_CODES.UNSUPPORTED_FORMAT, `Unsupported kind: ${String(kind)}.`, { path })],
		};
	}

	const extension = path.split('.').pop()?.toLowerCase() ?? '';
	if (extension !== kind) {
		return {
			ok: false,
			errors: [
				createAiError(
					AI_ERROR_CODES.SCHEMA_INVALID,
					`Path extension must be .${kind} for kind "${kind}".`,
					{ path, field: 'path' },
				),
			],
		};
	}

	if (options.paragraphs !== undefined && kind !== 'docx') {
		return {
			ok: false,
			errors: [
				createAiError(
					AI_ERROR_CODES.SCHEMA_INVALID,
					'paragraphs is only supported when kind is "docx".',
					{ path, field: 'paragraphs' },
				),
			],
		};
	}

	const existing = vault.getAbstractFileByPath(path);
	if (existing && options.overwrite !== true) {
		return {
			ok: false,
			errors: [
				createAiError(
					AI_ERROR_CODES.FILE_EXISTS,
					`File already exists: ${path}. Pass overwrite: true to replace it.`,
					{ path },
				),
			],
		};
	}

	try {
		await ensureParentFolders(vault, path);
		let buffer = kind === 'docx'
			? await buildBlankDocxArrayBuffer()
			: await buildBlankPptxArrayBuffer();

		if (kind === 'docx' && options.paragraphs) {
			const session = await DocxPatchSession.load(buffer);
			const nextXml = applyReplaceBodyParagraphs(session.getDocumentXml(), options.paragraphs);
			session.setDocumentXml(nextXml);
			buffer = await session.export();
		}

		if (existing) {
			await vault.modifyBinary(existing as never, buffer);
		} else {
			await vault.createBinary(path, buffer);
		}

		return { ok: true, path, errors: [] };
	} catch (error) {
		return {
			ok: false,
			errors: [
				createAiError(
					AI_ERROR_CODES.VALIDATION_FAILED,
					error instanceof Error ? error.message : String(error),
					{ path },
				),
			],
		};
	}
}
