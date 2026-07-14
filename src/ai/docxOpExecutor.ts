import type { Vault } from 'obsidian';
import { getImageMimeType } from '../PowerPointInsertModals';
import {
	getParagraphXml,
	getTableCellXmlFromPart,
	insertBlockAfterInPart,
	replaceParagraphXml,
	replaceTableCellXmlInPart,
} from './docxBlockResolver';
import { listReplaceTextPartPaths, resolvePartPath } from './docxParts';
import { addInlineImage, replaceInlineImage } from './docxMedia';
import type { DocxPatchSession } from './docxPatchSession';
import {
	buildEmptyTableXml,
	patchCellStyle,
	patchCellText,
	patchParagraphStyle,
	patchRunStyle,
	patchRunText,
	replacePartText,
	type DocxRunStylePatch,
} from './docxOoxmlWrite';
import { parseStableLocation } from './docxStableIds';
import { AI_ERROR_CODES, createAiError } from './errors';
import type { ApplyPreviewChange, DocumentOp } from './types';
import { readVaultBinaryFile } from './vaultBinary';

export interface DocxOpExecutionResult {
	changedIds: string[];
	preview: ApplyPreviewChange[];
	warnings: string[];
	documentXml: string;
}

export interface DocxOpExecutionContext {
	session: DocxPatchSession;
	vault: Vault;
	filePath: string;
	dryRun: boolean;
}

function asRecord(op: DocumentOp): Record<string, unknown> {
	return op as Record<string, unknown>;
}

function requireString(value: unknown, field: string): string {
	if (typeof value !== 'string') {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be a string.`, { field });
	}
	return value;
}

function requireNumber(value: unknown, field: string): number {
	if (typeof value !== 'number' || !Number.isFinite(value)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be a number.`, { field });
	}
	return value;
}

function asRunStylePatch(value: unknown): DocxRunStylePatch {
	if (!value || typeof value !== 'object') {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'style must be an object.', { field: 'style' });
	}
	const record = value as Record<string, unknown>;
	return {
		...(typeof record.bold === 'boolean' ? { bold: record.bold } : {}),
		...(typeof record.italic === 'boolean' ? { italic: record.italic } : {}),
		...(typeof record.underline === 'boolean' ? { underline: record.underline } : {}),
		...(typeof record.fontFamily === 'string' ? { fontFamily: record.fontFamily } : {}),
		...(typeof record.fontSizePt === 'number' ? { fontSizePt: record.fontSizePt } : {}),
		...(typeof record.color === 'string' || record.color === null ? { color: record.color as string | null } : {}),
	};
}

function rejectWriteOnlyExcludedId(id: string, field: string): void {
	if (id.startsWith('comments/')) {
		throw createAiError(
			AI_ERROR_CODES.VALIDATION_FAILED,
			'Comments are describe-only. trackChanges markup is not writable via AI ops.',
			{ field },
		);
	}
}

function getPartXmlForLocation(session: DocxPatchSession, location: ReturnType<typeof parseStableLocation>): string {
	if (!location) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Invalid stable id.', { field: 'blockId' });
	}
	const partPath = resolvePartPath(location);
	return session.getPartXml(partPath);
}

function setPartXmlForLocation(
	session: DocxPatchSession,
	location: NonNullable<ReturnType<typeof parseStableLocation>>,
	partXml: string,
): void {
	session.setPartXml(resolvePartPath(location), partXml);
}

export async function executeDocxOp(
	context: DocxOpExecutionContext,
	op: DocumentOp,
): Promise<DocxOpExecutionResult> {
	const record = asRecord(op);
	const opId = String(op.op);
	let documentXml = context.session.getDocumentXml();
	const changedIds: string[] = [];
	const preview: ApplyPreviewChange[] = [];
	const warnings: string[] = [];

	switch (opId) {
		case 'docx.setRunText': {
			const blockId = requireString(record.blockId, 'blockId');
			const runId = requireString(record.runId, 'runId');
			const text = requireString(record.text, 'text');
			rejectWriteOnlyExcludedId(runId, 'runId');
			const parsedRun = parseStableLocation(runId);
			if (!parsedRun || parsedRun.kind !== 'run') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid runId: ${runId}.`, { field: 'runId' });
			}
			const parsedBlock = parseStableLocation(blockId);
			if (
				!parsedBlock
				|| parsedBlock.kind !== 'paragraph'
				|| parsedBlock.part !== parsedRun.part
				|| parsedBlock.partNumber !== parsedRun.partNumber
				|| parsedBlock.paragraphIndex !== parsedRun.paragraphIndex
			) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `blockId ${blockId} does not match runId ${runId}.`, {
					field: 'blockId',
				});
			}
			let partXml = getPartXmlForLocation(context.session, parsedRun);
			const paragraphXml = getParagraphXml(partXml, parsedRun);
			const nextParagraphXml = patchRunText(paragraphXml, parsedRun.runIndex ?? 0, text);
			partXml = replaceParagraphXml(partXml, parsedRun, nextParagraphXml);
			setPartXmlForLocation(context.session, parsedRun, partXml);
			if (parsedRun.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(runId);
			preview.push({ id: runId, field: 'text', before: null, after: text });
			break;
		}
		case 'docx.setRunStyle': {
			const runId = requireString(record.runId, 'runId');
			const style = asRunStylePatch(record.style);
			rejectWriteOnlyExcludedId(runId, 'runId');
			const parsedRun = parseStableLocation(runId);
			if (!parsedRun || parsedRun.kind !== 'run') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid runId: ${runId}.`, { field: 'runId' });
			}
			let partXml = getPartXmlForLocation(context.session, parsedRun);
			const paragraphXml = getParagraphXml(partXml, parsedRun);
			const nextParagraphXml = patchRunStyle(paragraphXml, parsedRun.runIndex ?? 0, style);
			partXml = replaceParagraphXml(partXml, parsedRun, nextParagraphXml);
			setPartXmlForLocation(context.session, parsedRun, partXml);
			if (parsedRun.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(runId);
			preview.push({ id: runId, field: 'style', before: null, after: style });
			break;
		}
		case 'docx.setParagraphStyle': {
			const blockId = requireString(record.blockId, 'blockId');
			const style = record.style;
			if (!style || typeof style !== 'object') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'style must be an object.', { field: 'style' });
			}
			rejectWriteOnlyExcludedId(blockId, 'blockId');
			const parsed = parseStableLocation(blockId);
			if (!parsed || parsed.kind !== 'paragraph') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid blockId: ${blockId}.`, { field: 'blockId' });
			}
			let partXml = getPartXmlForLocation(context.session, parsed);
			const paragraphXml = getParagraphXml(partXml, parsed);
			const nextParagraphXml = patchParagraphStyle(paragraphXml, style as Record<string, unknown>);
			partXml = replaceParagraphXml(partXml, parsed, nextParagraphXml);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(blockId);
			preview.push({ id: blockId, field: 'style', before: null, after: style });
			break;
		}
		case 'docx.insertTable': {
			const afterBlockId = requireString(record.afterBlockId, 'afterBlockId');
			const rows = requireNumber(record.rows, 'rows');
			const cols = requireNumber(record.cols, 'cols');
			if (!Number.isInteger(rows) || rows < 1 || !Number.isInteger(cols) || cols < 1) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'rows and cols must be positive integers.', { field: 'rows' });
			}
			rejectWriteOnlyExcludedId(afterBlockId, 'afterBlockId');
			const anchor = parseStableLocation(afterBlockId);
			if (!anchor) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid afterBlockId: ${afterBlockId}.`, { field: 'afterBlockId' });
			}
			const tableXml = buildEmptyTableXml(rows, cols);
			let partXml = getPartXmlForLocation(context.session, anchor);
			partXml = insertBlockAfterInPart(partXml, afterBlockId, tableXml);
			setPartXmlForLocation(context.session, anchor, partXml);
			if (anchor.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(afterBlockId);
			preview.push({ id: afterBlockId, field: 'insertTable', before: null, after: { rows, cols } });
			break;
		}
		case 'docx.setCellText': {
			const cellId = requireString(record.cellId, 'cellId');
			const text = requireString(record.text, 'text');
			rejectWriteOnlyExcludedId(cellId, 'cellId');
			const parsed = parseStableLocation(cellId);
			if (!parsed || parsed.kind !== 'cell') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid cellId: ${cellId}.`, { field: 'cellId' });
			}
			let partXml = getPartXmlForLocation(context.session, parsed);
			const cellXml = getTableCellXmlFromPart(partXml, parsed);
			const nextCellXml = patchCellText(cellXml, text);
			partXml = replaceTableCellXmlInPart(partXml, parsed, nextCellXml);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(cellId);
			preview.push({ id: cellId, field: 'text', before: null, after: text });
			break;
		}
		case 'docx.setCellStyle': {
			const cellId = requireString(record.cellId, 'cellId');
			const style = record.style;
			if (!style || typeof style !== 'object') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'style must be an object.', { field: 'style' });
			}
			rejectWriteOnlyExcludedId(cellId, 'cellId');
			const parsed = parseStableLocation(cellId);
			if (!parsed || parsed.kind !== 'cell') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid cellId: ${cellId}.`, { field: 'cellId' });
			}
			let partXml = getPartXmlForLocation(context.session, parsed);
			const cellXml = getTableCellXmlFromPart(partXml, parsed);
			const nextCellXml = patchCellStyle(cellXml, style as Record<string, unknown>);
			partXml = replaceTableCellXmlInPart(partXml, parsed, nextCellXml);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(cellId);
			preview.push({ id: cellId, field: 'style', before: null, after: style });
			break;
		}
		case 'docx.insertImage': {
			const afterBlockId = requireString(record.afterBlockId, 'afterBlockId');
			const vaultImagePath = requireString(record.vaultImagePath, 'vaultImagePath');
			rejectWriteOnlyExcludedId(afterBlockId, 'afterBlockId');
			const anchor = parseStableLocation(afterBlockId);
			if (!anchor || anchor.part !== 'body') {
				throw createAiError(
					AI_ERROR_CODES.VALIDATION_FAILED,
					'insertImage is only supported in word/document.xml body blocks.',
					{ field: 'afterBlockId' },
				);
			}
			const image = await readVaultBinaryFile(context.vault, vaultImagePath);
			void getImageMimeType(image.extension);
			documentXml = await addInlineImage(
				context.session.getZip(),
				documentXml,
				afterBlockId,
				image.bytes,
				image.extension,
			);
			changedIds.push(afterBlockId);
			preview.push({ id: afterBlockId, field: 'insertImage', before: null, after: vaultImagePath });
			break;
		}
		case 'docx.replaceImage': {
			const blockId = requireString(record.blockId, 'blockId');
			const vaultImagePath = requireString(record.vaultImagePath, 'vaultImagePath');
			rejectWriteOnlyExcludedId(blockId, 'blockId');
			const parsed = parseStableLocation(blockId);
			if (!parsed || parsed.part !== 'body') {
				throw createAiError(
					AI_ERROR_CODES.VALIDATION_FAILED,
					'replaceImage is only supported in word/document.xml body blocks.',
					{ field: 'blockId' },
				);
			}
			const image = await readVaultBinaryFile(context.vault, vaultImagePath);
			documentXml = await replaceInlineImage(
				context.session.getZip(),
				documentXml,
				blockId,
				image.bytes,
				image.extension,
			);
			changedIds.push(blockId);
			preview.push({ id: blockId, field: 'replaceImage', before: null, after: vaultImagePath });
			break;
		}
		case 'docx.replaceText': {
			const query = requireString(record.query, 'query');
			const replacement = requireString(record.replacement, 'replacement');
			const matchCase = record.matchCase === true;
			const wholeWord = record.wholeWord === true;
			let replacementCount = 0;
			for (const partPath of listReplaceTextPartPaths(context.session.getZip())) {
				if (!context.session.hasPart(partPath)) continue;
				const result = replacePartText(
					context.session.getPartXml(partPath),
					query,
					replacement,
					{ matchCase, wholeWord },
				);
				if (result.replacementCount > 0) {
					context.session.setPartXml(partPath, result.partXml);
					replacementCount += result.replacementCount;
					if (partPath === resolvePartPath({ part: 'body', partNumber: null })) {
						documentXml = result.partXml;
					}
				}
			}
			if (replacementCount === 0) {
				warnings.push(`No matches found for query "${query}".`);
			} else {
				changedIds.push('document');
				preview.push({
					id: 'document',
					field: 'replaceText',
					before: query,
					after: replacement,
				});
			}
			break;
		}
		default:
			throw createAiError(AI_ERROR_CODES.UNKNOWN_OP, `Unknown DOCX operation: ${opId}.`, { op: opId });
	}

	context.session.setDocumentXml(documentXml);

	return { changedIds, preview, warnings, documentXml };
}
