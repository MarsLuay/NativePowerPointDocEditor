import type { Vault } from 'obsidian';
import { getImageMimeType } from '../PowerPointInsertModals';
import {
	deleteTableInPart,
	getParagraphXml,
	getTableCellXmlFromPart,
	insertBlockAfterInPart,
	replaceParagraphXml,
	replaceTableCellXmlInPart,
} from './docxBlockResolver';
import { DOCX_CORE_PROPERTIES_PATH, listReplaceTextPartPaths, resolvePartPath } from './docxParts';
import { removeAllDocxComments } from './docxComments';
import { patchDocxCoreProperties } from './docxCoreProperties';
import { addInlineImage, replaceInlineImage } from './docxMedia';
import { parseParagraph } from './docxOoxml';
import type { DocxPatchSession } from './docxPatchSession';
import {
	buildEmptyTableXml,
	patchCellStyle,
	patchCellText,
	patchParagraphBottomBorder,
	patchParagraphDefaultRunStyle,
	patchParagraphStyle,
	patchRunStyle,
	patchRunText,
	replacePartText,
	type DocxRunStylePatch,
} from './docxOoxmlWrite';
import {
	patchDocumentSectionLayout,
	patchParagraphLayout,
	type DocxParagraphAlignment,
	type DocxParagraphLayout,
	type DocxSectionLayoutPatch,
	type DocxTabAlignment,
} from './docxLayout';
import { applyReplaceBodyParagraphs } from './docxBodyParagraphs';
import { registerExternalHyperlink } from './docxHyperlink';
import {
	applyDeleteParagraphInPart,
	applyDeleteRangeInPart,
	applyInsertHyperlinkInPart,
	applyInsertParagraphBreakInPart,
	applyInsertTextInPart,
	applyRemoveHyperlinkInPart,
	type DocxTextPosition,
	type DocxTextRange,
} from './docxParagraphEdit';
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
	return op;
}

function requireString(value: unknown, field: string): string {
	if (typeof value !== 'string') {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be a string.`, { field });
	}
	return value;
}

function requireNonEmptyString(value: unknown, field: string): string {
	const text = requireString(value, field).trim();
	if (!text) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must not be empty.`, { field });
	}
	return text;
}

function requireNumber(value: unknown, field: string): number {
	if (typeof value !== 'number' || !Number.isFinite(value)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be a number.`, { field });
	}
	return value;
}

function asRunStylePatch(value: unknown): DocxRunStylePatch {
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'style must be an object.', { field: 'style' });
	}
	const record = value as Record<string, unknown>;
	for (const field of ['bold', 'italic', 'underline'] as const) {
		if (record[field] !== undefined && typeof record[field] !== 'boolean') {
			throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `style.${field} must be boolean.`, { field: `style.${field}` });
		}
	}
	if (record.fontFamily !== undefined && (typeof record.fontFamily !== 'string' || !record.fontFamily.trim())) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'style.fontFamily must be a non-empty string.', { field: 'style.fontFamily' });
	}
	if (record.fontSizePt !== undefined && (
		typeof record.fontSizePt !== 'number'
		|| !Number.isFinite(record.fontSizePt)
		|| record.fontSizePt <= 0
	)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'style.fontSizePt must be a number > 0.', { field: 'style.fontSizePt' });
	}
	if (record.color !== undefined && record.color !== null && typeof record.color !== 'string') {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'style.color must be a string or null.', { field: 'style.color' });
	}
	const style: DocxRunStylePatch = {
		...(typeof record.bold === 'boolean' ? { bold: record.bold } : {}),
		...(typeof record.italic === 'boolean' ? { italic: record.italic } : {}),
		...(typeof record.underline === 'boolean' ? { underline: record.underline } : {}),
		...(typeof record.fontFamily === 'string' ? { fontFamily: record.fontFamily } : {}),
		...(typeof record.fontSizePt === 'number' ? { fontSizePt: record.fontSizePt } : {}),
		...(typeof record.color === 'string' || record.color === null ? { color: record.color } : {}),
	};
	if (Object.keys(style).length === 0) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'style must include at least one supported property.', { field: 'style' });
	}
	return style;
}

function asParagraphBottomBorderPatch(value: unknown): {
	style: string;
	size?: number;
	space?: number;
	color?: string;
} {
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'border must be an object.', { field: 'border' });
	}
	const record = value as Record<string, unknown>;
	const style = requireNonEmptyString(record.style, 'border.style');
	const size = record.size === undefined ? undefined : requireInteger(record.size, 'border.size');
	const space = record.space === undefined ? undefined : requireInteger(record.space, 'border.space');
	const color = record.color === undefined ? undefined : requireNonEmptyString(record.color, 'border.color');
	return { style, ...(size !== undefined ? { size } : {}), ...(space !== undefined ? { space } : {}), ...(color ? { color } : {}) };
}

function asObject(value: unknown, field: string): Record<string, unknown> {
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be an object.`, { field });
	}
	return value as Record<string, unknown>;
}

function optionalInteger(value: unknown, field: string, allowNegative = false): number | undefined {
	if (value === undefined) return undefined;
	const numberValue = requireNumber(value, field);
	if (!Number.isInteger(numberValue) || (!allowNegative && numberValue < 0)) {
		throw createAiError(
			AI_ERROR_CODES.SCHEMA_INVALID,
			`${field} must be an integer${allowNegative ? '' : ' >= 0'}.`,
			{ field },
		);
	}
	return numberValue;
}

function asParagraphLayoutPatch(value: unknown): DocxParagraphLayout {
	const record = asObject(value, 'layout');
	const layout: DocxParagraphLayout = {};
	const alignments: DocxParagraphAlignment[] = ['left', 'center', 'right', 'both', 'distribute'];
	if (record.alignment !== undefined) {
		if (typeof record.alignment !== 'string' || !alignments.includes(record.alignment as DocxParagraphAlignment)) {
			throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'layout.alignment is invalid.', { field: 'layout.alignment' });
		}
		layout.alignment = record.alignment as DocxParagraphAlignment;
	}
	if (record.spacing !== undefined) {
		const spacingRecord = asObject(record.spacing, 'layout.spacing');
		const spacing: NonNullable<DocxParagraphLayout['spacing']> = {};
		for (const field of ['before', 'after', 'line'] as const) {
			const valueForField = optionalInteger(spacingRecord[field], `layout.spacing.${field}`);
			if (valueForField !== undefined) spacing[field] = valueForField;
		}
		if (spacingRecord.lineRule !== undefined) {
			const lineRule = spacingRecord.lineRule;
			if (lineRule !== 'auto' && lineRule !== 'exact' && lineRule !== 'atLeast') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'layout.spacing.lineRule is invalid.', { field: 'layout.spacing.lineRule' });
			}
			spacing.lineRule = lineRule;
		}
		layout.spacing = spacing;
	}
	if (record.indent !== undefined) {
		const indentRecord = asObject(record.indent, 'layout.indent');
		const indent: NonNullable<DocxParagraphLayout['indent']> = {};
		for (const field of ['left', 'right', 'firstLine', 'hanging'] as const) {
			const valueForField = optionalInteger(indentRecord[field], `layout.indent.${field}`, true);
			if (valueForField !== undefined) indent[field] = valueForField;
		}
		layout.indent = indent;
	}
	if (record.tabs !== undefined) {
		if (!Array.isArray(record.tabs)) {
			throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'layout.tabs must be an array.', { field: 'layout.tabs' });
		}
		const tabAlignments: DocxTabAlignment[] = ['left', 'center', 'right', 'decimal', 'bar', 'clear'];
		layout.tabs = record.tabs.map((entry, index) => {
			const tab = asObject(entry, `layout.tabs[${index}]`);
			const val = tab.val;
			if (typeof val !== 'string' || !tabAlignments.includes(val as DocxTabAlignment)) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `layout.tabs[${index}].val is invalid.`, { field: `layout.tabs[${index}].val` });
			}
			const pos = optionalInteger(tab.pos, `layout.tabs[${index}].pos`);
			if (pos === undefined) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `layout.tabs[${index}].pos is required.`, { field: `layout.tabs[${index}].pos` });
			}
			return { val: val as DocxTabAlignment, pos };
		});
	}
	for (const field of ['keepNext', 'keepLines', 'pageBreakBefore'] as const) {
		if (record[field] !== undefined) {
			if (typeof record[field] !== 'boolean') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `layout.${field} must be boolean.`, { field: `layout.${field}` });
			}
			layout[field] = record[field];
		}
	}
	return layout;
}

function asSectionLayoutPatch(value: unknown): DocxSectionLayoutPatch {
	const record = asObject(value, 'layout');
	const layout: DocxSectionLayoutPatch = {};
	if (record.pageSize !== undefined) {
		const pageSizeRecord = asObject(record.pageSize, 'layout.pageSize');
		const width = optionalInteger(pageSizeRecord.width, 'layout.pageSize.width');
		const height = optionalInteger(pageSizeRecord.height, 'layout.pageSize.height');
		if (width === undefined || width === 0 || height === undefined || height === 0) {
			throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'layout.pageSize.width and height are required and > 0.', { field: 'layout.pageSize' });
		}
		const orient = pageSizeRecord.orient;
		if (orient !== undefined && orient !== 'portrait' && orient !== 'landscape') {
			throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'layout.pageSize.orient is invalid.', { field: 'layout.pageSize.orient' });
		}
		layout.pageSize = { width, height, ...(orient ? { orient } : {}) };
	}
	if (record.margins !== undefined) {
		const marginsRecord = asObject(record.margins, 'layout.margins');
		const margins: NonNullable<DocxSectionLayoutPatch['margins']> = {};
		for (const field of ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'] as const) {
			const valueForField = optionalInteger(marginsRecord[field], `layout.margins.${field}`);
			if (valueForField !== undefined) margins[field] = valueForField;
		}
		layout.margins = margins;
	}
	return layout;
}

function requireInteger(value: unknown, field: string): number {
	const numberValue = requireNumber(value, field);
	if (!Number.isInteger(numberValue) || numberValue < 0) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be a non-negative integer.`, { field });
	}
	return numberValue;
}

function parseTextPosition(value: unknown, field: string): DocxTextPosition {
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be an object.`, { field });
	}
	const record = value as Record<string, unknown>;
	const blockId = requireString(record.blockId, `${field}.blockId`);
	const offset = requireInteger(record.offset, `${field}.offset`);
	const runId = typeof record.runId === 'string' ? record.runId : undefined;
	rejectWriteOnlyExcludedId(blockId, `${field}.blockId`);
	if (runId) {
		rejectWriteOnlyExcludedId(runId, `${field}.runId`);
	}
	return { blockId, offset, ...(runId ? { runId } : {}) };
}

function parseTextRange(value: unknown, field = 'range'): DocxTextRange {
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be an object.`, { field });
	}
	const record = value as Record<string, unknown>;
	return {
		start: parseTextPosition(record.start, `${field}.start`),
		end: parseTextPosition(record.end, `${field}.end`),
	};
}

function rejectWriteOnlyExcludedId(id: string, field: string): void {
	if (id.startsWith('comments/')) {
		throw createAiError(
			AI_ERROR_CODES.VALIDATION_FAILED,
			'Individual comments are describe-only. Use docx.removeComments to delete all comments; trackChanges markup is not writable via AI ops.',
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
		case 'docx.removeComments': {
			const removal = await removeAllDocxComments(context.session);
			documentXml = removal.documentXml;
			changedIds.push(...removal.changedPartPaths);
			if (removal.changedPartPaths.length > 0) {
				preview.push({
					id: 'comments',
					field: 'removeAll',
					before: removal.commentCount,
					after: 0,
				});
			}
			break;
		}
		case 'docx.setCoreProperties': {
			const creator = requireNonEmptyString(record.creator, 'creator');
			const lastModifiedBy = requireNonEmptyString(record.lastModifiedBy, 'lastModifiedBy');
			const existingCoreXml = context.session.hasPart(DOCX_CORE_PROPERTIES_PATH)
				? context.session.getPartXml(DOCX_CORE_PROPERTIES_PATH)
				: [
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
					'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"',
					' xmlns:dc="http://purl.org/dc/elements/1.1/"',
					' xmlns:dcterms="http://purl.org/dc/terms/"',
					' xmlns:dcmitype="http://purl.org/dc/dcmitype/"',
					' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
					'</cp:coreProperties>',
				].join('');
			const beforeCreator = /<dc:creator\b[^>]*>([^<]*)<\/dc:creator>/.exec(existingCoreXml)?.[1] ?? null;
			const beforeLastModifiedBy = /<cp:lastModifiedBy\b[^>]*>([^<]*)<\/cp:lastModifiedBy>/.exec(existingCoreXml)?.[1] ?? null;
			context.session.setPartXml(
				DOCX_CORE_PROPERTIES_PATH,
				patchDocxCoreProperties(existingCoreXml, { creator, lastModifiedBy }),
			);
			changedIds.push(DOCX_CORE_PROPERTIES_PATH);
			preview.push(
				{ id: DOCX_CORE_PROPERTIES_PATH, field: 'creator', before: beforeCreator, after: creator },
				{ id: DOCX_CORE_PROPERTIES_PATH, field: 'lastModifiedBy', before: beforeLastModifiedBy, after: lastModifiedBy },
			);
			break;
		}
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
			const nextParagraphXml = patchParagraphStyle(paragraphXml, style);
			partXml = replaceParagraphXml(partXml, parsed, nextParagraphXml);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(blockId);
			preview.push({ id: blockId, field: 'style', before: null, after: style });
			break;
		}
		case 'docx.setParagraphDefaultRunStyle': {
			const blockId = requireString(record.blockId, 'blockId');
			const style = asRunStylePatch(record.style);
			rejectWriteOnlyExcludedId(blockId, 'blockId');
			const parsed = parseStableLocation(blockId);
			if (!parsed || parsed.kind !== 'paragraph') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid blockId: ${blockId}.`, { field: 'blockId' });
			}
			let partXml = getPartXmlForLocation(context.session, parsed);
			const paragraphXml = getParagraphXml(partXml, parsed);
			const before = parseParagraph(paragraphXml).defaultRunStyle;
			const nextParagraphXml = patchParagraphDefaultRunStyle(paragraphXml, style);
			partXml = replaceParagraphXml(partXml, parsed, nextParagraphXml);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(blockId);
			preview.push({ id: blockId, field: 'defaultRunStyle', before, after: style });
			break;
		}
		case 'docx.setParagraphLayout': {
			const blockId = requireString(record.blockId, 'blockId');
			const layout = asParagraphLayoutPatch(record.layout);
			rejectWriteOnlyExcludedId(blockId, 'blockId');
			const parsed = parseStableLocation(blockId);
			if (!parsed || parsed.kind !== 'paragraph') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid blockId: ${blockId}.`, { field: 'blockId' });
			}
			let partXml = getPartXmlForLocation(context.session, parsed);
			const nextParagraphXml = patchParagraphLayout(getParagraphXml(partXml, parsed), layout);
			partXml = replaceParagraphXml(partXml, parsed, nextParagraphXml);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(blockId);
			preview.push({ id: blockId, field: 'layout', before: null, after: layout });
			break;
		}
		case 'docx.setSectionLayout': {
			const sectionIndex = requireInteger(record.sectionIndex, 'sectionIndex');
			const layout = asSectionLayoutPatch(record.layout);
			documentXml = patchDocumentSectionLayout(documentXml, sectionIndex, layout);
			changedIds.push(`body/sectPr[${sectionIndex}]`);
			preview.push({ id: `body/sectPr[${sectionIndex}]`, field: 'layout', before: null, after: layout });
			break;
		}
		case 'docx.setParagraphBottomBorder': {
			const blockId = requireString(record.blockId, 'blockId');
			const border = asParagraphBottomBorderPatch(record.border);
			rejectWriteOnlyExcludedId(blockId, 'blockId');
			const location = parseStableLocation(blockId);
			if (!location || location.kind !== 'paragraph') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid paragraph blockId: ${blockId}.`, { field: 'blockId' });
			}
			let partXml = getPartXmlForLocation(context.session, location);
			partXml = replaceParagraphXml(
				partXml,
				location,
				patchParagraphBottomBorder(getParagraphXml(partXml, location), border),
			);
			setPartXmlForLocation(context.session, location, partXml);
			if (location.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(blockId);
			preview.push({ id: blockId, field: 'bottomBorder', before: null, after: border });
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
		case 'docx.deleteTable': {
			const tableId = requireString(record.tableId, 'tableId');
			rejectWriteOnlyExcludedId(tableId, 'tableId');
			const parsed = parseStableLocation(tableId);
			if (!parsed || parsed.kind !== 'table') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid tableId: ${tableId}.`, { field: 'tableId' });
			}
			let partXml = getPartXmlForLocation(context.session, parsed);
			partXml = deleteTableInPart(partXml, parsed);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(tableId);
			preview.push({ id: tableId, field: 'deleteTable', before: 'table', after: null });
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
					'insertImage is only supported on body blocks in the main DOCX part.',
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
					'replaceImage is only supported on body blocks in the main DOCX part.',
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
		case 'docx.insertText': {
			const blockId = requireString(record.blockId, 'blockId');
			const offset = requireInteger(record.offset, 'offset');
			const text = requireString(record.text, 'text');
			const runId = typeof record.runId === 'string' ? record.runId : undefined;
			rejectWriteOnlyExcludedId(blockId, 'blockId');
			if (runId) rejectWriteOnlyExcludedId(runId, 'runId');
			const location = parseStableLocation(blockId);
			if (!location || location.kind !== 'paragraph') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid blockId: ${blockId}.`, { field: 'blockId' });
			}
			let partXml = getPartXmlForLocation(context.session, location);
			partXml = applyInsertTextInPart(partXml, { blockId, offset, ...(runId ? { runId } : {}) }, text);
			setPartXmlForLocation(context.session, location, partXml);
			if (location.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(blockId);
			preview.push({ id: blockId, field: 'insertText', before: null, after: { offset, text } });
			break;
		}
		case 'docx.deleteRange': {
			const range = parseTextRange(record.range, 'range');
			const startLocation = parseStableLocation(range.start.blockId);
			if (!startLocation) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid blockId: ${range.start.blockId}.`, { field: 'range.start.blockId' });
			}
			let partXml = getPartXmlForLocation(context.session, startLocation);
			partXml = applyDeleteRangeInPart(partXml, range);
			setPartXmlForLocation(context.session, startLocation, partXml);
			if (startLocation.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(range.start.blockId, range.end.blockId);
			preview.push({ id: range.start.blockId, field: 'deleteRange', before: range, after: null });
			break;
		}
		case 'docx.deleteBlock': {
			const blockId = requireString(record.blockId, 'blockId');
			rejectWriteOnlyExcludedId(blockId, 'blockId');
			const location = parseStableLocation(blockId);
			if (!location || location.kind !== 'paragraph') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid paragraph blockId: ${blockId}.`, { field: 'blockId' });
			}
			let partXml = getPartXmlForLocation(context.session, location);
			partXml = applyDeleteParagraphInPart(partXml, blockId);
			setPartXmlForLocation(context.session, location, partXml);
			if (location.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(blockId);
			preview.push({ id: blockId, field: 'deleteBlock', before: 'paragraph', after: null });
			break;
		}
		case 'docx.insertHyperlink': {
			const range = parseTextRange(record.range, 'range');
			const url = requireString(record.url, 'url');
			const displayText = typeof record.displayText === 'string' ? record.displayText : undefined;
			const tooltip = typeof record.tooltip === 'string' ? record.tooltip : undefined;
			const startLocation = parseStableLocation(range.start.blockId);
			if (!startLocation) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid blockId: ${range.start.blockId}.`, { field: 'range.start.blockId' });
			}
			const relationshipId = await registerExternalHyperlink(context.session.getZip(), startLocation, url);
			let partXml = getPartXmlForLocation(context.session, startLocation);
			partXml = applyInsertHyperlinkInPart(partXml, range, relationshipId, displayText, tooltip);
			setPartXmlForLocation(context.session, startLocation, partXml);
			if (startLocation.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(range.start.blockId);
			preview.push({ id: range.start.blockId, field: 'insertHyperlink', before: null, after: { url, relationshipId } });
			break;
		}
		case 'docx.removeHyperlink': {
			const range = parseTextRange(record.range, 'range');
			const startLocation = parseStableLocation(range.start.blockId);
			if (!startLocation) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid blockId: ${range.start.blockId}.`, { field: 'range.start.blockId' });
			}
			let partXml = getPartXmlForLocation(context.session, startLocation);
			partXml = applyRemoveHyperlinkInPart(partXml, range);
			setPartXmlForLocation(context.session, startLocation, partXml);
			if (startLocation.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(range.start.blockId);
			preview.push({ id: range.start.blockId, field: 'removeHyperlink', before: range, after: null });
			break;
		}
		case 'docx.insertParagraphBreak': {
			const blockId = requireString(record.blockId, 'blockId');
			const offset = requireInteger(record.offset, 'offset');
			const runId = typeof record.runId === 'string' ? record.runId : undefined;
			rejectWriteOnlyExcludedId(blockId, 'blockId');
			if (runId) rejectWriteOnlyExcludedId(runId, 'runId');
			const location = parseStableLocation(blockId);
			if (!location || location.kind !== 'paragraph') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid blockId: ${blockId}.`, { field: 'blockId' });
			}
			let partXml = getPartXmlForLocation(context.session, location);
			partXml = applyInsertParagraphBreakInPart(partXml, { blockId, offset, ...(runId ? { runId } : {}) });
			setPartXmlForLocation(context.session, location, partXml);
			if (location.part === 'body') {
				documentXml = partXml;
			}
			changedIds.push(blockId);
			preview.push({ id: blockId, field: 'insertParagraphBreak', before: null, after: { offset } });
			break;
		}
		case 'docx.replaceBodyParagraphs': {
			const paragraphsValue = record.paragraphs;
			if (!Array.isArray(paragraphsValue)) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'paragraphs must be an array of strings.', {
					field: 'paragraphs',
				});
			}
			const paragraphs = paragraphsValue.map((entry, index) => {
				if (typeof entry !== 'string') {
					throw createAiError(
						AI_ERROR_CODES.SCHEMA_INVALID,
						`paragraphs[${index}] must be a string.`,
						{ field: 'paragraphs' },
					);
				}
				return entry;
			});
			documentXml = applyReplaceBodyParagraphs(documentXml, paragraphs);
			changedIds.push('body');
			preview.push({
				id: 'body',
				field: 'replaceBodyParagraphs',
				before: null,
				after: { paragraphCount: paragraphs.length > 0 ? paragraphs.length : 1 },
			});
			break;
		}
		default:
			throw createAiError(AI_ERROR_CODES.UNKNOWN_OP, `Unknown DOCX operation: ${opId}.`, { op: opId });
	}

	context.session.setDocumentXml(documentXml);

	return { changedIds, preview, warnings, documentXml };
}
