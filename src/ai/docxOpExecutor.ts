import type { Vault } from 'obsidian';
import { getImageMimeType } from '../PowerPointInsertModals';
import {
	deleteTableInPart,
	findParagraphByAnchorInPart,
	getParagraphXml,
	resolveParagraphReferenceInPart,
	getTableCellXmlFromPart,
	insertBlockAfterInPart,
	replaceParagraphXml,
	replaceTableCellXmlInPart,
} from './docxBlockResolver';
import { DOCX_CORE_PROPERTIES_PATH, listDocxDescribeParts, listReplaceTextPartPaths, resolvePartPath } from './docxParts';
import { removeAllDocxComments } from './docxComments';
import { patchDocxCoreProperties } from './docxCoreProperties';
import { addInlineImage, replaceInlineImage } from './docxMedia';
import { ensureParagraphAnchors, parseParagraph } from './docxOoxml';
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
	getRunText,
	replacePartText,
	type DocxParagraphBottomBorderPatch,
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
	applyInsertParagraphsAfterInPart,
	applyInsertParagraphsInPart,
	applyInsertParagraphBreakInPart,
	FULL_PARAGRAPH_INHERITANCE,
	type DocxParagraphInheritance,
	applyInsertTextInPart,
	applyRemoveHyperlinkInPart,
	type DocxTextPosition,
	type DocxTextRange,
} from './docxParagraphEdit';
import { parseStableLocation } from './docxStableIds';
import { AI_ERROR_CODES, createAiError, isAiErrorDetail } from './errors';
import type { ApplyPreviewChange, ApplyResult, DocumentOp } from './types';
import { readVaultBinaryFile } from './vaultBinary';

export interface DocxOpExecutionResult {
	changedIds: string[];
	createdIds: string[];
	createdAnchors: string[];
	structuralMutations: NonNullable<ApplyResult['structuralMutations']>;
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

interface DocxOpAccumulator {
	documentXml: string;
	changedIds: string[];
	createdIds: string[];
	createdAnchors: string[];
	structuralMutations: NonNullable<ApplyResult['structuralMutations']>;
	preview: ApplyPreviewChange[];
	warnings: string[];
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
	const anchor = typeof record.anchor === 'string' ? record.anchor : undefined;
	rejectWriteOnlyExcludedId(blockId, `${field}.blockId`);
	if (runId) {
		rejectWriteOnlyExcludedId(runId, `${field}.runId`);
	}
	return { blockId, offset, ...(runId ? { runId } : {}), ...(anchor ? { anchor } : {}) };
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

function remapRunId(blockId: string, runId: string): string {
	const parsed = parseStableLocation(runId);
	if (!parsed || parsed.kind !== 'run') {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid runId: ${runId}.`, { field: 'runId' });
	}
	return `${blockId}/r[${parsed.runIndex ?? 0}]`;
}

function resolveParagraphTarget(
	session: DocxPatchSession,
	blockId: string,
	anchor: unknown,
	field: string,
): { location: NonNullable<ReturnType<typeof parseStableLocation>>; blockId: string; partXml: string; anchor: string } {
	const requested = parseStableLocation(blockId);
	if (!requested || requested.kind !== 'paragraph') {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid ${field}: ${blockId}.`, { field });
	}
	const partXml = getPartXmlForLocation(session, requested);
	const resolved = resolveParagraphReferenceInPart(
		partXml,
		blockId,
		typeof anchor === 'string' ? anchor : undefined,
	);
	return { location: resolved.location, blockId: resolved.blockId, partXml, anchor: resolved.anchor };
}

function resolveTextTarget(
	session: DocxPatchSession,
	position: DocxTextPosition,
	field: string,
): { position: DocxTextPosition; location: NonNullable<ReturnType<typeof parseStableLocation>>; partXml: string } {
	const resolved = resolveParagraphTarget(session, position.blockId, position.anchor, field);
	const runId = position.runId ? remapRunId(resolved.blockId, position.runId) : undefined;
	return {
		position: {
			...position,
			blockId: resolved.blockId,
			...(runId ? { runId } : {}),
			...(resolved.anchor ? { anchor: resolved.anchor } : {}),
		},
		location: resolved.location,
		partXml: resolved.partXml,
	};
}

function findParagraphAnchor(
	session: DocxPatchSession,
	anchor: string,
	field: string,
): { location: NonNullable<ReturnType<typeof parseStableLocation>>; blockId: string; partXml: string; paragraphXml: string } {
	const matches = [];
	for (const part of listDocxDescribeParts(session.getZip())) {
		if (!session.hasPart(part.path)) continue;
		const partXml = session.getPartXml(part.path);
		try {
			const block = findParagraphByAnchorInPart(partXml, part, anchor);
			const location = parseStableLocation(block.id);
			if (!location || location.kind !== 'paragraph') continue;
			matches.push({ location, blockId: block.id, partXml, paragraphXml: block.xml });
		} catch (error) {
			if (!isAiErrorDetail(error) || error.code !== AI_ERROR_CODES.BLOCK_NOT_FOUND) {
				throw error;
			}
		}
	}
	if (matches.length !== 1) {
		throw createAiError(
			AI_ERROR_CODES.BLOCK_NOT_FOUND,
			matches.length === 0
				? `Paragraph anchor ${anchor} was not found.`
				: `Paragraph anchor ${anchor} is not unique.`,
			{ field },
		);
	}
	return matches[0]!;
}

function parseParagraphInheritance(value: unknown): DocxParagraphInheritance {
	if (value === undefined) return { ...FULL_PARAGRAPH_INHERITANCE };
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'inherit must be an object.', { field: 'inherit' });
	}
	const record = value as Record<string, unknown>;
	const inheritance: DocxParagraphInheritance = {
		paragraph: false,
		run: false,
		layout: false,
		border: false,
		list: false,
	};
	for (const key of ['paragraph', 'run', 'layout', 'border', 'list'] as const) {
		if (record[key] === undefined) continue;
		if (typeof record[key] !== 'boolean') {
			throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `inherit.${key} must be boolean.`, { field: `inherit.${key}` });
		}
		inheritance[key] = record[key];
	}
	return inheritance;
}

async function executeCommentsAndMetadataOp(
	context: DocxOpExecutionContext,
	opId: string,
	record: Record<string, unknown>,
	acc: DocxOpAccumulator,
): Promise<void> {
	switch (opId) {
		case 'docx.removeComments': {
			const removal = await removeAllDocxComments(context.session);
			acc.documentXml = removal.documentXml;
			acc.changedIds.push(...removal.changedPartPaths);
			if (removal.changedPartPaths.length > 0) {
				acc.preview.push({
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
			acc.changedIds.push(DOCX_CORE_PROPERTIES_PATH);
			acc.preview.push(
				{ id: DOCX_CORE_PROPERTIES_PATH, field: 'creator', before: beforeCreator, after: creator },
				{ id: DOCX_CORE_PROPERTIES_PATH, field: 'lastModifiedBy', before: beforeLastModifiedBy, after: lastModifiedBy },
			);
			break;
		}
	}
}

function executeFormattingOp(
	context: DocxOpExecutionContext,
	opId: string,
	record: Record<string, unknown>,
	acc: DocxOpAccumulator,
): void {
	switch (opId) {
		case 'docx.setRunStyle': {
			const requestedRunId = requireString(record.runId, 'runId');
			const style = asRunStylePatch(record.style);
			rejectWriteOnlyExcludedId(requestedRunId, 'runId');
			const resolved = resolveParagraphTarget(
				context.session,
				requestedRunId.replace(/\/r\[\d+\]$/, ''),
				record.anchor,
				'runId',
			);
			const runId = remapRunId(resolved.blockId, requestedRunId);
			const parsedRun = parseStableLocation(runId);
			if (!parsedRun || parsedRun.kind !== 'run') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid runId: ${runId}.`, { field: 'runId' });
			}
			let partXml = resolved.partXml;
			const paragraphXml = getParagraphXml(partXml, parsedRun);
			const nextParagraphXml = patchRunStyle(paragraphXml, parsedRun.runIndex ?? 0, style);
			partXml = replaceParagraphXml(partXml, parsedRun, nextParagraphXml);
			setPartXmlForLocation(context.session, parsedRun, partXml);
			if (parsedRun.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(runId);
			acc.preview.push({ id: runId, field: 'style', before: null, after: style });
			break;
		}
		case 'docx.setParagraphStyle': {
			const requestedBlockId = requireString(record.blockId, 'blockId');
			const style = record.style;
			if (!style || typeof style !== 'object') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'style must be an object.', { field: 'style' });
			}
			rejectWriteOnlyExcludedId(requestedBlockId, 'blockId');
			const resolved = resolveParagraphTarget(context.session, requestedBlockId, record.anchor, 'blockId');
			const blockId = resolved.blockId;
			const parsed = resolved.location;
			let partXml = resolved.partXml;
			const paragraphXml = getParagraphXml(partXml, parsed);
			const nextParagraphXml = patchParagraphStyle(paragraphXml, style);
			partXml = replaceParagraphXml(partXml, parsed, nextParagraphXml);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(blockId);
			acc.preview.push({ id: blockId, field: 'style', before: null, after: style });
			break;
		}
		case 'docx.setParagraphDefaultRunStyle': {
			const requestedBlockId = requireString(record.blockId, 'blockId');
			const style = asRunStylePatch(record.style);
			rejectWriteOnlyExcludedId(requestedBlockId, 'blockId');
			const resolved = resolveParagraphTarget(context.session, requestedBlockId, record.anchor, 'blockId');
			const blockId = resolved.blockId;
			const parsed = resolved.location;
			let partXml = resolved.partXml;
			const paragraphXml = getParagraphXml(partXml, parsed);
			const before = parseParagraph(paragraphXml).defaultRunStyle;
			const nextParagraphXml = patchParagraphDefaultRunStyle(paragraphXml, style);
			partXml = replaceParagraphXml(partXml, parsed, nextParagraphXml);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(blockId);
			acc.preview.push({ id: blockId, field: 'defaultRunStyle', before, after: style });
			break;
		}
		case 'docx.setParagraphLayout': {
			const requestedBlockId = requireString(record.blockId, 'blockId');
			const layout = asParagraphLayoutPatch(record.layout);
			rejectWriteOnlyExcludedId(requestedBlockId, 'blockId');
			const resolved = resolveParagraphTarget(context.session, requestedBlockId, record.anchor, 'blockId');
			const blockId = resolved.blockId;
			const parsed = resolved.location;
			let partXml = resolved.partXml;
			const nextParagraphXml = patchParagraphLayout(getParagraphXml(partXml, parsed), layout);
			partXml = replaceParagraphXml(partXml, parsed, nextParagraphXml);
			setPartXmlForLocation(context.session, parsed, partXml);
			if (parsed.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(blockId);
			acc.preview.push({ id: blockId, field: 'layout', before: null, after: layout });
			break;
		}
		case 'docx.setSectionLayout': {
			const sectionIndex = requireInteger(record.sectionIndex, 'sectionIndex');
			const layout = asSectionLayoutPatch(record.layout);
			acc.documentXml = patchDocumentSectionLayout(acc.documentXml, sectionIndex, layout);
			acc.changedIds.push(`body/sectPr[${sectionIndex}]`);
			acc.preview.push({ id: `body/sectPr[${sectionIndex}]`, field: 'layout', before: null, after: layout });
			break;
		}
		case 'docx.setParagraphBottomBorder': {
			const requestedBlockId = requireString(record.blockId, 'blockId');
			const border = asParagraphBottomBorderPatch(record.border);
			rejectWriteOnlyExcludedId(requestedBlockId, 'blockId');
			const resolved = resolveParagraphTarget(context.session, requestedBlockId, record.anchor, 'blockId');
			const blockId = resolved.blockId;
			const location = resolved.location;
			let partXml = resolved.partXml;
			partXml = replaceParagraphXml(
				partXml,
				location,
				patchParagraphBottomBorder(getParagraphXml(partXml, location), border),
			);
			setPartXmlForLocation(context.session, location, partXml);
			if (location.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(blockId);
			acc.preview.push({ id: blockId, field: 'bottomBorder', before: null, after: border });
			break;
		}
	}
}

function executeTableOp(
	context: DocxOpExecutionContext,
	opId: string,
	record: Record<string, unknown>,
	acc: DocxOpAccumulator,
): void {
	switch (opId) {
		case 'docx.insertTable': {
			const afterBlockId = requireString(record.afterBlockId, 'afterBlockId');
			const rows = requireNumber(record.rows, 'rows');
			const cols = requireNumber(record.cols, 'cols');
			if (!Number.isInteger(rows) || rows < 1 || !Number.isInteger(cols) || cols < 1) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'rows and cols must be positive integers.', { field: 'rows' });
			}
			rejectWriteOnlyExcludedId(afterBlockId, 'afterBlockId');
			const resolved = resolveParagraphTarget(context.session, afterBlockId, record.anchor, 'afterBlockId');
			const anchor = resolved.location;
			const tableXml = buildEmptyTableXml(rows, cols);
			let partXml = resolved.partXml;
			partXml = insertBlockAfterInPart(partXml, resolved.blockId, tableXml);
			setPartXmlForLocation(context.session, anchor, partXml);
			if (anchor.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(resolved.blockId);
			acc.preview.push({ id: resolved.anchor || resolved.blockId, field: 'insertTable', before: null, after: { rows, cols } });
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
				acc.documentXml = partXml;
			}
			acc.changedIds.push(cellId);
			acc.preview.push({ id: cellId, field: 'text', before: null, after: text });
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
				acc.documentXml = partXml;
			}
			acc.changedIds.push(cellId);
			acc.preview.push({ id: cellId, field: 'style', before: null, after: style });
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
				acc.documentXml = partXml;
			}
			acc.changedIds.push(tableId);
			acc.preview.push({ id: tableId, field: 'deleteTable', before: 'table', after: null });
			break;
		}
	}
}

async function executeMediaOp(
	context: DocxOpExecutionContext,
	opId: string,
	record: Record<string, unknown>,
	acc: DocxOpAccumulator,
): Promise<void> {
	switch (opId) {
		case 'docx.insertImage': {
			const afterBlockId = requireString(record.afterBlockId, 'afterBlockId');
			const vaultImagePath = requireString(record.vaultImagePath, 'vaultImagePath');
			rejectWriteOnlyExcludedId(afterBlockId, 'afterBlockId');
			const resolved = resolveParagraphTarget(context.session, afterBlockId, record.anchor, 'afterBlockId');
			if (resolved.location.part !== 'body') {
				throw createAiError(
					AI_ERROR_CODES.VALIDATION_FAILED,
					'insertImage is only supported on body blocks in the main DOCX part.',
					{ field: 'afterBlockId' },
				);
			}
			const image = await readVaultBinaryFile(context.vault, vaultImagePath);
			void getImageMimeType(image.extension);
			acc.documentXml = await addInlineImage(
				context.session.getZip(),
				acc.documentXml,
				resolved.blockId,
				image.bytes,
				image.extension,
			);
			acc.changedIds.push(resolved.blockId);
			acc.preview.push({ id: resolved.anchor || resolved.blockId, field: 'insertImage', before: null, after: vaultImagePath });
			break;
		}
		case 'docx.replaceImage': {
			const requestedBlockId = requireString(record.blockId, 'blockId');
			const vaultImagePath = requireString(record.vaultImagePath, 'vaultImagePath');
			rejectWriteOnlyExcludedId(requestedBlockId, 'blockId');
			const resolved = resolveParagraphTarget(context.session, requestedBlockId, record.anchor, 'blockId');
			const blockId = resolved.blockId;
			if (resolved.location.part !== 'body') {
				throw createAiError(
					AI_ERROR_CODES.VALIDATION_FAILED,
					'replaceImage is only supported on body blocks in the main DOCX part.',
					{ field: 'blockId' },
				);
			}
			const image = await readVaultBinaryFile(context.vault, vaultImagePath);
			acc.documentXml = await replaceInlineImage(
				context.session.getZip(),
				acc.documentXml,
				blockId,
				image.bytes,
				image.extension,
			);
			acc.changedIds.push(blockId);
			acc.preview.push({ id: blockId, field: 'replaceImage', before: null, after: vaultImagePath });
			break;
		}
	}
}

async function executeHyperlinkOp(
	context: DocxOpExecutionContext,
	opId: string,
	record: Record<string, unknown>,
	acc: DocxOpAccumulator,
): Promise<void> {
	switch (opId) {
		case 'docx.insertHyperlink': {
			const parsedRange = parseTextRange(record.range, 'range');
			const start = resolveTextTarget(context.session, parsedRange.start, 'range.start.blockId');
			const end = resolveTextTarget(context.session, parsedRange.end, 'range.end.blockId');
			const range = { start: start.position, end: end.position };
			const url = requireString(record.url, 'url');
			const displayText = typeof record.displayText === 'string' ? record.displayText : undefined;
			const tooltip = typeof record.tooltip === 'string' ? record.tooltip : undefined;
			const startLocation = start.location;
			const relationshipId = await registerExternalHyperlink(context.session.getZip(), startLocation, url);
			let partXml = start.partXml;
			partXml = applyInsertHyperlinkInPart(partXml, range, relationshipId, displayText, tooltip);
			setPartXmlForLocation(context.session, startLocation, partXml);
			if (startLocation.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(range.start.blockId);
			acc.preview.push({ id: range.start.blockId, field: 'insertHyperlink', before: null, after: { url, relationshipId } });
			break;
		}
		case 'docx.removeHyperlink': {
			const parsedRange = parseTextRange(record.range, 'range');
			const start = resolveTextTarget(context.session, parsedRange.start, 'range.start.blockId');
			const end = resolveTextTarget(context.session, parsedRange.end, 'range.end.blockId');
			const range = { start: start.position, end: end.position };
			const startLocation = start.location;
			let partXml = start.partXml;
			partXml = applyRemoveHyperlinkInPart(partXml, range);
			setPartXmlForLocation(context.session, startLocation, partXml);
			if (startLocation.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(range.start.blockId);
			acc.preview.push({ id: range.start.blockId, field: 'removeHyperlink', before: range, after: null });
			break;
		}
	}
}

function executeTextEditOp(
	context: DocxOpExecutionContext,
	opId: string,
	record: Record<string, unknown>,
	acc: DocxOpAccumulator,
): void {
	switch (opId) {
		case 'docx.setRunText': {
			const requestedBlockId = requireString(record.blockId, 'blockId');
			const requestedRunId = requireString(record.runId, 'runId');
			const text = requireString(record.text, 'text');
			rejectWriteOnlyExcludedId(requestedBlockId, 'blockId');
			rejectWriteOnlyExcludedId(requestedRunId, 'runId');
			const resolved = resolveParagraphTarget(context.session, requestedBlockId, record.anchor, 'blockId');
			const blockId = resolved.blockId;
			const runId = remapRunId(blockId, requestedRunId);
			const parsedRun = parseStableLocation(runId);
			if (!parsedRun || parsedRun.kind !== 'run') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid runId: ${runId}.`, { field: 'runId' });
			}
			let partXml = resolved.partXml;
			const paragraphXml = getParagraphXml(partXml, parsedRun);
			if (getRunText(paragraphXml, parsedRun.runIndex ?? 0).length === 0) {
				throw createAiError(
					AI_ERROR_CODES.EMPTY_RUN_USE_INSERT_TEXT,
					`Run ${runId} is empty. Use docx.insertText with blockId ${blockId} and offset 0 to populate an empty paragraph.`,
					{ op: opId, field: 'runId' },
				);
			}
			const nextParagraphXml = patchRunText(paragraphXml, parsedRun.runIndex ?? 0, text);
			partXml = replaceParagraphXml(partXml, parsedRun, nextParagraphXml);
			setPartXmlForLocation(context.session, parsedRun, partXml);
			if (parsedRun.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(runId);
			acc.preview.push({ id: runId, field: 'text', before: null, after: text });
			break;
		}
		case 'docx.insertParagraphsAfter': {
			const afterBlockId = requireString(record.afterBlockId, 'afterBlockId');
			const paragraphsValue = record.paragraphs;
			if (!Array.isArray(paragraphsValue)) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'paragraphs must be an array.', { field: 'paragraphs' });
			}
			const paragraphs = paragraphsValue.map((entry, index) => {
				if (!entry || typeof entry !== 'object' || Array.isArray(entry)) {
					throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `paragraphs[${index}] must be an object.`, { field: 'paragraphs' });
				}
				const paragraph = entry as Record<string, unknown>;
				if (typeof paragraph.text !== 'string') {
					throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `paragraphs[${index}].text must be a string.`, { field: 'paragraphs' });
				}
				return {
					text: paragraph.text,
					...(typeof paragraph.listStyle === 'string' ? { listStyle: paragraph.listStyle as 'none' | 'bullet' | 'number' } : {}),
					...(typeof paragraph.bold === 'boolean' ? { bold: paragraph.bold } : {}),
					...(paragraph.runStyle && typeof paragraph.runStyle === 'object' && !Array.isArray(paragraph.runStyle) ? { runStyle: paragraph.runStyle } : {}),
					...(paragraph.layout && typeof paragraph.layout === 'object' && !Array.isArray(paragraph.layout) ? { layout: paragraph.layout } : {}),
					...(paragraph.border && typeof paragraph.border === 'object' && !Array.isArray(paragraph.border) ? { border: paragraph.border as DocxParagraphBottomBorderPatch } : {}),
					...(Number.isInteger(paragraph.listLevel) ? { listLevel: paragraph.listLevel as number } : {}),
					...(Number.isInteger(paragraph.numId) ? { numId: paragraph.numId as number } : {}),
				};
			});
			rejectWriteOnlyExcludedId(afterBlockId, 'afterBlockId');
			const resolved = resolveParagraphTarget(context.session, afterBlockId, record.anchor, 'afterBlockId');
			const anchor = resolved.location;
			let partXml = resolved.partXml;
			const result = applyInsertParagraphsAfterInPart(partXml, resolved.blockId, paragraphs);
			partXml = result.partXml;
			setPartXmlForLocation(context.session, anchor, partXml);
			if (anchor.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(afterBlockId, ...result.createdBlockIds);
			acc.createdIds.push(...result.createdBlockIds);
			acc.createdAnchors.push(...(result.createdAnchors ?? []));
			acc.preview.push({
				id: afterBlockId,
				field: 'insertParagraphsAfter',
				before: null,
				after: {
					createdBlockIds: result.createdBlockIds,
					paragraphCount: paragraphs.length,
					inheritedListProperties: result.inheritedListProperties,
				},
			});
			break;
		}
		case 'docx.insertParagraphs': {
			const placement = record.placement;
			if (placement !== 'before' && placement !== 'after') {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'placement must be before or after.', { field: 'placement' });
			}
			const anchor = requireString(record.anchor, 'anchor');
			const paragraphsValue = record.paragraphs;
			if (!Array.isArray(paragraphsValue) || paragraphsValue.length === 0) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'paragraphs must be a non-empty array.', { field: 'paragraphs' });
			}
			const paragraphs = paragraphsValue.map((entry, index) => {
				if (!entry || typeof entry !== 'object' || Array.isArray(entry)) {
					throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `paragraphs[${index}] must be an object.`, { field: 'paragraphs' });
				}
				const paragraph = entry as Record<string, unknown>;
				if (typeof paragraph.text !== 'string') {
					throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `paragraphs[${index}].text must be a string.`, { field: 'paragraphs' });
				}
				return {
					text: paragraph.text,
					...(typeof paragraph.listStyle === 'string' ? { listStyle: paragraph.listStyle as 'none' | 'bullet' | 'number' } : {}),
					...(typeof paragraph.bold === 'boolean' ? { bold: paragraph.bold } : {}),
					...(paragraph.runStyle && typeof paragraph.runStyle === 'object' && !Array.isArray(paragraph.runStyle) ? { runStyle: paragraph.runStyle } : {}),
					...(paragraph.layout && typeof paragraph.layout === 'object' && !Array.isArray(paragraph.layout) ? { layout: paragraph.layout } : {}),
					...(paragraph.border && typeof paragraph.border === 'object' && !Array.isArray(paragraph.border) ? { border: paragraph.border as DocxParagraphBottomBorderPatch } : {}),
					...(Number.isInteger(paragraph.listLevel) ? { listLevel: paragraph.listLevel as number } : {}),
					...(Number.isInteger(paragraph.numId) ? { numId: paragraph.numId as number } : {}),
				};
			});
			const inheritance = parseParagraphInheritance(record.inherit);
			const resolvedAnchor = findParagraphAnchor(context.session, anchor, 'anchor');
			const templateAnchor = typeof record.templateAnchor === 'string' ? record.templateAnchor : anchor;
			const resolvedTemplate = templateAnchor === anchor
				? resolvedAnchor
				: findParagraphAnchor(context.session, templateAnchor, 'templateAnchor');
			const result = applyInsertParagraphsInPart(resolvedAnchor.partXml, {
				anchorBlockId: resolvedAnchor.blockId,
				templateParagraphXml: resolvedTemplate.paragraphXml,
				paragraphs,
				placement,
				inherit: inheritance,
				anchor,
				templateAnchor,
			});
			setPartXmlForLocation(context.session, resolvedAnchor.location, result.partXml);
			if (resolvedAnchor.location.part === 'body') {
				acc.documentXml = result.partXml;
			}
			acc.changedIds.push(resolvedAnchor.blockId, ...result.createdAnchors);
			acc.createdIds.push(...result.createdAnchors);
			acc.createdAnchors.push(...result.createdAnchors);
			acc.structuralMutations.push({
				op: opId,
				anchor: result.anchor,
				placement: result.placement,
				templateBlockId: resolvedTemplate.blockId,
				createdBlockIds: result.createdBlockIds,
				createdAnchors: result.createdAnchors,
				inheritedListProperties: result.inheritedListProperties,
			});
			acc.preview.push({
				id: anchor,
				field: 'insertParagraphs',
				before: null,
				after: {
					placement: result.placement,
					anchor: result.anchor,
					templateAnchor: result.templateAnchor,
					relationship: {
						placement: result.placement,
						anchor: result.anchor,
						templateAnchor: result.templateAnchor,
					},
					inheritance: result.inheritance,
					inheritedListProperties: result.inheritedListProperties,
					createdAnchors: result.createdAnchors,
					createdBlockIds: result.createdBlockIds,
					paragraphCount: paragraphs.length,
				},
			});
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
						acc.documentXml = result.partXml;
					}
				}
			}
			if (replacementCount === 0) {
				acc.warnings.push(`No matches found for query "${query}".`);
			} else {
				acc.changedIds.push('document');
				acc.preview.push({
					id: 'document',
					field: 'replaceText',
					before: query,
					after: replacement,
				});
			}
			break;
		}
		case 'docx.insertText': {
			const requestedBlockId = requireString(record.blockId, 'blockId');
			const offset = requireInteger(record.offset, 'offset');
			const text = requireString(record.text, 'text');
			const runId = typeof record.runId === 'string' ? record.runId : undefined;
			rejectWriteOnlyExcludedId(requestedBlockId, 'blockId');
			if (runId) rejectWriteOnlyExcludedId(runId, 'runId');
			const resolved = resolveTextTarget(context.session, {
				blockId: requestedBlockId,
				offset,
				...(runId ? { runId } : {}),
				...(typeof record.anchor === 'string' ? { anchor: record.anchor } : {}),
			}, 'blockId');
			const blockId = resolved.position.blockId;
			const location = resolved.location;
			let partXml = resolved.partXml;
			partXml = applyInsertTextInPart(partXml, resolved.position, text);
			setPartXmlForLocation(context.session, location, partXml);
			if (location.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(blockId);
			acc.preview.push({ id: blockId, field: 'insertText', before: null, after: { offset, text } });
			break;
		}
		case 'docx.deleteRange': {
			const parsedRange = parseTextRange(record.range, 'range');
			const start = resolveTextTarget(context.session, parsedRange.start, 'range.start.blockId');
			const end = resolveTextTarget(context.session, parsedRange.end, 'range.end.blockId');
			const range = { start: start.position, end: end.position };
			const startLocation = start.location;
			let partXml = start.partXml;
			partXml = applyDeleteRangeInPart(partXml, range);
			setPartXmlForLocation(context.session, startLocation, partXml);
			if (startLocation.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(range.start.blockId, range.end.blockId);
			acc.preview.push({ id: range.start.blockId, field: 'deleteRange', before: range, after: null });
			break;
		}
		case 'docx.deleteBlock': {
			const requestedBlockId = requireString(record.blockId, 'blockId');
			rejectWriteOnlyExcludedId(requestedBlockId, 'blockId');
			const resolved = resolveParagraphTarget(context.session, requestedBlockId, record.anchor, 'blockId');
			const blockId = resolved.blockId;
			const location = resolved.location;
			let partXml = resolved.partXml;
			partXml = applyDeleteParagraphInPart(partXml, blockId);
			setPartXmlForLocation(context.session, location, partXml);
			if (location.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(blockId);
			acc.preview.push({ id: blockId, field: 'deleteBlock', before: 'paragraph', after: null });
			break;
		}
		case 'docx.insertParagraphBreak': {
			const requestedBlockId = requireString(record.blockId, 'blockId');
			const offset = requireInteger(record.offset, 'offset');
			const runId = typeof record.runId === 'string' ? record.runId : undefined;
			rejectWriteOnlyExcludedId(requestedBlockId, 'blockId');
			if (runId) rejectWriteOnlyExcludedId(runId, 'runId');
			const resolved = resolveTextTarget(context.session, {
				blockId: requestedBlockId,
				offset,
				...(runId ? { runId } : {}),
				...(typeof record.anchor === 'string' ? { anchor: record.anchor } : {}),
			}, 'blockId');
			const blockId = resolved.position.blockId;
			const location = resolved.location;
			let partXml = resolved.partXml;
			const result = applyInsertParagraphBreakInPart(partXml, resolved.position);
			partXml = result.partXml;
			setPartXmlForLocation(context.session, location, partXml);
			if (location.part === 'body') {
				acc.documentXml = partXml;
			}
			acc.changedIds.push(blockId, ...result.createdBlockIds);
			acc.createdIds.push(...result.createdBlockIds);
			acc.preview.push({
				id: blockId,
				field: 'insertParagraphBreak',
				before: null,
				after: { offset, createdBlockIds: result.createdBlockIds, inheritedListProperties: result.inheritedListProperties },
			});
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
			acc.documentXml = applyReplaceBodyParagraphs(acc.documentXml, paragraphs);
			acc.changedIds.push('body');
			acc.preview.push({
				id: 'body',
				field: 'replaceBodyParagraphs',
				before: null,
				after: { paragraphCount: paragraphs.length > 0 ? paragraphs.length : 1 },
			});
			break;
		}
	}
}

export async function executeDocxOp(
	context: DocxOpExecutionContext,
	op: DocumentOp,
): Promise<DocxOpExecutionResult> {
	const record = asRecord(op);
	const opId = String(op.op);
	// Structural mutations need persistent identities before positional resolution.
	// Dry runs operate on a cloned session, so this remains side-effect free there.
	for (const partPath of context.session.listLoadedPartPaths()) {
		const partXml = context.session.getPartXml(partPath);
		const anchoredXml = ensureParagraphAnchors(partXml);
		if (anchoredXml !== partXml) {
			context.session.setPartXml(partPath, anchoredXml);
		}
	}
	const acc: DocxOpAccumulator = {
		documentXml: context.session.getDocumentXml(),
		changedIds: [],
		createdIds: [],
		createdAnchors: [],
		structuralMutations: [],
		preview: [],
		warnings: [],
	};

	switch (opId) {
		case 'docx.removeComments':
		case 'docx.setCoreProperties':
			await executeCommentsAndMetadataOp(context, opId, record, acc);
			break;
		case 'docx.setRunStyle':
		case 'docx.setParagraphStyle':
		case 'docx.setParagraphDefaultRunStyle':
		case 'docx.setParagraphLayout':
		case 'docx.setSectionLayout':
		case 'docx.setParagraphBottomBorder':
			executeFormattingOp(context, opId, record, acc);
			break;
		case 'docx.insertTable':
		case 'docx.setCellText':
		case 'docx.setCellStyle':
		case 'docx.deleteTable':
			executeTableOp(context, opId, record, acc);
			break;
		case 'docx.insertImage':
		case 'docx.replaceImage':
			await executeMediaOp(context, opId, record, acc);
			break;
		case 'docx.insertHyperlink':
		case 'docx.removeHyperlink':
			await executeHyperlinkOp(context, opId, record, acc);
			break;
		case 'docx.insertParagraphsBefore':
			executeTextEditOp(context, 'docx.insertParagraphs', { ...record, placement: 'before' }, acc);
			break;
		case 'docx.setRunText':
		case 'docx.insertParagraphsAfter':
		case 'docx.insertParagraphs':
		case 'docx.replaceText':
		case 'docx.insertText':
		case 'docx.deleteRange':
		case 'docx.deleteBlock':
		case 'docx.insertParagraphBreak':
		case 'docx.replaceBodyParagraphs':
			executeTextEditOp(context, opId, record, acc);
			break;
		default:
			throw createAiError(AI_ERROR_CODES.UNKNOWN_OP, `Unknown DOCX operation: ${opId}.`, { op: opId });
	}

	context.session.setDocumentXml(acc.documentXml);

	return acc;
}
