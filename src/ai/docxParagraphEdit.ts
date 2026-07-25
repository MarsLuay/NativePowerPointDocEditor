import {
	enumerateTopLevelBlockPositions,
	getParagraphXml,
	replaceParagraphXml,
} from './docxBlockResolver';
import { buildHyperlinkXml } from './docxHyperlink';
import { extractDocxRunText } from '../docxXmlText';
import { wrapperTagForPart, type DocxWrapperTag } from './docxParts';
import type { DocxStableLocation } from './docxStableIds';
import { docxIdPrefix, paragraphIdForLocation, parseStableLocation } from './docxStableIds';
import {
	getFootnoteInner,
	getWrapperInner,
	replaceFootnoteInner,
	replaceWrapperInner,
} from './docxOoxml';
import { AI_ERROR_CODES, createAiError } from './errors';

export interface DocxTextPosition {
	blockId: string;
	offset: number;
	runId?: string;
}

export interface DocxTextRange {
	start: DocxTextPosition;
	end: DocxTextPosition;
}

interface FlatRun {
	xml: string;
	text: string;
	globalStart: number;
	globalEnd: number;
	insideHyperlink: boolean;
	hyperlinkOpen?: string;
	hyperlinkClose?: string;
}

function encodeXmlText(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function extractElementXml(
	source: string,
	startIndex: number,
	tag: string,
): { xml: string; nextIndex: number } | null {
	const openPrefix = `<w:${tag}`;
	const start = source.indexOf(openPrefix, startIndex);
	if (start === -1) return null;

	const openEnd = source.indexOf('>', start);
	if (openEnd === -1) return null;

	const openTag = source.slice(start, openEnd + 1);
	if (openTag.endsWith('/>')) {
		return { xml: openTag, nextIndex: openEnd + 1 };
	}

	const closeTag = `</w:${tag}>`;
	const closeIndex = source.indexOf(closeTag, openEnd + 1);
	if (closeIndex === -1) return null;

	return {
		xml: source.slice(start, closeIndex + closeTag.length),
		nextIndex: closeIndex + closeTag.length,
	};
}

function decomposeParagraph(paragraphXml: string): {
	openTag: string;
	prefixXml: string;
	contentXml: string;
	closeTag: string;
} {
	const trimmed = paragraphXml.trim();
	const openMatch = /^<w:p\b[^>]*>/.exec(trimmed);
	if (!openMatch) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Invalid paragraph XML.');
	}
	const openTag = openMatch[0];
	const closeTag = '</w:p>';
	const closeIndex = trimmed.lastIndexOf(closeTag);
	if (closeIndex === -1) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Invalid paragraph XML.');
	}
	const inner = trimmed.slice(openTag.length, closeIndex);
	const pPr = /<w:pPr\b[\s\S]*?<\/w:pPr>/.exec(inner)?.[0]
		?? /<w:pPr\b[^>]*\/>/.exec(inner)?.[0]
		?? '';
	const contentStart = pPr ? inner.indexOf(pPr) + pPr.length : 0;
	return {
		openTag,
		prefixXml: pPr,
		contentXml: inner.slice(contentStart),
		closeTag,
	};
}

function composeParagraph(openTag: string, prefixXml: string, contentXml: string, closeTag: string): string {
	return `${openTag}${prefixXml}${contentXml}${closeTag}`;
}

function flattenParagraphRuns(contentXml: string): FlatRun[] {
	const runs: FlatRun[] = [];
	let cursor = 0;
	let globalOffset = 0;

	while (cursor < contentXml.length) {
		if (contentXml.startsWith('<w:hyperlink', cursor)) {
			const hyperlink = extractElementXml(contentXml, cursor, 'hyperlink');
			if (!hyperlink) break;
			const openTagEnd = hyperlink.xml.indexOf('>');
			const hyperlinkOpen = hyperlink.xml.slice(0, openTagEnd + 1);
			const hyperlinkClose = '</w:hyperlink>';
			const hyperlinkInner = hyperlink.xml.slice(openTagEnd + 1, hyperlink.xml.length - hyperlinkClose.length);
			let innerCursor = 0;
			while (innerCursor < hyperlinkInner.length) {
				const run = extractElementXml(hyperlinkInner, innerCursor, 'r');
				if (!run) break;
				const text = extractDocxRunText(run.xml);
				runs.push({
					xml: run.xml,
					text,
					globalStart: globalOffset,
					globalEnd: globalOffset + text.length,
					insideHyperlink: true,
					hyperlinkOpen,
					hyperlinkClose,
				});
				globalOffset += text.length;
				innerCursor = run.nextIndex;
			}
			cursor = hyperlink.nextIndex;
			continue;
		}

		const run = extractElementXml(contentXml, cursor, 'r');
		if (!run) {
			const nextTag = contentXml.indexOf('<', cursor + 1);
			if (nextTag === -1) break;
			cursor = nextTag;
			continue;
		}
		const text = extractDocxRunText(run.xml);
		runs.push({
			xml: run.xml,
			text,
			globalStart: globalOffset,
			globalEnd: globalOffset + text.length,
			insideHyperlink: false,
		});
		globalOffset += text.length;
		cursor = run.nextIndex;
	}

	return runs;
}

function setRunTextContent(runXml: string, text: string): string {
	const encoded = encodeXmlText(text);
	if (/<w:t\b/.test(runXml)) {
		return runXml.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/, `<w:t>${encoded}</w:t>`);
	}
	if (runXml.trim().endsWith('/>')) {
		return runXml.replace(/\/>$/, `><w:t>${encoded}</w:t></w:r>`);
	}
	return runXml.replace(/<\/w:r>$/, `<w:t>${encoded}</w:t></w:r>`);
}

function buildRunXml(text: string): string {
	return `<w:r><w:t>${encodeXmlText(text)}</w:t></w:r>`;
}

function rebuildContentFromRuns(runs: FlatRun[]): string {
	const parts: string[] = [];
	let index = 0;
	while (index < runs.length) {
		const run = runs[index]!;
		if (run.insideHyperlink && run.hyperlinkOpen && run.hyperlinkClose) {
			const group: FlatRun[] = [];
			while (index < runs.length && runs[index]?.hyperlinkOpen === run.hyperlinkOpen) {
				group.push(runs[index]!);
				index++;
			}
			parts.push(`${run.hyperlinkOpen}${group.map((entry) => entry.xml).join('')}${run.hyperlinkClose}`);
			continue;
		}
		parts.push(run.xml);
		index++;
	}
	return parts.join('');
}

function cloneRunTemplate(runXml: string): string {
	if (/<w:t\b/.test(runXml)) {
		return runXml.replace(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/, '<w:t></w:t>');
	}
	if (runXml.trim().endsWith('/>')) {
		return runXml.replace(/\/>$/, '><w:t></w:t></w:r>');
	}
	return runXml.replace(/<\/w:r>$/, '<w:t></w:t></w:r>');
}

export function getParagraphPlainText(paragraphXml: string): string {
	const { contentXml } = decomposeParagraph(paragraphXml);
	return flattenParagraphRuns(contentXml).map((run) => run.text).join('');
}

function insertTextInParagraphContent(contentXml: string, offset: number, text: string): string {
	const runs = flattenParagraphRuns(contentXml);
	const totalLength = runs.reduce((sum, run) => sum + run.text.length, 0);
	const clampedOffset = Math.max(0, Math.min(offset, totalLength));

	if (runs.length === 0) {
		return buildRunXml(text);
	}

	const nextRuns: FlatRun[] = [];
	let inserted = false;
	for (const run of runs) {
		if (!inserted && clampedOffset >= run.globalStart && clampedOffset <= run.globalEnd) {
			const localOffset = clampedOffset - run.globalStart;
			const before = run.text.slice(0, localOffset);
			const after = run.text.slice(localOffset);
			if (before.length > 0) {
				nextRuns.push({ ...run, xml: setRunTextContent(run.xml, before), text: before });
			}
			const template = cloneRunTemplate(run.xml);
			nextRuns.push({
				...run,
				xml: setRunTextContent(template, text),
				text,
				insideHyperlink: false,
				hyperlinkOpen: undefined,
				hyperlinkClose: undefined,
			});
			if (after.length > 0) {
				nextRuns.push({ ...run, xml: setRunTextContent(run.xml, after), text: after });
			}
			inserted = true;
			continue;
		}
		nextRuns.push(run);
	}

	if (!inserted) {
		const template = cloneRunTemplate(runs[runs.length - 1]!.xml);
		nextRuns.push({
			...runs[runs.length - 1]!,
			xml: setRunTextContent(template, text),
			text,
			globalStart: totalLength,
			globalEnd: totalLength + text.length,
			insideHyperlink: false,
			hyperlinkOpen: undefined,
			hyperlinkClose: undefined,
		});
	}

	return rebuildContentFromRuns(nextRuns);
}

function deleteRangeInParagraphContent(contentXml: string, startOffset: number, endOffset: number): string {
	if (startOffset >= endOffset) {
		return contentXml;
	}

	const runs = flattenParagraphRuns(contentXml);
	const nextRuns: FlatRun[] = [];

	for (const run of runs) {
		if (run.globalEnd <= startOffset || run.globalStart >= endOffset) {
			nextRuns.push(run);
			continue;
		}

		const localStart = Math.max(0, startOffset - run.globalStart);
		const localEnd = Math.min(run.text.length, endOffset - run.globalStart);
		const before = run.text.slice(0, localStart);
		const after = run.text.slice(localEnd);

		if (before.length > 0) {
			nextRuns.push({ ...run, xml: setRunTextContent(run.xml, before), text: before });
		}
		if (after.length > 0) {
			nextRuns.push({ ...run, xml: setRunTextContent(run.xml, after), text: after });
		}
	}

	return rebuildContentFromRuns(nextRuns);
}

function splitParagraphContent(contentXml: string, offset: number): { before: string; after: string } {
	const totalLength = getParagraphPlainText(`<w:p>${contentXml}</w:p>`).length;
	const clampedOffset = Math.max(0, Math.min(offset, totalLength));
	return {
		before: deleteRangeInParagraphContent(contentXml, clampedOffset, totalLength),
		after: deleteRangeInParagraphContent(contentXml, 0, clampedOffset),
	};
}

function wrapRangeWithHyperlink(
	contentXml: string,
	startOffset: number,
	endOffset: number,
	relationshipId: string,
	displayText: string | undefined,
	tooltip: string | undefined,
): string {
	if (startOffset >= endOffset) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Hyperlink range must not be empty.', { field: 'range' });
	}

	const runs = flattenParagraphRuns(contentXml);
	const selected: FlatRun[] = [];
	const before: FlatRun[] = [];
	const after: FlatRun[] = [];

	for (const run of runs) {
		if (run.globalEnd <= startOffset) {
			before.push(run);
			continue;
		}
		if (run.globalStart >= endOffset) {
			after.push(run);
			continue;
		}

		const localStart = Math.max(0, startOffset - run.globalStart);
		const localEnd = Math.min(run.text.length, endOffset - run.globalStart);
		const prefix = run.text.slice(0, localStart);
		const sliceText = run.text.slice(localStart, localEnd);
		const suffix = run.text.slice(localEnd);

		if (prefix.length > 0) {
			before.push({ ...run, xml: setRunTextContent(run.xml, prefix), text: prefix });
		}
		if (sliceText.length > 0) {
			selected.push({
				...run,
				xml: setRunTextContent(run.xml, sliceText),
				text: sliceText,
				insideHyperlink: false,
				hyperlinkOpen: undefined,
				hyperlinkClose: undefined,
			});
		}
		if (suffix.length > 0) {
			after.push({ ...run, xml: setRunTextContent(run.xml, suffix), text: suffix });
		}
	}

	const linkRunsXml = displayText !== undefined
		? buildRunXml(displayText)
		: selected.map((run) => run.xml).join('');
	const hyperlinkXml = buildHyperlinkXml(relationshipId, linkRunsXml, tooltip);
	return `${rebuildContentFromRuns(before)}${hyperlinkXml}${rebuildContentFromRuns(after)}`;
}

function unwrapHyperlinksInRange(contentXml: string, startOffset: number, endOffset: number): string {
	const runs = flattenParagraphRuns(contentXml);
	const nextRuns = runs.map((run) => {
		if (!run.insideHyperlink) return run;
		if (run.globalEnd <= startOffset || run.globalStart >= endOffset) return run;
		return {
			...run,
			insideHyperlink: false,
			hyperlinkOpen: undefined,
			hyperlinkClose: undefined,
		};
	});
	return rebuildContentFromRuns(nextRuns);
}

export function insertTextInParagraph(paragraphXml: string, offset: number, text: string): string {
	const { openTag, prefixXml, contentXml, closeTag } = decomposeParagraph(paragraphXml);
	const nextContent = insertTextInParagraphContent(contentXml, offset, text);
	return composeParagraph(openTag, prefixXml, nextContent, closeTag);
}

export function deleteRangeInParagraph(paragraphXml: string, startOffset: number, endOffset: number): string {
	const { openTag, prefixXml, contentXml, closeTag } = decomposeParagraph(paragraphXml);
	const nextContent = deleteRangeInParagraphContent(contentXml, startOffset, endOffset);
	return composeParagraph(openTag, prefixXml, nextContent, closeTag);
}

export function splitParagraphAtOffset(paragraphXml: string, offset: number): { before: string; after: string } {
	const { openTag, prefixXml, contentXml, closeTag } = decomposeParagraph(paragraphXml);
	const split = splitParagraphContent(contentXml, offset);
	return {
		before: composeParagraph(openTag, prefixXml, split.before, closeTag),
		after: composeParagraph(openTag, prefixXml, split.after, closeTag),
	};
}

export function wrapParagraphRangeWithHyperlink(
	paragraphXml: string,
	startOffset: number,
	endOffset: number,
	relationshipId: string,
	displayText: string | undefined,
	tooltip: string | undefined,
): string {
	const { openTag, prefixXml, contentXml, closeTag } = decomposeParagraph(paragraphXml);
	const nextContent = wrapRangeWithHyperlink(
		contentXml,
		startOffset,
		endOffset,
		relationshipId,
		displayText,
		tooltip,
	);
	return composeParagraph(openTag, prefixXml, nextContent, closeTag);
}

export function removeHyperlinkInParagraphRange(
	paragraphXml: string,
	startOffset: number,
	endOffset: number,
): string {
	const { openTag, prefixXml, contentXml, closeTag } = decomposeParagraph(paragraphXml);
	const nextContent = unwrapHyperlinksInRange(contentXml, startOffset, endOffset);
	return composeParagraph(openTag, prefixXml, nextContent, closeTag);
}

function footnoteContainerTag(part: DocxStableLocation['part']): 'footnotes' | 'endnotes' | null {
	if (part === 'footnotes') return 'footnotes';
	if (part === 'endnotes') return 'endnotes';
	return null;
}

function getEditableInner(partXml: string, location: Pick<DocxStableLocation, 'part' | 'partNumber'>): string {
	const footnoteTag = footnoteContainerTag(location.part);
	if (footnoteTag) {
		const inner = getFootnoteInner(partXml, footnoteTag, location.partNumber ?? 0);
		if (inner === null) {
			throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Footnote ${location.partNumber ?? 0} was not found.`, { field: 'blockId' });
		}
		return inner;
	}

	const wrapperTag = wrapperTagForPart(location.part) as DocxWrapperTag;
	const inner = getWrapperInner(partXml, wrapperTag);
	if (inner === null) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, `Missing w:${wrapperTag} wrapper in part XML.`);
	}
	return inner;
}

function setEditableInner(
	partXml: string,
	location: Pick<DocxStableLocation, 'part' | 'partNumber'>,
	inner: string,
): string {
	const footnoteTag = footnoteContainerTag(location.part);
	if (footnoteTag) {
		return replaceFootnoteInner(partXml, footnoteTag, location.partNumber ?? 0, inner);
	}
	const wrapperTag = wrapperTagForPart(location.part) as DocxWrapperTag;
	return replaceWrapperInner(partXml, wrapperTag, inner);
}

function idPrefixForLocation(location: Pick<DocxStableLocation, 'part' | 'partNumber'>): string {
	return docxIdPrefix(location.part, location.partNumber);
}

function resolveParagraphLocation(blockId: string): DocxStableLocation {
	const location = parseStableLocation(blockId);
	if (!location || location.kind !== 'paragraph') {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid blockId: ${blockId}.`, { field: 'blockId' });
	}
	return location;
}

function validateOptionalRunId(blockId: string, runId: string | undefined, offset: number, partXml: string, location: DocxStableLocation): void {
	if (!runId) return;
	const parsedRun = parseStableLocation(runId);
	if (!parsedRun || parsedRun.kind !== 'run') {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Invalid runId: ${runId}.`, { field: 'runId' });
	}
	if (
		parsedRun.part !== location.part
		|| parsedRun.partNumber !== location.partNumber
		|| parsedRun.paragraphIndex !== location.paragraphIndex
	) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `runId ${runId} does not match blockId ${blockId}.`, { field: 'runId' });
	}
	const paragraphXml = getParagraphXml(partXml, location);
	const runs = flattenParagraphRuns(decomposeParagraph(paragraphXml).contentXml);
	const target = runs[parsedRun.runIndex ?? 0];
	if (!target) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Run ${runId} was not found.`, { field: 'runId' });
	}
	if (offset < target.globalStart || offset > target.globalEnd) {
		throw createAiError(
			AI_ERROR_CODES.SCHEMA_INVALID,
			`offset ${offset} is outside run ${runId} (${target.globalStart}..${target.globalEnd}).`,
			{ field: 'offset' },
		);
	}
}

function validateRangePositions(range: DocxTextRange): {
	startLocation: DocxStableLocation;
	endLocation: DocxStableLocation;
} {
	const startLocation = resolveParagraphLocation(range.start.blockId);
	const endLocation = resolveParagraphLocation(range.end.blockId);
	if (
		startLocation.part !== endLocation.part
		|| startLocation.partNumber !== endLocation.partNumber
	) {
		throw createAiError(
			AI_ERROR_CODES.VALIDATION_FAILED,
			'Range start and end must be in the same DOCX part.',
			{ field: 'range' },
		);
	}
	if (startLocation.paragraphIndex > endLocation.paragraphIndex) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Range end must not precede range start.', { field: 'range' });
	}
	if (
		startLocation.paragraphIndex === endLocation.paragraphIndex
		&& range.start.offset > range.end.offset
	) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Range end offset must be >= start offset.', { field: 'range' });
	}
	return { startLocation, endLocation };
}

export function applyInsertTextInPart(
	partXml: string,
	position: DocxTextPosition,
	text: string,
): string {
	const location = resolveParagraphLocation(position.blockId);
	validateOptionalRunId(position.blockId, position.runId, position.offset, partXml, location);
	const paragraphXml = getParagraphXml(partXml, location);
	const nextParagraphXml = insertTextInParagraph(paragraphXml, position.offset, text);
	return replaceParagraphXml(partXml, location, nextParagraphXml);
}

export function applyDeleteRangeInPart(partXml: string, range: DocxTextRange): string {
	const { startLocation, endLocation } = validateRangePositions(range);

	if (startLocation.paragraphIndex === endLocation.paragraphIndex) {
		const paragraphXml = getParagraphXml(partXml, startLocation);
		const nextParagraphXml = deleteRangeInParagraph(paragraphXml, range.start.offset, range.end.offset);
		return replaceParagraphXml(partXml, startLocation, nextParagraphXml);
	}

	const inner = getEditableInner(partXml, startLocation);
	const idPrefix = idPrefixForLocation(startLocation);
	const blocks = enumerateTopLevelBlockPositions(inner, idPrefix);
	const startBlock = blocks.find((block) => block.id === range.start.blockId);
	const endBlock = blocks.find((block) => block.id === range.end.blockId);
	if (!startBlock || !endBlock) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, 'Range block anchors were not found.', { field: 'range' });
	}

	let startParagraphXml = deleteRangeInParagraph(
		getParagraphXml(partXml, startLocation),
		range.start.offset,
		getParagraphPlainText(getParagraphXml(partXml, startLocation)).length,
	);
	const endParagraphXml = deleteRangeInParagraph(
		getParagraphXml(partXml, endLocation),
		0,
		range.end.offset,
	);
	const endPlain = getParagraphPlainText(endParagraphXml);
	startParagraphXml = insertTextInParagraph(
		startParagraphXml,
		getParagraphPlainText(startParagraphXml).length,
		endPlain,
	);

	const removeEnd = endBlock.endInBody;
	const nextInner = `${inner.slice(0, startBlock.startInBody)}${startParagraphXml}${inner.slice(removeEnd)}`;
	return setEditableInner(partXml, startLocation, nextInner);
}

/**
 * Delete one complete paragraph without joining it to an adjacent paragraph.
 * Text-range deletion deliberately merges endpoints, which is wrong for
 * removing an accidental blank paragraph while retaining the heading after it.
 */
export function applyDeleteParagraphInPart(partXml: string, blockId: string): string {
	const location = resolveParagraphLocation(blockId);
	const inner = getEditableInner(partXml, location);
	const idPrefix = idPrefixForLocation(location);
	const blocks = enumerateTopLevelBlockPositions(inner, idPrefix);
	const block = blocks.find((entry) => entry.id === blockId);
	if (!block) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Block ${blockId} was not found.`, { field: 'blockId' });
	}
	if (blocks.length <= 1) {
		throw createAiError(
			AI_ERROR_CODES.VALIDATION_FAILED,
			'Cannot delete the only editable block in a DOCX part.',
			{ field: 'blockId' },
		);
	}

	return setEditableInner(
		partXml,
		location,
		`${inner.slice(0, block.startInBody)}${inner.slice(block.endInBody)}`,
	);
}

export function applyInsertParagraphBreakInPart(partXml: string, position: DocxTextPosition): string {
	const location = resolveParagraphLocation(position.blockId);
	validateOptionalRunId(position.blockId, position.runId, position.offset, partXml, location);
	const paragraphXml = getParagraphXml(partXml, location);
	const split = splitParagraphAtOffset(paragraphXml, position.offset);
	const inner = getEditableInner(partXml, location);
	const idPrefix = idPrefixForLocation(location);
	const block = enumerateTopLevelBlockPositions(inner, idPrefix).find((entry) => entry.id === position.blockId);
	if (!block) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Block ${position.blockId} was not found.`, { field: 'blockId' });
	}
	const nextInner = `${inner.slice(0, block.startInBody)}${split.before}${split.after}${inner.slice(block.endInBody)}`;
	return setEditableInner(partXml, location, nextInner);
}

export function applyInsertHyperlinkInPart(
	partXml: string,
	range: DocxTextRange,
	relationshipId: string,
	displayText: string | undefined,
	tooltip: string | undefined,
): string {
	const { startLocation, endLocation } = validateRangePositions(range);
	if (startLocation.paragraphIndex !== endLocation.paragraphIndex) {
		throw createAiError(
			AI_ERROR_CODES.VALIDATION_FAILED,
			'insertHyperlink is only supported within a single paragraph.',
			{ field: 'range' },
		);
	}
	const paragraphXml = getParagraphXml(partXml, startLocation);
	const nextParagraphXml = wrapParagraphRangeWithHyperlink(
		paragraphXml,
		range.start.offset,
		range.end.offset,
		relationshipId,
		displayText,
		tooltip,
	);
	return replaceParagraphXml(partXml, startLocation, nextParagraphXml);
}

export function applyRemoveHyperlinkInPart(partXml: string, range: DocxTextRange): string {
	const { startLocation, endLocation } = validateRangePositions(range);
	if (startLocation.paragraphIndex !== endLocation.paragraphIndex) {
		throw createAiError(
			AI_ERROR_CODES.VALIDATION_FAILED,
			'removeHyperlink is only supported within a single paragraph.',
			{ field: 'range' },
		);
	}
	const paragraphXml = getParagraphXml(partXml, startLocation);
	const nextParagraphXml = removeHyperlinkInParagraphRange(
		paragraphXml,
		range.start.offset,
		range.end.offset,
	);
	return replaceParagraphXml(partXml, startLocation, nextParagraphXml);
}

export function paragraphIdFromLocation(location: DocxStableLocation): string {
	return paragraphIdForLocation(location);
}
