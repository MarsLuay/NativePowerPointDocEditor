import {
	cellIdForLocation,
	docxIdPrefix,
	parseStableLocation,
	paragraphIdForLocation,
	type DocxStableLocation,
} from './docxStableIds';
import {
	extractElementXmlAt,
	getDocumentBodyInner,
	getFootnoteInner,
	getWrapperInner,
	parseFootnotesContainer,
	replaceFootnoteInner,
	replaceWrapperInner,
	splitTopLevelBodyBlocks,
} from './docxOoxml';
import { wrapperTagForPart, type DocxWrapperTag } from './docxParts';
import { AI_ERROR_CODES, createAiError } from './errors';

export interface TopLevelBlockPosition {
	kind: 'paragraph' | 'table';
	id: string;
	xml: string;
	startInBody: number;
	endInBody: number;
}

export type ParsedStableId =
	| { kind: 'paragraph'; paragraphIndex: number }
	| { kind: 'run'; paragraphIndex: number; runIndex: number }
	| { kind: 'table'; tableIndex: number }
	| { kind: 'cell'; tableIndex: number; rowIndex: number; colIndex: number };

export function parseStableId(id: string): ParsedStableId | null {
	const location = parseStableLocation(id);
	if (!location || location.part !== 'body') {
		return null;
	}
	switch (location.kind) {
		case 'paragraph':
			return { kind: 'paragraph', paragraphIndex: location.paragraphIndex };
		case 'run':
			return { kind: 'run', paragraphIndex: location.paragraphIndex, runIndex: location.runIndex ?? 0 };
		case 'table':
			return { kind: 'table', tableIndex: location.tableIndex ?? 0 };
		case 'cell':
			return {
				kind: 'cell',
				tableIndex: location.tableIndex ?? 0,
				rowIndex: location.rowIndex ?? 0,
				colIndex: location.colIndex ?? 0,
			};
	}
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
			throw createAiError(
				AI_ERROR_CODES.BLOCK_NOT_FOUND,
				`${footnoteTag} fn[${location.partNumber ?? 0}] was not found.`,
				{ field: 'blockId' },
			);
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

export function enumerateTopLevelBlockPositions(bodyInner: string, idPrefix = 'body'): TopLevelBlockPosition[] {
	const positions: TopLevelBlockPosition[] = [];
	let paragraphIndex = 0;
	let tableIndex = 0;
	let cursor = 0;

	while (cursor < bodyInner.length) {
		const nextParagraph = bodyInner.indexOf('<w:p', cursor);
		const nextTable = bodyInner.indexOf('<w:tbl', cursor);
		if (nextParagraph === -1 && nextTable === -1) break;

		let tag: 'p' | 'tbl';
		if (nextTable === -1 || (nextParagraph !== -1 && nextParagraph < nextTable)) {
			tag = 'p';
			cursor = nextParagraph;
		} else {
			tag = 'tbl';
			cursor = nextTable;
		}

		const extracted = extractElementXmlAt(bodyInner, cursor, tag);
		if (!extracted) break;

		const startInBody = cursor;
		const endInBody = extracted.nextIndex;
		if (tag === 'p') {
			positions.push({
				kind: 'paragraph',
				id: `${idPrefix}/p[${paragraphIndex}]`,
				xml: extracted.xml,
				startInBody,
				endInBody,
			});
			paragraphIndex++;
		} else {
			positions.push({
				kind: 'table',
				id: `${idPrefix}/tbl[${tableIndex}]`,
				xml: extracted.xml,
				startInBody,
				endInBody,
			});
			tableIndex++;
		}
		cursor = endInBody;
	}

	return positions;
}

export function findTopLevelBlock(bodyInner: string, blockId: string): TopLevelBlockPosition {
	const location = parseStableLocation(blockId);
	const idPrefix = location ? idPrefixForLocation(location) : 'body';
	const block = enumerateTopLevelBlockPositions(bodyInner, idPrefix).find((entry) => entry.id === blockId);
	if (!block) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Block ${blockId} was not found.`, { field: 'blockId' });
	}
	return block;
}

export function getParagraphXml(partXml: string, location: DocxStableLocation): string {
	const inner = getEditableInner(partXml, location);
	let currentIndex = 0;
	for (const { tag, xml } of splitTopLevelBodyBlocks(inner)) {
		if (tag !== 'p') continue;
		if (currentIndex === location.paragraphIndex) {
			return xml;
		}
		currentIndex++;
	}

	throw createAiError(
		AI_ERROR_CODES.BLOCK_NOT_FOUND,
		`Paragraph ${paragraphIdForLocation(location)} was not found.`,
		{ field: 'blockId' },
	);
}

export function replaceParagraphXml(
	partXml: string,
	location: DocxStableLocation,
	nextParagraphXml: string,
): string {
	const inner = getEditableInner(partXml, location);
	const idPrefix = idPrefixForLocation(location);
	let currentIndex = 0;

	for (const block of enumerateTopLevelBlockPositions(inner, idPrefix)) {
		if (block.kind !== 'paragraph') continue;
		if (currentIndex === location.paragraphIndex) {
			const nextInner = `${inner.slice(0, block.startInBody)}${nextParagraphXml}${inner.slice(block.endInBody)}`;
			return setEditableInner(partXml, location, nextInner);
		}
		currentIndex++;
	}

	throw createAiError(
		AI_ERROR_CODES.BLOCK_NOT_FOUND,
		`Paragraph ${paragraphIdForLocation(location)} was not found.`,
		{ field: 'blockId' },
	);
}

export function getTableCellXmlFromPart(partXml: string, location: DocxStableLocation): string {
	if (location.tableIndex === null || location.rowIndex === null || location.colIndex === null) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Cell location is missing table coordinates.', { field: 'cellId' });
	}

	const inner = getEditableInner(partXml, location);
	const idPrefix = idPrefixForLocation(location);
	let currentTableIndex = 0;

	for (const block of enumerateTopLevelBlockPositions(inner, idPrefix)) {
		if (block.kind !== 'table') continue;
		if (currentTableIndex !== location.tableIndex) {
			currentTableIndex++;
			continue;
		}

		const rowPattern = /<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g;
		let rowMatch: RegExpExecArray | null;
		let currentRowIndex = 0;
		while ((rowMatch = rowPattern.exec(block.xml)) !== null) {
			if (currentRowIndex !== location.rowIndex) {
				currentRowIndex++;
				continue;
			}

			const rowInner = rowMatch[1] ?? '';
			let currentColIndex = 0;
			const cellPattern = /<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/g;
			let cellMatch: RegExpExecArray | null;
			while ((cellMatch = cellPattern.exec(rowInner)) !== null) {
				if (currentColIndex === location.colIndex) {
					return cellMatch[0];
				}
				currentColIndex++;
			}

			throw createAiError(
				AI_ERROR_CODES.BLOCK_NOT_FOUND,
				`Cell ${cellIdForLocation(location)} was not found.`,
				{ field: 'cellId' },
			);
		}

		throw createAiError(
			AI_ERROR_CODES.BLOCK_NOT_FOUND,
			`Row ${location.rowIndex} in ${docxIdPrefix(location.part, location.partNumber)}/tbl[${location.tableIndex}] was not found.`,
			{ field: 'cellId' },
		);
	}

	throw createAiError(
		AI_ERROR_CODES.BLOCK_NOT_FOUND,
		`Table ${docxIdPrefix(location.part, location.partNumber)}/tbl[${location.tableIndex}] was not found.`,
		{ field: 'cellId' },
	);
}

export function replaceTableCellXmlInPart(
	partXml: string,
	location: DocxStableLocation,
	nextCellXml: string,
): string {
	if (location.tableIndex === null || location.rowIndex === null || location.colIndex === null) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Cell location is missing table coordinates.', { field: 'cellId' });
	}

	const inner = getEditableInner(partXml, location);
	const idPrefix = idPrefixForLocation(location);
	let currentTableIndex = 0;

	for (const block of enumerateTopLevelBlockPositions(inner, idPrefix)) {
		if (block.kind !== 'table') continue;
		if (currentTableIndex !== location.tableIndex) {
			currentTableIndex++;
			continue;
		}

		const rowPattern = /<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g;
		let rowMatch: RegExpExecArray | null;
		let currentRowIndex = 0;
		let nextTableXml = block.xml;
		let replaced = false;

		const rowMatches: Array<{ full: string; start: number; end: number; inner: string }> = [];
		while ((rowMatch = rowPattern.exec(block.xml)) !== null) {
			rowMatches.push({
				full: rowMatch[0],
				start: rowMatch.index,
				end: rowMatch.index + rowMatch[0].length,
				inner: rowMatch[1] ?? '',
			});
		}

		for (const row of rowMatches) {
			if (currentRowIndex !== location.rowIndex) {
				currentRowIndex++;
				continue;
			}

			let currentColIndex = 0;
			const cellPattern = /<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/g;
			let cellMatch: RegExpExecArray | null;
			while ((cellMatch = cellPattern.exec(row.inner)) !== null) {
				if (currentColIndex === location.colIndex) {
					const nextRowInner = `${row.inner.slice(0, cellMatch.index)}${nextCellXml}${row.inner.slice(cellMatch.index + cellMatch[0].length)}`;
					const nextRowXml = row.full.replace(row.inner, nextRowInner);
					nextTableXml = `${block.xml.slice(0, row.start)}${nextRowXml}${block.xml.slice(row.end)}`;
					replaced = true;
					break;
				}
				currentColIndex++;
			}
			break;
		}

		if (!replaced) {
			throw createAiError(
				AI_ERROR_CODES.BLOCK_NOT_FOUND,
				`Cell ${cellIdForLocation(location)} was not found.`,
				{ field: 'cellId' },
			);
		}

		const nextInner = `${inner.slice(0, block.startInBody)}${nextTableXml}${inner.slice(block.endInBody)}`;
		return setEditableInner(partXml, location, nextInner);
	}

	throw createAiError(
		AI_ERROR_CODES.BLOCK_NOT_FOUND,
		`Table ${docxIdPrefix(location.part, location.partNumber)}/tbl[${location.tableIndex}] was not found.`,
		{ field: 'cellId' },
	);
}

export function insertBlockAfterInPart(partXml: string, afterBlockId: string, blockXml: string): string {
	const location = parseStableLocation(afterBlockId);
	if (!location) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Block ${afterBlockId} was not found.`, { field: 'afterBlockId' });
	}
	if (location.part === 'footnotes' || location.part === 'endnotes') {
		throw createAiError(
			AI_ERROR_CODES.VALIDATION_FAILED,
			'insertTable and insertImage are only supported in body, header, and footer parts.',
			{ field: 'afterBlockId' },
		);
	}

	const inner = getEditableInner(partXml, location);
	const anchor = findTopLevelBlock(inner, afterBlockId);
	const sectPrMatch = /<w:sectPr\b[\s\S]*?<\/w:sectPr>\s*$/.exec(inner.slice(anchor.endInBody));
	if (sectPrMatch && sectPrMatch.index === 0) {
		const insertAt = anchor.endInBody;
		const nextInner = `${inner.slice(0, insertAt)}${blockXml}${inner.slice(insertAt)}`;
		return setEditableInner(partXml, location, nextInner);
	}

	const nextInner = `${inner.slice(0, anchor.endInBody)}${blockXml}${inner.slice(anchor.endInBody)}`;
	return setEditableInner(partXml, location, nextInner);
}

export function getTopLevelParagraphXml(documentXml: string, paragraphIndex: number): string {
	return getParagraphXml(documentXml, {
		part: 'body',
		partNumber: null,
		kind: 'paragraph',
		paragraphIndex,
		runIndex: null,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	});
}

export function replaceTopLevelParagraphXml(
	documentXml: string,
	paragraphIndex: number,
	nextParagraphXml: string,
): string {
	return replaceParagraphXml(documentXml, {
		part: 'body',
		partNumber: null,
		kind: 'paragraph',
		paragraphIndex,
		runIndex: null,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	}, nextParagraphXml);
}

export function getTableCellXml(
	documentXml: string,
	tableIndex: number,
	rowIndex: number,
	colIndex: number,
): string {
	return getTableCellXmlFromPart(documentXml, {
		part: 'body',
		partNumber: null,
		kind: 'cell',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex,
		rowIndex,
		colIndex,
	});
}

export function replaceTableCellXml(
	documentXml: string,
	tableIndex: number,
	rowIndex: number,
	colIndex: number,
	nextCellXml: string,
): string {
	return replaceTableCellXmlInPart(documentXml, {
		part: 'body',
		partNumber: null,
		kind: 'cell',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex,
		rowIndex,
		colIndex,
	}, nextCellXml);
}

export function insertBlockAfter(documentXml: string, afterBlockId: string, blockXml: string): string {
	return insertBlockAfterInPart(documentXml, afterBlockId, blockXml);
}

export function getDocumentBodyInnerOrThrow(documentXml: string): string {
	const bodyInner = getDocumentBodyInner(documentXml);
	if (bodyInner === null) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Missing w:body in document.xml.');
	}
	return bodyInner;
}

export function listFootnoteIds(partXml: string, part: 'footnotes' | 'endnotes'): number[] {
	const containerTag = part === 'footnotes' ? 'footnotes' : 'endnotes';
	return parseFootnotesContainer(partXml, containerTag)
		.filter((footnote) => footnote.type !== 'separator' && footnote.type !== 'continuationSeparator')
		.map((footnote) => footnote.id);
}
