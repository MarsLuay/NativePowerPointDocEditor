export type DocxPartKind = 'body' | 'header' | 'footer' | 'footnotes' | 'endnotes';

export type DocxStableBlockKind = 'paragraph' | 'run' | 'table' | 'cell';

export interface DocxStableLocation {
	part: DocxPartKind;
	partNumber: number | null;
	kind: DocxStableBlockKind;
	paragraphIndex: number;
	runIndex: number | null;
	tableIndex: number | null;
	rowIndex: number | null;
	colIndex: number | null;
}

const BODY_PARAGRAPH_ID = /^body\/p\[(\d+)\]$/;
const BODY_RUN_ID = /^body\/p\[(\d+)\]\/r\[(\d+)\]$/;
const BODY_TABLE_ID = /^body\/tbl\[(\d+)\]$/;
const BODY_CELL_ID = /^body\/tbl\[(\d+)\]\/tr\[(\d+)\]\/tc\[(\d+)\]$/;

const HEADER_PARAGRAPH_ID = /^header\/(\d+)\/p\[(\d+)\]$/;
const HEADER_RUN_ID = /^header\/(\d+)\/p\[(\d+)\]\/r\[(\d+)\]$/;
const HEADER_TABLE_ID = /^header\/(\d+)\/tbl\[(\d+)\]$/;
const HEADER_CELL_ID = /^header\/(\d+)\/tbl\[(\d+)\]\/tr\[(\d+)\]\/tc\[(\d+)\]$/;

const FOOTER_PARAGRAPH_ID = /^footer\/(\d+)\/p\[(\d+)\]$/;
const FOOTER_RUN_ID = /^footer\/(\d+)\/p\[(\d+)\]\/r\[(\d+)\]$/;
const FOOTER_TABLE_ID = /^footer\/(\d+)\/tbl\[(\d+)\]$/;
const FOOTER_CELL_ID = /^footer\/(\d+)\/tbl\[(\d+)\]\/tr\[(\d+)\]\/tc\[(\d+)\]$/;

const FOOTNOTE_PARAGRAPH_ID = /^footnotes\/fn\[(\d+)\]\/p\[(\d+)\]$/;
const FOOTNOTE_RUN_ID = /^footnotes\/fn\[(\d+)\]\/p\[(\d+)\]\/r\[(\d+)\]$/;
const FOOTNOTE_TABLE_ID = /^footnotes\/fn\[(\d+)\]\/tbl\[(\d+)\]$/;
const FOOTNOTE_CELL_ID = /^footnotes\/fn\[(\d+)\]\/tbl\[(\d+)\]\/tr\[(\d+)\]\/tc\[(\d+)\]$/;

const ENDNOTE_PARAGRAPH_ID = /^endnotes\/en\[(\d+)\]\/p\[(\d+)\]$/;
const ENDNOTE_RUN_ID = /^endnotes\/en\[(\d+)\]\/p\[(\d+)\]\/r\[(\d+)\]$/;
const ENDNOTE_TABLE_ID = /^endnotes\/en\[(\d+)\]\/tbl\[(\d+)\]$/;
const ENDNOTE_CELL_ID = /^endnotes\/en\[(\d+)\]\/tbl\[(\d+)\]\/tr\[(\d+)\]\/tc\[(\d+)\]$/;

export function docxIdPrefix(part: DocxPartKind, partNumber: number | null): string {
	switch (part) {
		case 'body':
			return 'body';
		case 'header':
			return `header/${partNumber ?? 1}`;
		case 'footer':
			return `footer/${partNumber ?? 1}`;
		case 'footnotes':
			return `footnotes/fn[${partNumber ?? 0}]`;
		case 'endnotes':
			return `endnotes/en[${partNumber ?? 0}]`;
	}
}

export function parseStableLocation(id: string): DocxStableLocation | null {
	const bodyParagraph = BODY_PARAGRAPH_ID.exec(id);
	if (bodyParagraph) {
		return {
			part: 'body',
			partNumber: null,
			kind: 'paragraph',
			paragraphIndex: Number(bodyParagraph[1]),
			runIndex: null,
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const bodyRun = BODY_RUN_ID.exec(id);
	if (bodyRun) {
		return {
			part: 'body',
			partNumber: null,
			kind: 'run',
			paragraphIndex: Number(bodyRun[1]),
			runIndex: Number(bodyRun[2]),
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const bodyTable = BODY_TABLE_ID.exec(id);
	if (bodyTable) {
		return {
			part: 'body',
			partNumber: null,
			kind: 'table',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(bodyTable[1]),
			rowIndex: null,
			colIndex: null,
		};
	}

	const bodyCell = BODY_CELL_ID.exec(id);
	if (bodyCell) {
		return {
			part: 'body',
			partNumber: null,
			kind: 'cell',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(bodyCell[1]),
			rowIndex: Number(bodyCell[2]),
			colIndex: Number(bodyCell[3]),
		};
	}

	const headerParagraph = HEADER_PARAGRAPH_ID.exec(id);
	if (headerParagraph) {
		return {
			part: 'header',
			partNumber: Number(headerParagraph[1]),
			kind: 'paragraph',
			paragraphIndex: Number(headerParagraph[2]),
			runIndex: null,
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const headerRun = HEADER_RUN_ID.exec(id);
	if (headerRun) {
		return {
			part: 'header',
			partNumber: Number(headerRun[1]),
			kind: 'run',
			paragraphIndex: Number(headerRun[2]),
			runIndex: Number(headerRun[3]),
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const headerTable = HEADER_TABLE_ID.exec(id);
	if (headerTable) {
		return {
			part: 'header',
			partNumber: Number(headerTable[1]),
			kind: 'table',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(headerTable[2]),
			rowIndex: null,
			colIndex: null,
		};
	}

	const headerCell = HEADER_CELL_ID.exec(id);
	if (headerCell) {
		return {
			part: 'header',
			partNumber: Number(headerCell[1]),
			kind: 'cell',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(headerCell[2]),
			rowIndex: Number(headerCell[3]),
			colIndex: Number(headerCell[4]),
		};
	}

	const footerParagraph = FOOTER_PARAGRAPH_ID.exec(id);
	if (footerParagraph) {
		return {
			part: 'footer',
			partNumber: Number(footerParagraph[1]),
			kind: 'paragraph',
			paragraphIndex: Number(footerParagraph[2]),
			runIndex: null,
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const footerRun = FOOTER_RUN_ID.exec(id);
	if (footerRun) {
		return {
			part: 'footer',
			partNumber: Number(footerRun[1]),
			kind: 'run',
			paragraphIndex: Number(footerRun[2]),
			runIndex: Number(footerRun[3]),
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const footerTable = FOOTER_TABLE_ID.exec(id);
	if (footerTable) {
		return {
			part: 'footer',
			partNumber: Number(footerTable[1]),
			kind: 'table',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(footerTable[2]),
			rowIndex: null,
			colIndex: null,
		};
	}

	const footerCell = FOOTER_CELL_ID.exec(id);
	if (footerCell) {
		return {
			part: 'footer',
			partNumber: Number(footerCell[1]),
			kind: 'cell',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(footerCell[2]),
			rowIndex: Number(footerCell[3]),
			colIndex: Number(footerCell[4]),
		};
	}

	const footnoteParagraph = FOOTNOTE_PARAGRAPH_ID.exec(id);
	if (footnoteParagraph) {
		return {
			part: 'footnotes',
			partNumber: Number(footnoteParagraph[1]),
			kind: 'paragraph',
			paragraphIndex: Number(footnoteParagraph[2]),
			runIndex: null,
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const footnoteRun = FOOTNOTE_RUN_ID.exec(id);
	if (footnoteRun) {
		return {
			part: 'footnotes',
			partNumber: Number(footnoteRun[1]),
			kind: 'run',
			paragraphIndex: Number(footnoteRun[2]),
			runIndex: Number(footnoteRun[3]),
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const footnoteTable = FOOTNOTE_TABLE_ID.exec(id);
	if (footnoteTable) {
		return {
			part: 'footnotes',
			partNumber: Number(footnoteTable[1]),
			kind: 'table',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(footnoteTable[2]),
			rowIndex: null,
			colIndex: null,
		};
	}

	const footnoteCell = FOOTNOTE_CELL_ID.exec(id);
	if (footnoteCell) {
		return {
			part: 'footnotes',
			partNumber: Number(footnoteCell[1]),
			kind: 'cell',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(footnoteCell[2]),
			rowIndex: Number(footnoteCell[3]),
			colIndex: Number(footnoteCell[4]),
		};
	}

	const endnoteParagraph = ENDNOTE_PARAGRAPH_ID.exec(id);
	if (endnoteParagraph) {
		return {
			part: 'endnotes',
			partNumber: Number(endnoteParagraph[1]),
			kind: 'paragraph',
			paragraphIndex: Number(endnoteParagraph[2]),
			runIndex: null,
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const endnoteRun = ENDNOTE_RUN_ID.exec(id);
	if (endnoteRun) {
		return {
			part: 'endnotes',
			partNumber: Number(endnoteRun[1]),
			kind: 'run',
			paragraphIndex: Number(endnoteRun[2]),
			runIndex: Number(endnoteRun[3]),
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		};
	}

	const endnoteTable = ENDNOTE_TABLE_ID.exec(id);
	if (endnoteTable) {
		return {
			part: 'endnotes',
			partNumber: Number(endnoteTable[1]),
			kind: 'table',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(endnoteTable[2]),
			rowIndex: null,
			colIndex: null,
		};
	}

	const endnoteCell = ENDNOTE_CELL_ID.exec(id);
	if (endnoteCell) {
		return {
			part: 'endnotes',
			partNumber: Number(endnoteCell[1]),
			kind: 'cell',
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(endnoteCell[2]),
			rowIndex: Number(endnoteCell[3]),
			colIndex: Number(endnoteCell[4]),
		};
	}

	return null;
}

export function paragraphIdForLocation(location: DocxStableLocation): string {
	const prefix = docxIdPrefix(location.part, location.partNumber);
	return `${prefix}/p[${location.paragraphIndex}]`;
}

export function runIdForLocation(location: DocxStableLocation): string {
	if (location.runIndex === null) {
		throw new Error('runIdForLocation requires runIndex.');
	}
	return `${paragraphIdForLocation(location)}/r[${location.runIndex}]`;
}

export function tableIdForLocation(location: DocxStableLocation): string {
	const prefix = docxIdPrefix(location.part, location.partNumber);
	return `${prefix}/tbl[${location.tableIndex ?? 0}]`;
}

export function cellIdForLocation(location: DocxStableLocation): string {
	return `${tableIdForLocation(location)}/tr[${location.rowIndex ?? 0}]/tc[${location.colIndex ?? 0}]`;
}
