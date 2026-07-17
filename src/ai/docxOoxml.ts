const RUN_PATTERN = /<w:r\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:r>)/g;
import { extractDocxRunText } from '../docxXmlText';

export interface ParsedDocxRun {
	text: string;
	bold: boolean;
	italic: boolean;
	underline: boolean;
	fontFamily: string | null;
	fontSizePt: number | null;
	color: string | null;
}

export interface ParsedDocxParagraph {
	style: string | null;
	runs: ParsedDocxRun[];
	text: string;
	inlineImage?: {
		relationshipId: string;
	};
}

export interface ParsedDocxInlineImage {
	relationshipId: string;
}

export interface ParsedDocxCell {
	paragraphs: ParsedDocxParagraph[];
	text: string;
}

export interface ParsedDocxTable {
	rows: ParsedDocxCell[][];
}

export type ParsedDocxBodyBlock =
	| { kind: 'paragraph'; paragraph: ParsedDocxParagraph }
	| { kind: 'image'; paragraph: ParsedDocxParagraph }
	| { kind: 'table'; table: ParsedDocxTable };

export function paragraphContainsDrawing(paragraphXml: string): boolean {
	return /<w:drawing\b/.test(paragraphXml);
}

export function extractBlipEmbedId(paragraphXml: string): string | null {
	return /r:embed="([^"]+)"/.exec(paragraphXml)?.[1] ?? null;
}

export function parseInlineImageInfo(paragraphXml: string): ParsedDocxInlineImage | null {
	if (!paragraphContainsDrawing(paragraphXml)) {
		return null;
	}
	const relationshipId = extractBlipEmbedId(paragraphXml);
	if (!relationshipId) {
		return null;
	}
	return { relationshipId };
}

function resolveRelationshipTarget(relsXml: string | null, relationshipId: string): string | null {
	if (!relsXml) return null;
	const match = new RegExp(
		`<Relationship\\b[^>]*Id="${relationshipId}"[^>]*Target="([^"]+)"`,
	).exec(relsXml);
	return match?.[1] ?? null;
}

export function resolveInlineImageMediaPath(
	relsXml: string | null,
	relationshipId: string,
): string | null {
	const target = resolveRelationshipTarget(relsXml, relationshipId);
	if (!target) return null;
	return target.startsWith('word/') ? target : `word/${target.replace(/^media\//, 'media/')}`;
}

function getAttribute(xml: string, name: string): string | null {
	const pattern = new RegExp(`(?:\\w+:)?${name}="([^"]*)"`);
	return xml.match(pattern)?.[1] ?? null;
}

function isFlagEnabled(propertiesXml: string, tag: string): boolean {
	const match = new RegExp(`<w:${tag}\\b([^>]*)>|<w:${tag}\\b([^>]*)/>`).exec(propertiesXml);
	if (!match) return false;
	const attributes = `${match[1] ?? ''}${match[2] ?? ''}`;
	const valMatch = /w:val="([^"]*)"/.exec(attributes);
	if (!valMatch) return true;
	const val = (valMatch[1] ?? '').toLowerCase();
	return val !== 'false' && val !== '0' && val !== 'off';
}

function getRunProperties(runXml: string): string {
	const match = /<w:rPr\b[^>]*>([\s\S]*?)<\/w:rPr>/.exec(runXml);
	return match?.[1] ?? '';
}

function getParagraphProperties(paragraphXml: string): string {
	const match = /<w:pPr\b[^>]*>([\s\S]*?)<\/w:pPr>/.exec(paragraphXml);
	return match?.[1] ?? '';
}

function parseRun(runXml: string, runInnerXml: string): ParsedDocxRun {
	const runProperties = getRunProperties(runXml);
	const fontSizeHalfPoints = getAttribute(runProperties, 'sz');
	const parsedFontSize = fontSizeHalfPoints ? Number(fontSizeHalfPoints) : Number.NaN;
	const latinMatch = /<w:rFonts\b[^>]*w:ascii="([^"]*)"/.exec(runProperties)
		?? /<w:rFonts\b[^>]*w:hAnsi="([^"]*)"/.exec(runProperties);

	return {
		text: extractDocxRunText(runInnerXml),
		bold: isFlagEnabled(runProperties, 'b'),
		italic: isFlagEnabled(runProperties, 'i'),
		underline: isFlagEnabled(runProperties, 'u'),
		fontFamily: latinMatch?.[1] ?? null,
		fontSizePt: Number.isFinite(parsedFontSize) ? parsedFontSize / 2 : null,
		color: getAttribute(runProperties, 'color'),
	};
}

export function parseParagraph(paragraphXml: string): ParsedDocxParagraph {
	const propertiesXml = getParagraphProperties(paragraphXml);
	const styleMatch = /<w:pStyle\b[^>]*w:val="([^"]*)"/.exec(propertiesXml);
	const runs: ParsedDocxRun[] = [];
	let runMatch: RegExpExecArray | null;
	RUN_PATTERN.lastIndex = 0;

	while ((runMatch = RUN_PATTERN.exec(paragraphXml)) !== null) {
		const runInner = runMatch[1];
		if (runInner === undefined) continue;
		const parsed = parseRun(runMatch[0], runInner);
		if (parsed.text.length > 0 || runs.length === 0) {
			runs.push(parsed);
		}
	}

	const text = runs.map((run) => run.text).join('');
	const inlineImage = parseInlineImageInfo(paragraphXml);
	return {
		style: styleMatch?.[1] ?? null,
		runs,
		text,
		...(inlineImage ? { inlineImage } : {}),
	};
}

function parseCell(cellXml: string): ParsedDocxCell {
	const paragraphs: ParsedDocxParagraph[] = [];
	const paragraphPattern = /<w:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:p>)/g;
	let match: RegExpExecArray | null;
	paragraphPattern.lastIndex = 0;

	while ((match = paragraphPattern.exec(cellXml)) !== null) {
		const inner = match[1];
		const paragraphXml = inner === undefined ? match[0] : `<w:p>${inner}</w:p>`;
		paragraphs.push(parseParagraph(paragraphXml));
	}

	return {
		paragraphs,
		text: paragraphs.map((paragraph) => paragraph.text).join('\n'),
	};
}

function parseTable(tableXml: string): ParsedDocxTable {
	const rows: ParsedDocxCell[][] = [];
	const rowPattern = /<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g;
	let rowMatch: RegExpExecArray | null;
	rowPattern.lastIndex = 0;

	while ((rowMatch = rowPattern.exec(tableXml)) !== null) {
		const rowInner = rowMatch[1] ?? '';
		const cells: ParsedDocxCell[] = [];
		const cellPattern = /<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/g;
		let cellMatch: RegExpExecArray | null;
		cellPattern.lastIndex = 0;

		while ((cellMatch = cellPattern.exec(rowInner)) !== null) {
			cells.push(parseCell(cellMatch[0]));
		}

		rows.push(cells);
	}

	return { rows };
}

function extractElementXml(source: string, startIndex: number, tag: 'p' | 'tbl'): { xml: string; nextIndex: number } | null {
	const openPrefix = `<w:${tag}`;
	const start = source.indexOf(openPrefix, startIndex);
	if (start === -1) return null;

	const openEnd = source.indexOf('>', start);
	if (openEnd === -1) return null;

	const openTag = source.slice(start, openEnd + 1);
	if (openTag.endsWith('/>')) {
		return {
			xml: openTag,
			nextIndex: openEnd + 1,
		};
	}

	const closeTag = `</w:${tag}>`;
	const closeIndex = source.indexOf(closeTag, openEnd + 1);
	if (closeIndex === -1) return null;

	return {
		xml: source.slice(start, closeIndex + closeTag.length),
		nextIndex: closeIndex + closeTag.length,
	};
}

export function extractElementXmlAt(
	source: string,
	startIndex: number,
	tag: 'p' | 'tbl',
): { xml: string; nextIndex: number } | null {
	return extractElementXml(source, startIndex, tag);
}

export function splitTopLevelBodyBlocks(bodyInnerXml: string): Array<{ tag: 'p' | 'tbl'; xml: string }> {
	const blocks: Array<{ tag: 'p' | 'tbl'; xml: string }> = [];
	let cursor = 0;

	while (cursor < bodyInnerXml.length) {
		const nextParagraph = bodyInnerXml.indexOf('<w:p', cursor);
		const nextTable = bodyInnerXml.indexOf('<w:tbl', cursor);
		if (nextParagraph === -1 && nextTable === -1) break;

		let tag: 'p' | 'tbl';
		if (nextTable === -1 || (nextParagraph !== -1 && nextParagraph < nextTable)) {
			tag = 'p';
			cursor = nextParagraph;
		} else {
			tag = 'tbl';
			cursor = nextTable;
		}

		const extracted = extractElementXml(bodyInnerXml, cursor, tag);
		if (!extracted) break;
		blocks.push({ tag, xml: extracted.xml });
		cursor = extracted.nextIndex;
	}

	return blocks;
}

export function getDocumentBodyInner(documentXml: string): string | null {
	return /<w:body\b[^>]*>([\s\S]*?)<\/w:body>/.exec(documentXml)?.[1] ?? null;
}

export function replaceDocumentBodyInner(documentXml: string, inner: string): string {
	return documentXml.replace(/(<w:body\b[^>]*>)([\s\S]*?)(<\/w:body>)/, `$1${inner}$3`);
}

export function parseTopLevelBlocks(innerXml: string): ParsedDocxBodyBlock[] {
	const blocks: ParsedDocxBodyBlock[] = [];
	for (const { tag, xml } of splitTopLevelBodyBlocks(innerXml)) {
		if (tag === 'p') {
			const paragraph = parseParagraph(xml);
			if (paragraph.inlineImage) {
				blocks.push({ kind: 'image', paragraph });
			} else {
				blocks.push({ kind: 'paragraph', paragraph });
			}
			continue;
		}
		blocks.push({ kind: 'table', table: parseTable(xml) });
	}

	return blocks;
}

export function getWrapperInner(documentXml: string, wrapperTag: 'body' | 'hdr' | 'ftr'): string | null {
	const pattern = new RegExp(`<w:${wrapperTag}\\b[^>]*>([\\s\\S]*?)<\\/w:${wrapperTag}>`);
	return pattern.exec(documentXml)?.[1] ?? null;
}

export function replaceWrapperInner(documentXml: string, wrapperTag: 'body' | 'hdr' | 'ftr', inner: string): string {
	const pattern = new RegExp(`(<w:${wrapperTag}\\b[^>]*>)([\\s\\S]*?)(<\\/w:${wrapperTag}>)`);
	return documentXml.replace(pattern, `$1${inner}$3`);
}

export interface ParsedDocxFootnote {
	id: number;
	type: string | null;
	blocks: ParsedDocxBodyBlock[];
}

export function parseFootnotesContainer(xml: string, containerTag: 'footnotes' | 'endnotes'): ParsedDocxFootnote[] {
	const footnotes: ParsedDocxFootnote[] = [];
	const footnoteTag = containerTag === 'footnotes' ? 'footnote' : 'endnote';
	const footnotePattern = new RegExp(`<w:${footnoteTag}\\b([^>]*)>([\\s\\S]*?)<\\/w:${footnoteTag}>`, 'g');
	let match: RegExpExecArray | null;

	while ((match = footnotePattern.exec(xml)) !== null) {
		const attributes = match[1] ?? '';
		const inner = match[2] ?? '';
		const idMatch = /w:id="(\d+)"/.exec(attributes);
		const typeMatch = /w:type="([^"]*)"/.exec(attributes);
		if (!idMatch) continue;
		footnotes.push({
			id: Number(idMatch[1]),
			type: typeMatch?.[1] ?? null,
			blocks: parseTopLevelBlocks(inner),
		});
	}

	return footnotes;
}

export function getFootnoteInner(xml: string, containerTag: 'footnotes' | 'endnotes', footnoteId: number): string | null {
	const footnoteTag = containerTag === 'footnotes' ? 'footnote' : 'endnote';
	const pattern = new RegExp(
		`<w:${footnoteTag}\\b[^>]*w:id="${footnoteId}"[^>]*>([\\s\\S]*?)<\\/w:${footnoteTag}>`,
	);
	return pattern.exec(xml)?.[1] ?? null;
}

export function replaceFootnoteInner(
	xml: string,
	containerTag: 'footnotes' | 'endnotes',
	footnoteId: number,
	inner: string,
): string {
	const footnoteTag = containerTag === 'footnotes' ? 'footnote' : 'endnote';
	const pattern = new RegExp(
		`(<w:${footnoteTag}\\b[^>]*w:id="${footnoteId}"[^>]*>)([\\s\\S]*?)(<\\/w:${footnoteTag}>)`,
	);
	return xml.replace(pattern, `$1${inner}$3`);
}

export interface ParsedDocxComment {
	id: number;
	author: string | null;
	date: string | null;
	paragraphs: ParsedDocxParagraph[];
	text: string;
}

export function parseCommentsXml(xml: string): ParsedDocxComment[] {
	const comments: ParsedDocxComment[] = [];
	const commentPattern = /<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g;
	let match: RegExpExecArray | null;

	while ((match = commentPattern.exec(xml)) !== null) {
		const attributes = match[1] ?? '';
		const inner = match[2] ?? '';
		const idMatch = /w:id="(\d+)"/.exec(attributes);
		if (!idMatch) continue;
		const paragraphs: ParsedDocxParagraph[] = [];
		const paragraphPattern = /<w:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:p>)/g;
		let paragraphMatch: RegExpExecArray | null;
		paragraphPattern.lastIndex = 0;
		while ((paragraphMatch = paragraphPattern.exec(inner)) !== null) {
			const paragraphInner = paragraphMatch[1];
			const paragraphXml = paragraphInner === undefined
				? paragraphMatch[0]
				: `<w:p>${paragraphInner}</w:p>`;
			paragraphs.push(parseParagraph(paragraphXml));
		}
		comments.push({
			id: Number(idMatch[1]),
			author: /w:author="([^"]*)"/.exec(attributes)?.[1] ?? null,
			date: /w:date="([^"]*)"/.exec(attributes)?.[1] ?? null,
			paragraphs,
			text: paragraphs.map((paragraph) => paragraph.text).join('\n'),
		});
	}

	return comments;
}

export function parseDocumentBody(documentXml: string): ParsedDocxBodyBlock[] {
	const bodyMatch = /<w:body\b[^>]*>([\s\S]*?)<\/w:body>/.exec(documentXml);
	if (!bodyMatch?.[1]) return [];
	return parseTopLevelBlocks(bodyMatch[1]);
}
