import JSZip from 'jszip';
import { listDocxDescribeParts, scanDocxReviewState } from './docxParts';
import {
	parseCommentsXml,
	parseDocumentBody,
	parseFootnotesContainer,
	parseTopLevelBlocks,
	getWrapperInner,
	resolveInlineImageMediaPath,
	type ParsedDocxParagraph,
	type ParsedDocxRunStyle,
	type ParsedDocxTable,
} from './docxOoxml';
import { docxIdPrefix } from './docxStableIds';
import {
	parseDocumentSectionLayouts,
	type DocxParagraphLayout,
	type DocxSectionLayout,
} from './docxLayout';

export interface DocxDescribedRun {
	id: string;
	text: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	fontFamily?: string | null;
	fontSizePt?: number | null;
	color?: string | null;
}

export type DocxDescribedRunStyle = Omit<DocxDescribedRun, 'id' | 'text'>;

export interface DocxDescribeScope {
	sources: string[];
	body: 'top-level p and tbl in each included part';
	writeExcluded: string[];
	review?: {
		hasTrackChanges: boolean;
		hasComments: boolean;
	};
}

export interface DocxDescribedBlock {
	id: string;
	kind: 'paragraph' | 'image' | 'table' | 'tableCell' | 'comment';
	part?: string;
	style?: string | null;
	layout?: DocxParagraphLayout;
	defaultRunStyle?: DocxDescribedRunStyle;
	text?: string | null;
	runs?: DocxDescribedRun[];
	relationshipId?: string;
	mediaPath?: string | null;
	rows?: number;
	cols?: number;
	row?: number;
	col?: number;
	cells?: DocxDescribedBlock[];
	author?: string | null;
	date?: string | null;
	/** Present on comment replies when `commentsExtended.xml` threaded them. */
	parentId?: number | null;
}

export interface DocxDescribeSnapshot {
	format: 'docx';
	file: string;
	blockCount: number;
	scope: DocxDescribeScope;
	blocks: DocxDescribedBlock[];
	sections: DocxSectionLayout[];
}

function mapRunStyle(run: ParsedDocxRunStyle): DocxDescribedRunStyle {
	return {
		...(run.bold ? { bold: true } : {}),
		...(run.italic ? { italic: true } : {}),
		...(run.underline ? { underline: true } : {}),
		...(run.fontFamily ? { fontFamily: run.fontFamily } : {}),
		...(run.fontSizePt !== null ? { fontSizePt: run.fontSizePt } : {}),
		...(run.color ? { color: run.color } : {}),
	};
}

function mapRuns(idPrefix: string, paragraphIndex: number, paragraph: ParsedDocxParagraph): DocxDescribedRun[] {
	return paragraph.runs.map((run, runIndex) => ({
		id: `${idPrefix}/p[${paragraphIndex}]/r[${runIndex}]`,
		text: run.text,
		...mapRunStyle(run),
	}));
}

function mapParagraphBlock(
	idPrefix: string,
	paragraphIndex: number,
	paragraph: ParsedDocxParagraph,
	partLabel: string,
): DocxDescribedBlock {
	return {
		id: `${idPrefix}/p[${paragraphIndex}]`,
		kind: 'paragraph',
		part: partLabel,
		style: paragraph.style,
		layout: paragraph.layout,
		...(paragraph.defaultRunStyle ? { defaultRunStyle: mapRunStyle(paragraph.defaultRunStyle) } : {}),
		text: paragraph.text,
		runs: mapRuns(idPrefix, paragraphIndex, paragraph),
	};
}

function mapImageBlock(
	idPrefix: string,
	paragraphIndex: number,
	paragraph: ParsedDocxParagraph,
	mediaPath: string | null,
	partLabel: string,
): DocxDescribedBlock {
	const relationshipId = paragraph.inlineImage?.relationshipId;
	if (!relationshipId) {
		throw new Error(`Image paragraph ${idPrefix}/p[${paragraphIndex}] is missing a relationship id.`);
	}

	return {
		id: `${idPrefix}/p[${paragraphIndex}]`,
		kind: 'image',
		part: partLabel,
		style: paragraph.style,
		text: paragraph.text.length > 0 ? paragraph.text : null,
		relationshipId,
		mediaPath,
	};
}

function mapTableBlock(
	idPrefix: string,
	tableIndex: number,
	table: ParsedDocxTable,
	partLabel: string,
): DocxDescribedBlock {
	const rows = table.rows.length;
	const cols = table.rows.reduce((max, row) => Math.max(max, row.length), 0);
	const cells: DocxDescribedBlock[] = [];

	for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex++) {
		const row = table.rows[rowIndex] ?? [];
		for (let colIndex = 0; colIndex < row.length; colIndex++) {
			const cell = row[colIndex];
			if (!cell) continue;
			const primaryParagraph = cell.paragraphs[0];
			const cellId = `${idPrefix}/tbl[${tableIndex}]/tr[${rowIndex}]/tc[${colIndex}]`;
			cells.push({
				id: cellId,
				kind: 'tableCell',
				part: partLabel,
				row: rowIndex,
				col: colIndex,
				text: cell.text,
				style: primaryParagraph?.style ?? null,
				...(primaryParagraph ? { layout: primaryParagraph.layout } : {}),
				...(primaryParagraph?.defaultRunStyle
					? { defaultRunStyle: mapRunStyle(primaryParagraph.defaultRunStyle) }
					: {}),
				runs: primaryParagraph
					? primaryParagraph.runs.map((run, runIndex) => ({
						id: `${cellId}/r[${runIndex}]`,
						text: run.text,
						...(run.bold ? { bold: true } : {}),
						...(run.italic ? { italic: true } : {}),
						...(run.underline ? { underline: true } : {}),
					}))
					: [],
			});
		}
	}

	return {
		id: `${idPrefix}/tbl[${tableIndex}]`,
		kind: 'table',
		part: partLabel,
		rows,
		cols,
		cells,
	};
}

function appendParsedBlocks(
	blocks: ParsedDocxBodyBlock[],
	target: DocxDescribedBlock[],
	idPrefix: string,
	partLabel: string,
	relsXml: string | null,
): void {
	let paragraphIndex = 0;
	let tableIndex = 0;

	for (const block of blocks) {
		if (block.kind === 'image') {
			const mediaPath = resolveInlineImageMediaPath(
				relsXml,
				block.paragraph.inlineImage!.relationshipId,
			);
			target.push(mapImageBlock(idPrefix, paragraphIndex, block.paragraph, mediaPath, partLabel));
			paragraphIndex++;
			continue;
		}
		if (block.kind === 'paragraph') {
			target.push(mapParagraphBlock(idPrefix, paragraphIndex, block.paragraph, partLabel));
			paragraphIndex++;
			continue;
		}
		target.push(mapTableBlock(idPrefix, tableIndex, block.table, partLabel));
		tableIndex++;
	}
}

type ParsedDocxBodyBlock = ReturnType<typeof parseDocumentBody>[number];

function describePartLabel(part: string, partNumber: number | null): string {
	if (part === 'body') return 'body';
	if (partNumber === null) return part;
	return `${part}/${partNumber}`;
}

export async function describeDocxFromBuffer(buffer: ArrayBuffer, filePath: string): Promise<DocxDescribeSnapshot> {
	const zip = await JSZip.loadAsync(buffer.slice(0));
	const documentXml = (await zip.file('word/document.xml')?.async('string')) ?? '';
	const relsXml = (await zip.file('word/_rels/document.xml.rels')?.async('string')) ?? null;
	const blocks: DocxDescribedBlock[] = [];
	const sources: string[] = [];

	for (const listed of listDocxDescribeParts(zip)) {
		const partXml = await zip.file(listed.path)?.async('string');
		if (!partXml) continue;

		sources.push(listed.path);
		const label = describePartLabel(listed.part, listed.partNumber);

		if (listed.part === 'footnotes' || listed.part === 'endnotes') {
			const containerTag = listed.part === 'footnotes' ? 'footnotes' : 'endnotes';
			const footnotes = parseFootnotesContainer(partXml, containerTag);
			for (const footnote of footnotes) {
				if (footnote.type === 'separator' || footnote.type === 'continuationSeparator') {
					continue;
				}
				const idPrefix = docxIdPrefix(listed.part, footnote.id);
				appendParsedBlocks(footnote.blocks, blocks, idPrefix, `${label}/fn[${footnote.id}]`, relsXml);
			}
			continue;
		}

		const wrapperTag = listed.part === 'body' ? 'body' : listed.part === 'header' ? 'hdr' : 'ftr';
		const inner = getWrapperInner(partXml, wrapperTag);
		if (!inner) continue;
		const idPrefix = docxIdPrefix(listed.part, listed.partNumber);
		appendParsedBlocks(parseTopLevelBlocks(inner), blocks, idPrefix, label, relsXml);
	}

	const commentsXml = await zip.file('word/comments.xml')?.async('string');
	if (commentsXml) {
		sources.push('word/comments.xml');
		const commentsExtendedXml = await zip.file('word/commentsExtended.xml')?.async('string');
		if (commentsExtendedXml) {
			sources.push('word/commentsExtended.xml');
		}
		for (const comment of parseCommentsXml(commentsXml, commentsExtendedXml)) {
			blocks.push({
				id: `comments/c[${comment.id}]`,
				kind: 'comment',
				part: 'comments',
				author: comment.author,
				date: comment.date,
				...(comment.parentId != null ? { parentId: comment.parentId } : {}),
				text: comment.text,
			});
		}
	}

	const review = await scanDocxReviewState(zip);

	return {
		format: 'docx',
		file: filePath,
		blockCount: blocks.length,
		scope: {
			sources,
			body: 'top-level p and tbl in each included part',
			writeExcluded: ['comments', 'trackChanges'],
			review,
		},
		blocks,
		sections: parseDocumentSectionLayouts(documentXml),
	};
}
