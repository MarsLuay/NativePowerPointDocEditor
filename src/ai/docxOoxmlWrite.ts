import { AI_ERROR_CODES, createAiError } from './errors';

export { extractBlipEmbedId, paragraphContainsDrawing } from './docxOoxml';

const RUN_PATTERN = /<w:r\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:r>)/g;

export interface DocxRunStylePatch {
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	fontFamily?: string;
	fontSizePt?: number;
	color?: string | null;
}

export interface DocxParagraphStylePatch {
	name?: string;
	style?: string;
}

function encodeXmlText(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function decodeXmlText(value: string): string {
	return value
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&amp;/g, '&')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'");
}

function extractRuns(paragraphXml: string): Array<{ full: string; inner: string }> {
	const runs: Array<{ full: string; inner: string }> = [];
	let match: RegExpExecArray | null;
	RUN_PATTERN.lastIndex = 0;
	while ((match = RUN_PATTERN.exec(paragraphXml)) !== null) {
		runs.push({ full: match[0], inner: match[1] ?? '' });
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

function upsertRunProperties(runInner: string, patchXml: string): string {
	if (/<w:rPr\b/.test(runInner)) {
		return runInner.replace(/<w:rPr\b[^>]*>([\s\S]*?)<\/w:rPr>/, `<w:rPr>${patchXml}</w:rPr>`);
	}
	return `<w:rPr>${patchXml}</w:rPr>${runInner}`;
}

function buildRunPropertiesPatch(style: DocxRunStylePatch): string {
	const parts: string[] = [];
	if (style.bold === true) parts.push('<w:b/>');
	if (style.bold === false) parts.push('<w:b w:val="false"/>');
	if (style.italic === true) parts.push('<w:i/>');
	if (style.italic === false) parts.push('<w:i w:val="false"/>');
	if (style.underline === true) parts.push('<w:u w:val="single"/>');
	if (style.underline === false) parts.push('<w:u w:val="none"/>');
	if (typeof style.fontFamily === 'string' && style.fontFamily.length > 0) {
		parts.push(`<w:rFonts w:ascii="${encodeXmlText(style.fontFamily)}" w:hAnsi="${encodeXmlText(style.fontFamily)}"/>`);
	}
	if (typeof style.fontSizePt === 'number' && Number.isFinite(style.fontSizePt)) {
		const halfPoints = Math.round(style.fontSizePt * 2);
		parts.push(`<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`);
	}
	if (typeof style.color === 'string' && style.color.length > 0) {
		const hex = style.color.replace(/^#/, '').toUpperCase();
		parts.push(`<w:color w:val="${hex}"/>`);
	}
	return parts.join('');
}

function replaceRunAtIndex(paragraphXml: string, runIndex: number, nextRunXml: string): string {
	const runs = extractRuns(paragraphXml);
	if (runIndex < 0 || runIndex >= runs.length) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Run index ${runIndex} is out of range.`, { field: 'runId' });
	}

	let cursor = 0;
	for (let index = 0; index < runs.length; index++) {
		const run = runs[index];
		if (!run) continue;
		const start = paragraphXml.indexOf(run.full, cursor);
		if (start === -1) {
			throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Failed to locate run XML for patch.');
		}
		if (index === runIndex) {
			return `${paragraphXml.slice(0, start)}${nextRunXml}${paragraphXml.slice(start + run.full.length)}`;
		}
		cursor = start + run.full.length;
	}

	throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Run index ${runIndex} is out of range.`, { field: 'runId' });
}

export function patchRunText(paragraphXml: string, runIndex: number, text: string): string {
	const runs = extractRuns(paragraphXml);
	if (runIndex < 0 || runIndex >= runs.length) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Run index ${runIndex} is out of range.`, { field: 'runId' });
	}
	return replaceRunAtIndex(paragraphXml, runIndex, setRunTextContent(runs[runIndex]!.full, text));
}

export function patchRunStyle(paragraphXml: string, runIndex: number, style: DocxRunStylePatch): string {
	const runs = extractRuns(paragraphXml);
	if (runIndex < 0 || runIndex >= runs.length) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Run index ${runIndex} is out of range.`, { field: 'runId' });
	}

	const run = runs[runIndex]!;
	const patchXml = buildRunPropertiesPatch(style);
	const nextInner = upsertRunProperties(run.inner, patchXml);
	const nextRunXml = run.full.includes('/>')
		? run.full
		: run.full.replace(run.inner, nextInner);
	const patched = run.full.includes('/>')
		? run.full.replace(/\/>$/, `>${nextInner}</w:r>`)
		: nextRunXml;
	return replaceRunAtIndex(paragraphXml, runIndex, patched);
}

export function patchParagraphStyle(paragraphXml: string, style: DocxParagraphStylePatch): string {
	const styleName = typeof style.name === 'string'
		? style.name
		: typeof style.style === 'string'
			? style.style
			: null;
	if (!styleName) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Paragraph style requires a name.', { field: 'style' });
	}

	const styleXml = `<w:pStyle w:val="${encodeXmlText(styleName)}"/>`;
	if (/<w:pPr\b/.test(paragraphXml)) {
		if (/<w:pStyle\b/.test(paragraphXml)) {
			return paragraphXml.replace(/<w:pStyle\b[^>]*\/?>/, styleXml);
		}
		return paragraphXml.replace(/<w:pPr\b[^>]*>/, (match) => `${match}${styleXml}`);
	}

	return paragraphXml.replace(/^<w:p\b[^>]*>/, (match) => `${match}<w:pPr>${styleXml}</w:pPr>`);
}

export function patchCellText(cellXml: string, text: string): string {
	const encoded = encodeXmlText(text);
	const paragraphXml = `<w:p><w:r><w:t>${encoded}</w:t></w:r></w:p>`;
	if (/<w:p\b/.test(cellXml)) {
		return cellXml.replace(/<w:p\b[\s\S]*?<\/w:p>/, paragraphXml);
	}
	return cellXml.replace(/<\/w:tc>$/, `${paragraphXml}</w:tc>`);
}

export function patchCellStyle(cellXml: string, style: Record<string, unknown>): string {
	const styleName = typeof style.name === 'string'
		? style.name
		: typeof style.style === 'string'
			? style.style
			: null;
	if (!styleName) {
		return cellXml;
	}

	const paragraphXml = /<w:p\b[\s\S]*?<\/w:p>/.exec(cellXml)?.[0];
	if (!paragraphXml) {
		return cellXml.replace(
			/<\/w:tc>$/,
			`<w:p><w:pPr><w:pStyle w:val="${encodeXmlText(styleName)}"/></w:pPr><w:r><w:t></w:t></w:r></w:p></w:tc>`,
		);
	}
	const nextParagraph = patchParagraphStyle(paragraphXml, { name: styleName });
	return cellXml.replace(paragraphXml, nextParagraph);
}

export function buildEmptyTableXml(rows: number, cols: number): string {
	const cellXml = '<w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>';
	const rowXml = Array.from({ length: cols }, () => cellXml).join('');
	const rowsXml = Array.from({ length: rows }, () => `<w:tr>${rowXml}</w:tr>`).join('');
	const gridCols = Array.from({ length: cols }, () => '<w:gridCol w:w="2400"/>').join('');
	return [
		'<w:tbl>',
		'<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>',
		`<w:tblGrid>${gridCols}</w:tblGrid>`,
		rowsXml,
		'</w:tbl>',
	].join('');
}

export function replacePartText(
	partXml: string,
	query: string,
	replacement: string,
	options: { matchCase?: boolean; wholeWord?: boolean } = {},
): { partXml: string; replacementCount: number } {
	if (!query) {
		return { partXml, replacementCount: 0 };
	}

	let replacementCount = 0;
	const flags = options.matchCase ? 'g' : 'gi';
	const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	const patternSource = options.wholeWord ? `\\b${escapedQuery}\\b` : escapedQuery;
	const pattern = new RegExp(patternSource, flags);

	const nextPartXml = partXml.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g, (full, rawText: string) => {
		const decoded = decodeXmlText(rawText);
		const replaced = decoded.replace(pattern, () => {
			replacementCount++;
			return replacement;
		});
		if (replaced === decoded) {
			return full;
		}
		return full.replace(rawText, encodeXmlText(replaced));
	});

	return { partXml: nextPartXml, replacementCount };
}

export function replaceDocumentText(
	documentXml: string,
	query: string,
	replacement: string,
	options: { matchCase?: boolean; wholeWord?: boolean } = {},
): { documentXml: string; replacementCount: number } {
	const result = replacePartText(documentXml, query, replacement, options);
	return { documentXml: result.partXml, replacementCount: result.replacementCount };
}

export function buildInlineImageParagraphXml(relationshipId: string): string {
	const extentCx = 4_572_000;
	const extentCy = 3_429_000;
	return [
		'<w:p>',
		'<w:r>',
		'<w:drawing>',
		'<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">',
		`<wp:extent cx="${extentCx}" cy="${extentCy}"/>`,
		'<wp:docPr id="1" name="Picture 1"/>',
		'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
		'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">',
		'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">',
		'<pic:nvPicPr><pic:cNvPr id="0" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>',
		'<pic:blipFill>',
		`<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${relationshipId}"/>`,
		'<a:stretch><a:fillRect/></a:stretch>',
		'</pic:blipFill>',
		`<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${extentCx}" cy="${extentCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`,
		'</pic:pic>',
		'</a:graphicData>',
		'</a:graphic>',
		'</wp:inline>',
		'</w:drawing>',
		'</w:r>',
		'</w:p>',
	].join('');
}
