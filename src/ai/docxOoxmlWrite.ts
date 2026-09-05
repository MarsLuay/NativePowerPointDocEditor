import { AI_ERROR_CODES, createAiError } from './errors';
import { extractDocxRunText } from '../docxXmlText';

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

function getChildElementXml(propertiesXml: string, tag: string): string | null {
	return propertiesXml.match(new RegExp(`<w:${tag}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/w:${tag}>)`))?.[0] ?? null;
}

function upsertChildElement(propertiesXml: string, tag: string, replacement: string | null): string {
	const pattern = new RegExp(`<w:${tag}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/w:${tag}>)`);
	if (pattern.test(propertiesXml)) {
		return propertiesXml.replace(pattern, replacement ?? '');
	}
	return replacement === null ? propertiesXml : `${propertiesXml}${replacement}`;
}

function patchElementAttributes(
	existingXml: string | null,
	tag: string,
	patch: Record<string, string>,
): string {
	const attributes = new Map<string, string>();
	if (existingXml) {
		for (const match of existingXml.matchAll(/([\w:-]+)="([^"]*)"/g)) {
			const name = match[1];
			const value = match[2];
			if (name && value !== undefined) attributes.set(name, value);
		}
	}
	for (const [name, value] of Object.entries(patch)) {
		attributes.set(`w:${name}`, value);
	}
	const serialized = [...attributes.entries()]
		.map(([name, value]) => `${name}="${encodeXmlText(value)}"`)
		.join(' ');
	return `<w:${tag}${serialized ? ` ${serialized}` : ''}/>`;
}

function patchRunProperties(propertiesXml: string, style: DocxRunStylePatch): string {
	let next = propertiesXml;
	if (style.bold !== undefined) {
		next = upsertChildElement(next, 'b', style.bold ? '<w:b/>' : '<w:b w:val="false"/>');
	}
	if (style.italic !== undefined) {
		next = upsertChildElement(next, 'i', style.italic ? '<w:i/>' : '<w:i w:val="false"/>');
	}
	if (style.underline !== undefined) {
		next = upsertChildElement(next, 'u', `<w:u w:val="${style.underline ? 'single' : 'none'}"/>`);
	}
	if (style.fontFamily !== undefined) {
		const fontFamily = style.fontFamily;
		next = upsertChildElement(
			next,
			'rFonts',
			patchElementAttributes(getChildElementXml(next, 'rFonts'), 'rFonts', {
				ascii: fontFamily,
				hAnsi: fontFamily,
				cs: fontFamily,
			}),
		);
	}
	if (style.fontSizePt !== undefined) {
		const halfPoints = String(Math.round(style.fontSizePt * 2));
		for (const tag of ['sz', 'szCs']) {
			next = upsertChildElement(
				next,
				tag,
				patchElementAttributes(getChildElementXml(next, tag), tag, { val: halfPoints }),
			);
		}
	}
	if (style.color !== undefined) {
		const color = style.color;
		next = upsertChildElement(
			next,
			'color',
			color === null
				? null
				: patchElementAttributes(getChildElementXml(next, 'color'), 'color', {
					val: color.replace(/^#/, '').toUpperCase(),
				}),
		);
	}
	return next;
}

function patchRunPropertiesContainer(
	containerXml: string,
	style: DocxRunStylePatch,
	placement: 'prepend' | 'append' = 'prepend',
): string {
	const paired = /<w:rPr\b([^>]*)>([\s\S]*?)<\/w:rPr>/.exec(containerXml);
	if (paired) {
		const [full, attributes = '', properties = ''] = paired;
		return containerXml.replace(full, `<w:rPr${attributes}>${patchRunProperties(properties, style)}</w:rPr>`);
	}
	if (/<w:rPr\b[^>]*\/>/.test(containerXml)) {
		return containerXml.replace(/<w:rPr\b([^>]*)\/>/, (_match, attributes: string) => (
			`<w:rPr${attributes}>${patchRunProperties('', style)}</w:rPr>`
		));
	}
	const properties = `<w:rPr>${patchRunProperties('', style)}</w:rPr>`;
	return placement === 'prepend' ? `${properties}${containerXml}` : `${containerXml}${properties}`;
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

export function getRunText(paragraphXml: string, runIndex: number): string {
	const runs = extractRuns(paragraphXml);
	if (runIndex < 0 || runIndex >= runs.length) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Run index ${runIndex} is out of range.`, { field: 'runId' });
	}
	return extractDocxRunText(runs[runIndex]!.full);
}

export function patchRunStyle(paragraphXml: string, runIndex: number, style: DocxRunStylePatch): string {
	const runs = extractRuns(paragraphXml);
	if (runIndex < 0 || runIndex >= runs.length) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Run index ${runIndex} is out of range.`, { field: 'runId' });
	}

	const run = runs[runIndex]!;
	const nextInner = patchRunPropertiesContainer(run.inner, style);
	const isSelfClosingRun = /^<w:r\b[^>]*\/>$/.test(run.full.trim());
	const nextRunXml = isSelfClosingRun
		? run.full
		: run.full.replace(run.inner, nextInner);
	const patched = isSelfClosingRun
		? run.full.replace(/\/>$/, `>${nextInner}</w:r>`)
		: nextRunXml;
	return replaceRunAtIndex(paragraphXml, runIndex, patched);
}

/** Patch w:pPr/w:rPr, including empty paragraphs, without replacing other paragraph or run properties. */
export function patchParagraphDefaultRunStyle(
	paragraphXml: string,
	style: DocxRunStylePatch,
): string {
	const match = /^(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)$/.exec(paragraphXml);
	if (!match) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Failed to parse paragraph XML.');
	}
	const [, opening, body = '', closing] = match;
	const pairedProperties = /^(<w:pPr\b[^>]*>)([\s\S]*?)(<\/w:pPr>)([\s\S]*)$/.exec(body);
	if (pairedProperties) {
		const [, propertiesOpening, propertiesBody = '', propertiesClosing, remainder = ''] = pairedProperties;
		return `${opening}${propertiesOpening}${patchRunPropertiesContainer(propertiesBody, style, 'append')}${propertiesClosing}${remainder}${closing}`;
	}
	const selfClosingProperties = /^(<w:pPr\b([^>]*)\/>)([\s\S]*)$/.exec(body);
	if (selfClosingProperties) {
		const [, , attributes = '', remainder = ''] = selfClosingProperties;
		return `${opening}<w:pPr${attributes}>${patchRunPropertiesContainer('', style, 'append')}</w:pPr>${remainder}${closing}`;
	}
	return `${opening}<w:pPr>${patchRunPropertiesContainer('', style, 'append')}</w:pPr>${body}${closing}`;
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

export interface DocxParagraphBottomBorderPatch {
	style: string;
	size?: number;
	space?: number;
	color?: string;
}

/**
 * Set only w:pBdr/w:bottom while preserving every other paragraph property
 * and border side. Useful for heading rules without recreating a paragraph.
 */
export function patchParagraphBottomBorder(
	paragraphXml: string,
	border: DocxParagraphBottomBorderPatch,
): string {
	const attrs = [
		`w:val="${encodeXmlText(border.style)}"`,
		...(typeof border.size === 'number' ? [`w:sz="${border.size}"`] : []),
		...(typeof border.space === 'number' ? [`w:space="${border.space}"`] : []),
		...(border.color ? [`w:color="${encodeXmlText(border.color)}"`] : []),
	].join(' ');
	const bottomXml = `<w:bottom ${attrs}/>`;

	if (/<w:pBdr\b/.test(paragraphXml)) {
		if (/<w:bottom\b[^>]*\/>/.test(paragraphXml)) {
			return paragraphXml.replace(/<w:bottom\b[^>]*\/>/, bottomXml);
		}
		return paragraphXml.replace(/<\/w:pBdr>/, `${bottomXml}</w:pBdr>`);
	}
	if (/<w:pPr\b/.test(paragraphXml)) {
		return paragraphXml.replace(/<w:pPr\b[^>]*>/, (match) => `${match}<w:pBdr>${bottomXml}</w:pBdr>`);
	}
	return paragraphXml.replace(
		/^<w:p\b[^>]*>/,
		(match) => `${match}<w:pPr><w:pBdr>${bottomXml}</w:pBdr></w:pPr>`,
	);
}

export function patchCellText(cellXml: string, text: string): string {
	const encoded = encodeXmlText(text);
	const paragraphXml = `<w:p><w:r><w:t>${encoded}</w:t></w:r></w:p>`;
	const openTag = /^<w:tc\b[^>]*>/.exec(cellXml)?.[0];
	const closeIndex = cellXml.lastIndexOf('</w:tc>');
	if (!openTag || closeIndex === -1) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Failed to parse table cell XML.');
	}

	// A cell can contain many paragraphs, drawings, and nested tables. Replacing
	// only its first paragraph leaves stale template content behind. Preserve the
	// cell properties while replacing the complete cell body with one paragraph.
	const inner = cellXml.slice(openTag.length, closeIndex);
	const properties = /^\s*(<w:tcPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:tcPr>))/.exec(inner)?.[1] ?? '';
	return `${openTag}${properties}${paragraphXml}</w:tc>`;
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

	const nextPartXml = partXml.replace(/(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g, (full, openTag: string, rawText: string, closeTag: string) => {
		const decoded = decodeXmlText(rawText);
		const replaced = decoded.replace(pattern, () => {
			replacementCount++;
			return replacement;
		});
		if (replaced === decoded) {
			return full;
		}
		return openTag + encodeXmlText(replaced) + closeTag;
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
