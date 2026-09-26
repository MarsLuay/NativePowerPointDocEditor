import { AI_ERROR_CODES, createAiError } from './errors';

export type DocxParagraphAlignment = 'left' | 'center' | 'right' | 'both' | 'distribute';
export type DocxLineRule = 'auto' | 'exact' | 'atLeast';
export type DocxTabAlignment = 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear';

export interface DocxParagraphSpacing {
	before?: number;
	after?: number;
	line?: number;
	lineRule?: DocxLineRule;
}

export interface DocxParagraphIndent {
	left?: number;
	right?: number;
	firstLine?: number;
	hanging?: number;
}

export interface DocxParagraphTab {
	val: DocxTabAlignment;
	pos: number;
}

export interface DocxParagraphLayout {
	alignment?: DocxParagraphAlignment;
	spacing?: DocxParagraphSpacing;
	indent?: DocxParagraphIndent;
	tabs?: DocxParagraphTab[];
	keepNext?: boolean;
	keepLines?: boolean;
	pageBreakBefore?: boolean;
}

export interface DocxSectionPageSize {
	width: number;
	height: number;
	orient?: 'portrait' | 'landscape';
}

export interface DocxSectionMargins {
	top?: number;
	right?: number;
	bottom?: number;
	left?: number;
	header?: number;
	footer?: number;
	gutter?: number;
}

export interface DocxSectionLayout {
	index: number;
	pageSize?: DocxSectionPageSize;
	margins?: DocxSectionMargins;
}

export interface DocxSectionLayoutPatch {
	pageSize?: DocxSectionPageSize;
	margins?: DocxSectionMargins;
}

function getAttribute(xml: string, name: string): string | null {
	return xml.match(new RegExp(`(?:w:)?${name}="([^"]*)"`))?.[1] ?? null;
}

function parseInteger(value: string | null): number | undefined {
	if (value === null) return undefined;
	const parsed = Number(value);
	return Number.isInteger(parsed) ? parsed : undefined;
}

function parseBooleanTag(propertiesXml: string, tag: string): boolean | undefined {
	const match = propertiesXml.match(new RegExp(`<w:${tag}\\b([^>]*)\\/?>`));
	if (!match) return undefined;
	const value = getAttribute(match[1] ?? '', 'val');
	return value === null || !['0', 'false', 'off'].includes(value.toLowerCase());
}

function getParagraphProperties(paragraphXml: string): string {
	return paragraphXml.match(/<w:pPr\b[^>]*>([\s\S]*?)<\/w:pPr>/)?.[1] ?? '';
}

function getChildXml(propertiesXml: string, tag: string): string | null {
	return propertiesXml.match(new RegExp(`<w:${tag}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/w:${tag}>)`))?.[0] ?? null;
}

function parseAttributes(xml: string, names: readonly string[]): Record<string, number> {
	const result: Record<string, number> = {};
	for (const name of names) {
		const value = parseInteger(getAttribute(xml, name));
		if (value !== undefined) result[name] = value;
	}
	return result;
}

export function parseParagraphLayout(paragraphXml: string): DocxParagraphLayout {
	const propertiesXml = getParagraphProperties(paragraphXml);
	const layout: DocxParagraphLayout = {};
	const alignment = getAttribute(getChildXml(propertiesXml, 'jc') ?? '', 'val');
	if (alignment && ['left', 'center', 'right', 'both', 'distribute'].includes(alignment)) {
		layout.alignment = alignment as DocxParagraphAlignment;
	}

	const spacingXml = getChildXml(propertiesXml, 'spacing');
	if (spacingXml) {
		const spacing: DocxParagraphSpacing = parseAttributes(spacingXml, ['before', 'after', 'line']);
		const lineRule = getAttribute(spacingXml, 'lineRule');
		if (lineRule === 'auto' || lineRule === 'exact' || lineRule === 'atLeast') spacing.lineRule = lineRule;
		if (Object.keys(spacing).length > 0) layout.spacing = spacing;
	}

	const indentXml = getChildXml(propertiesXml, 'ind');
	if (indentXml) {
		const indent = parseAttributes(indentXml, ['left', 'right', 'firstLine', 'hanging']);
		if (Object.keys(indent).length > 0) layout.indent = indent;
	}

	const tabsXml = getChildXml(propertiesXml, 'tabs');
	if (tabsXml) {
		layout.tabs = [...tabsXml.matchAll(/<w:tab\b[^>]*\/?>(?:)/g)].flatMap((match) => {
			const val = getAttribute(match[0], 'val');
			const pos = parseInteger(getAttribute(match[0], 'pos'));
			return val && pos !== undefined && ['left', 'center', 'right', 'decimal', 'bar', 'clear'].includes(val)
				? [{ val: val as DocxTabAlignment, pos }]
				: [];
		});
	}

	for (const tag of ['keepNext', 'keepLines', 'pageBreakBefore'] as const) {
		const value = parseBooleanTag(propertiesXml, tag);
		if (value !== undefined) layout[tag] = value;
	}

	return layout;
}

function encodeXmlText(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function upsertChild(propertiesXml: string, tag: string, replacement: string): string {
	const pattern = new RegExp(`<w:${tag}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/w:${tag}>)`);
	if (pattern.test(propertiesXml)) return propertiesXml.replace(pattern, replacement);
	return `${propertiesXml}${replacement}`;
}

function ensureParagraphProperties(paragraphXml: string): { prefix: string; properties: string; suffix: string } {
	const match = /^(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)$/.exec(paragraphXml);
	if (!match) throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Failed to parse paragraph XML.');
	const [, opening, body, closing] = match;
	if (!opening || body === undefined || !closing) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Failed to parse paragraph XML parts.');
	}
	const propertiesMatch = /^(<w:pPr\b[^>]*>)([\s\S]*?)(<\/w:pPr>)([\s\S]*)$/.exec(body);
	if (propertiesMatch) {
		const [, propertiesOpening, propertiesBody, propertiesClosing, remainder] = propertiesMatch;
		if (!propertiesOpening || propertiesBody === undefined || !propertiesClosing || remainder === undefined) {
			throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Failed to parse paragraph properties.');
		}
		return {
			prefix: `${opening}${propertiesOpening}`,
			properties: propertiesBody,
			suffix: `${propertiesClosing}${remainder}${closing}`,
		};
	}
	const selfClosingPropertiesMatch = /^(<w:pPr\b[^>]*)\/>([\s\S]*)$/.exec(body);
	if (selfClosingPropertiesMatch) {
		const [, propertiesOpening, remainder] = selfClosingPropertiesMatch;
		if (!propertiesOpening || remainder === undefined) {
			throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Failed to parse self-closing paragraph properties.');
		}
		return {
			prefix: `${opening}${propertiesOpening}>`,
			properties: '',
			suffix: `</w:pPr>${remainder}${closing}`,
		};
	}
	return {
		prefix: `${opening}<w:pPr>`,
		properties: '',
		suffix: `</w:pPr>${body}${closing}`,
	};
}

function patchAttributes(existingXml: string | null, tag: string, patch: Record<string, string | number | undefined>): string {
	const openingAttributes = existingXml?.match(new RegExp(`^<w:${tag}\\b([^>]*)`))?.[1] ?? '';
	const attributes = [...openingAttributes.matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*"([^"]*)"/g)].map((match) => ({
		name: match[1]!,
		value: match[2]!,
	}));
	for (const [name, value] of Object.entries(patch)) {
		if (value === undefined) continue;
		const encoded = encodeXmlText(String(value));
		const existing = attributes.find((attribute) => attribute.name === `w:${name}` || attribute.name === name);
		if (existing) existing.value = encoded;
		else attributes.push({ name: `w:${name}`, value: encoded });
	}
	const serialized = attributes.map(({ name, value }) => `${name}="${value}"`).join(' ');
	return `<w:${tag}${serialized ? ` ${serialized}` : ''}/>`;
}

export function patchParagraphLayout(paragraphXml: string, patch: DocxParagraphLayout): string {
	const parts = ensureParagraphProperties(paragraphXml);
	let properties = parts.properties;

	if (patch.alignment !== undefined) {
		properties = upsertChild(properties, 'jc', `<w:jc w:val="${patch.alignment}"/>`);
	}
	if (patch.spacing !== undefined) {
		const existing = getChildXml(properties, 'spacing');
		properties = upsertChild(properties, 'spacing', patchAttributes(existing, 'spacing', { ...patch.spacing }));
	}
	if (patch.indent !== undefined) {
		const existing = getChildXml(properties, 'ind');
		properties = upsertChild(properties, 'ind', patchAttributes(existing, 'ind', { ...patch.indent }));
	}
	if (patch.tabs !== undefined) {
		const tabs = patch.tabs.map((tab) => `<w:tab w:val="${tab.val}" w:pos="${tab.pos}"/>`).join('');
		properties = upsertChild(properties, 'tabs', `<w:tabs>${tabs}</w:tabs>`);
	}
	for (const tag of ['keepNext', 'keepLines', 'pageBreakBefore'] as const) {
		if (patch[tag] !== undefined) {
			properties = upsertChild(properties, tag, patch[tag] ? `<w:${tag}/>` : `<w:${tag} w:val="0"/>`);
		}
	}

	return `${parts.prefix}${properties}${parts.suffix}`;
}

function sectionXmls(documentXml: string): Array<{ full: string; start: number }> {
	const sections: Array<{ full: string; start: number }> = [];
	const pattern = /<w:sectPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:sectPr>)/g;
	let match: RegExpExecArray | null;
	while ((match = pattern.exec(documentXml)) !== null) {
		sections.push({ full: match[0], start: match.index });
	}
	return sections;
}

function parseSectionLayout(sectionXml: string, index: number): DocxSectionLayout {
	const pageSizeXml = getChildXml(sectionXml, 'pgSz');
	const marginsXml = getChildXml(sectionXml, 'pgMar');
	const pageSize = pageSizeXml
		? {
			width: parseInteger(getAttribute(pageSizeXml, 'w')) ?? 0,
			height: parseInteger(getAttribute(pageSizeXml, 'h')) ?? 0,
			...(getAttribute(pageSizeXml, 'orient') ? { orient: getAttribute(pageSizeXml, 'orient') as 'portrait' | 'landscape' } : {}),
		}
		: undefined;
	const margins = marginsXml
		? parseAttributes(marginsXml, ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'])
		: undefined;
	return {
		index,
		...(pageSize && (pageSize.width > 0 || pageSize.height > 0) ? { pageSize } : {}),
		...(margins && Object.keys(margins).length > 0 ? { margins } : {}),
	};
}

export function parseDocumentSectionLayouts(documentXml: string): DocxSectionLayout[] {
	return sectionXmls(documentXml).map(({ full }, index) => parseSectionLayout(full, index));
}

export function patchDocumentSectionLayout(
	documentXml: string,
	sectionIndex: number,
	patch: DocxSectionLayoutPatch,
): string {
	const sections = sectionXmls(documentXml);
	const section = sections[sectionIndex];
	if (!section) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Section ${sectionIndex} was not found.`, { field: 'sectionIndex' });
	}

	const expandedSection = /^(<w:sectPr\b[^>]*>)([\s\S]*)(<\/w:sectPr>)$/.exec(section.full);
	const selfClosingSection = /^(<w:sectPr\b[^>]*)\/>$/.exec(section.full);
	const sectionOpening = expandedSection?.[1] ?? (selfClosingSection?.[1] ? `${selfClosingSection[1]}>` : null);
	let sectionBody = expandedSection?.[2] ?? (selfClosingSection ? '' : null);
	if (!sectionOpening || sectionBody === null) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, `Failed to parse section ${sectionIndex}.`, {
			field: 'sectionIndex',
		});
	}
	if (patch.pageSize !== undefined) {
		const existing = getChildXml(sectionBody, 'pgSz');
		const pageSize = patchAttributes(existing, 'pgSz', {
			w: patch.pageSize.width,
			h: patch.pageSize.height,
			orient: patch.pageSize.orient,
		});
		sectionBody = upsertChild(sectionBody, 'pgSz', pageSize);
	}
	if (patch.margins !== undefined) {
		const existing = getChildXml(sectionBody, 'pgMar');
		const margins = patchAttributes(existing, 'pgMar', { ...patch.margins });
		sectionBody = upsertChild(sectionBody, 'pgMar', margins);
	}
	const nextSection = `${sectionOpening}${sectionBody}</w:sectPr>`;

	return `${documentXml.slice(0, section.start)}${nextSection}${documentXml.slice(section.start + section.full.length)}`;
}
