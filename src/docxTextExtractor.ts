import JSZip from 'jszip';
import {
	extractDocxRunText,
	extractDocxTextFromXml,
	normalizeDocxExtractedText,
} from './docxXmlText';

const TEXT_PART_PATTERNS = [
	/^word\/document\.xml$/,
	/^word\/headers\/header\d+\.xml$/,
	/^word\/footers\/footer\d+\.xml$/,
	/^word\/footnotes\.xml$/,
	/^word\/endnotes\.xml$/,
	/^word\/comments\.xml$/,
];

function isTextPart(path: string): boolean {
	return TEXT_PART_PATTERNS.some(pattern => pattern.test(path));
}

function sortTextParts(left: string, right: string): number {
	if (left === 'word/document.xml') {
		return -1;
	}
	if (right === 'word/document.xml') {
		return 1;
	}

	return left.localeCompare(right);
}

export async function extractDocxText(buffer: ArrayBuffer): Promise<string> {
	const zip = await JSZip.loadAsync(buffer.slice(0));
	const partPaths = Object.keys(zip.files)
		.filter(isTextPart)
		.sort(sortTextParts);
	const textParts: string[] = [];

	for (const partPath of partPaths) {
		const xml = await zip.file(partPath)?.async('string');
		if (!xml) {
			continue;
		}

		const text = extractDocxTextFromXml(xml);
		if (text) {
			textParts.push(text);
		}
	}

	return normalizeDocxExtractedText(textParts.join('\n\n'));
}

const PARAGRAPH_PATTERN = /<w:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:p>)/g;
const RUN_PATTERN = /<w:r\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:r>)/g;
const MARKDOWN_HARD_BREAK = ' '.repeat(2).concat('\n');

function normalizeMarkdown(value: string): string {
	return value
		.replace(/\r/g, '\n')
		// Strip accidental trailing whitespace while preserving intentional
		// two-space Markdown hard breaks.
		.replace(/([ \t]+)\n/g, (_match, whitespace: string) =>
			whitespace.endsWith('  ') ? MARKDOWN_HARD_BREAK : '\n')
		.replace(/[ \t]+$/g, '')
		.replace(/\n{3,}/g, '\n\n')
		.trim();
}

function isFlagEnabled(propertiesXml: string, tag: string): boolean {
	const match = new RegExp(`<w:${tag}\\b([^>]*)>|<w:${tag}\\b([^>]*)/>`).exec(propertiesXml);
	if (!match) {
		return false;
	}
	const attributes = `${match[1] ?? ''}${match[2] ?? ''}`;
	const valMatch = /w:val="([^"]*)"/.exec(attributes);
	if (!valMatch) {
		// A bare <w:b/> toggles the property on.
		return true;
	}
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

function applyInlineMarkdown(text: string, bold: boolean, italic: boolean): string {
	if (!bold && !italic) {
		return text;
	}
	// Don't wrap a run that is only whitespace; that would emit empty emphasis.
	if (text.trim().length === 0) {
		return text;
	}

	const marker = `${bold ? '**' : ''}${italic ? '*' : ''}`;
	// Keep surrounding whitespace outside the emphasis markers so Markdown
	// renders the emphasis correctly.
	const parts = /^(\s*)([\s\S]*?)(\s*)$/.exec(text);
	if (!parts) {
		return `${marker}${text}${marker}`;
	}
	const leading = parts[1] ?? '';
	const core = parts[2] ?? '';
	const trailing = parts[3] ?? '';
	return `${leading}${marker}${core}${marker}${trailing}`;
}

function getHeadingLevel(paragraphProperties: string): number {
	const styleMatch = /<w:pStyle\b[^>]*w:val="([^"]*)"/.exec(paragraphProperties);
	if (!styleMatch) {
		return 0;
	}
	const style = styleMatch[1] ?? '';
	if (/^title$/i.test(style)) {
		return 1;
	}
	const headingMatch = /heading\s*([1-9])/i.exec(style);
	if (!headingMatch) {
		return 0;
	}
	return Math.min(6, Number(headingMatch[1]));
}

function getListIndentLevel(paragraphProperties: string): number | null {
	if (!/<w:numPr\b/.test(paragraphProperties)) {
		return null;
	}
	const ilvlMatch = /<w:ilvl\b[^>]*w:val="([^"]*)"/.exec(paragraphProperties);
	if (!ilvlMatch) {
		return 0;
	}
	const level = Number(ilvlMatch[1]);
	return Number.isFinite(level) && level > 0 ? level : 0;
}

function paragraphToMarkdown(paragraphInnerXml: string): string {
	const propertiesXml = getParagraphProperties(paragraphInnerXml);
	let inlineText = '';
	let runMatch: RegExpExecArray | null;
	RUN_PATTERN.lastIndex = 0;

	while ((runMatch = RUN_PATTERN.exec(paragraphInnerXml)) !== null) {
		const runInner = runMatch[1];
		if (runInner === undefined) {
			continue;
		}
		const runText = extractDocxRunText(runInner);
		if (!runText) {
			continue;
		}
		const runProperties = getRunProperties(runInner);
		const bold = isFlagEnabled(runProperties, 'b');
		const italic = isFlagEnabled(runProperties, 'i');
		inlineText += applyInlineMarkdown(runText, bold, italic);
	}

	// Soft line breaks inside a paragraph become Markdown hard breaks.
	inlineText = inlineText.replace(/\n/g, MARKDOWN_HARD_BREAK).trim();
	if (!inlineText) {
		return '';
	}

	const headingLevel = getHeadingLevel(propertiesXml);
	if (headingLevel > 0) {
		return `${'#'.repeat(headingLevel)} ${inlineText}`;
	}

	const listLevel = getListIndentLevel(propertiesXml);
	if (listLevel !== null) {
		return `${'  '.repeat(listLevel)}- ${inlineText}`;
	}

	return inlineText;
}

function extractMarkdownFromXml(xml: string): string {
	const blocks: string[] = [];
	let match: RegExpExecArray | null;
	PARAGRAPH_PATTERN.lastIndex = 0;

	while ((match = PARAGRAPH_PATTERN.exec(xml)) !== null) {
		const inner = match[1];
		if (inner === undefined) {
			continue;
		}
		const markdown = paragraphToMarkdown(inner);
		if (markdown) {
			blocks.push(markdown);
		}
	}

	return normalizeMarkdown(blocks.join('\n\n'));
}

/**
 * Serializes a DOCX into Markdown. Paragraph styles (headings, list items) and
 * run-level bold/italic that are present in the OOXML are mapped to Markdown.
 * Anything richer than that (tables, images, links, colors) is flattened to its
 * text content, so the output is a faithful-but-minimal Markdown rendering.
 */
export async function extractDocxMarkdown(buffer: ArrayBuffer): Promise<string> {
	const zip = await JSZip.loadAsync(buffer.slice(0));
	const partPaths = Object.keys(zip.files)
		.filter(isTextPart)
		.sort(sortTextParts);
	const markdownParts: string[] = [];

	for (const partPath of partPaths) {
		const xml = await zip.file(partPath)?.async('string');
		if (!xml) {
			continue;
		}

		const markdown = extractMarkdownFromXml(xml);
		if (markdown) {
			markdownParts.push(markdown);
		}
	}

	return normalizeMarkdown(markdownParts.join('\n\n'));
}
