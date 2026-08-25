import JSZip from 'jszip';
import type { App, TFile } from 'obsidian';

import { DEFAULT_DOCX_STYLES_XML } from '../docxStyleDefaults';
import { getAvailableNumberedPath } from '../export/artifactPaths';
import { errorLog, infoLog } from '../logger';

const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_RELS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

interface InlineFormatting {
	bold?: boolean;
	italic?: boolean;
	strike?: boolean;
	code?: boolean;
}

type MarkdownBlock =
	| { kind: 'paragraph'; text: string }
	| { kind: 'heading'; level: number; text: string }
	| { kind: 'list'; level: number; ordered: boolean; text: string }
	| { kind: 'quote'; text: string }
	| { kind: 'code'; lines: string[] }
	| { kind: 'rule' };

const INLINE_TOKEN_PATTERN = /(?<!\\)(`+)([\s\S]*?)\1|(?<!\\)!\[([^\]]*)\]\(([^)\n]+)\)|(?<!\\)\[([^\]]+)\]\(([^)\n]+)\)|(?<!\\)!?\[\[([^\]]+)\]\]|(?<!\\)\*\*([^*\n]+)\*\*|(?<!\\)__([^_\n]+)__|(?<!\\)~~([^~\n]+)~~|(?<![\\*])\*([^*\n]+)\*(?!\*)|(?<![\\\w])_([^_\n]+)_(?!\w)|\n/g;

function stripInvalidXmlControlCharacters(value: string): string {
	let sanitized = '';
	for (const character of value) {
		const codePoint = character.codePointAt(0) ?? 0;
		if (codePoint <= 0x08 || codePoint === 0x0B || codePoint === 0x0C || (codePoint >= 0x0E && codePoint <= 0x1F)) {
			continue;
		}
		sanitized += character;
	}
	return sanitized;
}

function escapeXml(value: string): string {
	return stripInvalidXmlControlCharacters(value)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function removeMarkdownEscapes(value: string): string {
	return value.replace(/\\([\\`*_[\]{}()#+\-.!>])/g, '$1');
}

function renderRun(text: string, formatting: InlineFormatting = {}): string {
	const properties = [
		formatting.bold ? '<w:b/><w:bCs/>' : '',
		formatting.italic ? '<w:i/><w:iCs/>' : '',
		formatting.strike ? '<w:strike/>' : '',
		formatting.code
			? '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:shd w:val="clear" w:color="auto" w:fill="EDEDED"/>'
			: '',
	].join('');
	return `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

function readableLinkTarget(rawTarget: string): string {
	const trimmed = rawTarget.trim();
	const angleWrapped = trimmed.match(/^<([^>]+)>/);
	if (angleWrapped?.[1]) {
		return angleWrapped[1];
	}
	return trimmed.match(/^\S+/)?.[0] ?? trimmed;
}

function renderInline(markdown: string, formatting: InlineFormatting = {}): string {
	let output = '';
	let cursor = 0;
	INLINE_TOKEN_PATTERN.lastIndex = 0;

	for (const match of markdown.matchAll(INLINE_TOKEN_PATTERN)) {
		const index = match.index ?? 0;
		if (index > cursor) {
			output += renderRun(removeMarkdownEscapes(markdown.slice(cursor, index)), formatting);
		}

		if (match[0] === '\n') {
			output += '<w:r><w:br/></w:r>';
		} else if (match[2] !== undefined) {
			output += renderRun(match[2], { ...formatting, code: true });
		} else if (match[3] !== undefined) {
			const target = readableLinkTarget(match[4] ?? '');
			output += renderRun(match[3] || target, formatting);
		} else if (match[5] !== undefined) {
			const target = readableLinkTarget(match[6] ?? '');
			output += renderInline(match[5], formatting);
			if (target && target !== match[5]) {
				output += renderRun(` (${target})`, formatting);
			}
		} else if (match[7] !== undefined) {
			const [target, alias] = match[7].split('|', 2);
			output += renderRun(alias || target || '', formatting);
		} else if (match[8] !== undefined || match[9] !== undefined) {
			output += renderInline(match[8] ?? match[9] ?? '', { ...formatting, bold: true });
		} else if (match[10] !== undefined) {
			output += renderInline(match[10], { ...formatting, strike: true });
		} else if (match[11] !== undefined || match[12] !== undefined) {
			output += renderInline(match[11] ?? match[12] ?? '', { ...formatting, italic: true });
		}
		cursor = index + match[0].length;
	}

	if (cursor < markdown.length) {
		output += renderRun(removeMarkdownEscapes(markdown.slice(cursor)), formatting);
	}
	return output || renderRun('');
}

function stripYamlFrontmatter(lines: string[]): string[] {
	if (lines[0]?.trim() !== '---') {
		return lines;
	}
	const end = lines.findIndex((line, index) => index > 0 && /^(---|\.\.\.)\s*$/.test(line));
	if (end < 2 || !lines.slice(1, end).some(line => /^[\w.-]+\s*:/.test(line))) {
		return lines;
	}
	return lines.slice(end + 1);
}

function isFence(line: string): RegExpMatchArray | null {
	return line.match(/^\s*(`{3,}|~{3,})/);
}

function isRule(line: string): boolean {
	return /^\s{0,3}(?:(?:\*\s*){3,}|(?:-\s*){3,}|(?:_\s*){3,})$/.test(line);
}

function joinParagraphLines(lines: string[]): string {
	return lines.map((line, index) => {
		const hardBreak = /(?: {2,}|\\)$/.test(line);
		const text = hardBreak ? line.replace(/(?: {2,}|\\)$/, '') : line.trim();
		return index < lines.length - 1 ? `${text}${hardBreak ? '\n' : ' '}` : text;
	}).join('');
}

function parseCodeBlock(lines: string[], startIndex: number, fenceMatch: RegExpMatchArray): { block: MarkdownBlock; nextIndex: number } {
	const fenceToken = fenceMatch[1] ?? '```';
	const delimiter = fenceToken[0] ?? '`';
	const minimumLength = fenceToken.length;
	const codeLines: string[] = [];
	let index = startIndex + 1;
	const closingFence = new RegExp(`^\\s*${delimiter}{${minimumLength},}\\s*$`);
	while (index < lines.length) {
		const codeLine = lines[index];
		if (codeLine === undefined || closingFence.test(codeLine)) {
			break;
		}
		codeLines.push(codeLine);
		index += 1;
	}
	if (index < lines.length) index += 1;
	return { block: { kind: 'code', lines: codeLines.length ? codeLines : [''] }, nextIndex: index };
}

function parseListBlock(lines: string[], startIndex: number, listMatch: RegExpMatchArray): { block: MarkdownBlock; nextIndex: number } {
	const continuation: string[] = [listMatch[4] ?? ''];
	const indentation = (listMatch[1] ?? '').replace(/\t/g, '    ').length;
	let index = startIndex + 1;
	while (index < lines.length) {
		const continuationLine = lines[index];
		if (continuationLine === undefined || !/^\s+\S/.test(continuationLine) || /^(\s*)(?:[-+*]|\d+[.)])\s+/.test(continuationLine)) {
			break;
		}
		continuation.push(continuationLine.trim());
		index += 1;
	}
	return {
		block: {
			kind: 'list',
			level: Math.min(8, Math.floor(indentation / 2)),
			ordered: listMatch[3] !== undefined,
			text: joinParagraphLines(continuation).replace(/^\[ \]\s+/, '☐ ').replace(/^\[[xX]\]\s+/, '☒ '),
		},
		nextIndex: index,
	};
}

function parseQuoteBlock(lines: string[], startIndex: number): { block: MarkdownBlock; nextIndex: number } {
	const quoteLines: string[] = [];
	let index = startIndex;
	while (index < lines.length) {
		const quoteLine = lines[index];
		if (quoteLine === undefined || !/^\s{0,3}>/.test(quoteLine)) {
			break;
		}
		quoteLines.push(quoteLine.replace(/^\s{0,3}>\s?/, ''));
		index += 1;
	}
	return { block: { kind: 'quote', text: joinParagraphLines(quoteLines) }, nextIndex: index };
}

function parseParagraphBlock(lines: string[], startIndex: number): { block: MarkdownBlock; nextIndex: number } {
	const line = lines[startIndex];
	const paragraphLines = line !== undefined ? [line] : [];
	let index = startIndex + 1;
	while (index < lines.length) {
		const next = lines[index];
		if (next === undefined) {
			break;
		}
		if (!next.trim() || isFence(next) || /^\s{0,3}(?:#{1,6}\s+|>|(?:[-+*]|\d+[.)])\s+)/.test(next) || isRule(next)) {
			break;
		}
		const following = lines[index + 1];
		if (following !== undefined && /^\s{0,3}(?:=+|-+)\s*$/.test(following)) {
			break;
		}
		paragraphLines.push(next);
		index += 1;
	}
	return { block: { kind: 'paragraph', text: joinParagraphLines(paragraphLines) }, nextIndex: index };
}

function parseMarkdownBlocks(markdown: string): MarkdownBlock[] {
	const lines = stripYamlFrontmatter(markdown.replace(/\r\n?/g, '\n').split('\n'));
	const blocks: MarkdownBlock[] = [];

	for (let index = 0; index < lines.length;) {
		const line = lines[index];
		if (line === undefined) {
			break;
		}
		if (!line.trim()) {
			index += 1;
			continue;
		}

		const fence = isFence(line);
		if (fence) {
			const { block, nextIndex } = parseCodeBlock(lines, index, fence);
			blocks.push(block);
			index = nextIndex;
			continue;
		}

		const heading = line.match(/^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$/);
		if (heading) {
			blocks.push({ kind: 'heading', level: (heading[1] ?? '#').length, text: heading[2] ?? '' });
			index += 1;
			continue;
		}

		const setext = lines[index + 1]?.match(/^\s{0,3}(=+|-+)\s*$/);
		if (setext) {
			blocks.push({ kind: 'heading', level: setext[1]?.[0] === '=' ? 1 : 2, text: line.trim() });
			index += 2;
			continue;
		}

		const list = line.match(/^(\s*)(?:([-+*])|(\d+)[.)])\s+(.+)$/);
		if (list) {
			const { block, nextIndex } = parseListBlock(lines, index, list);
			blocks.push(block);
			index = nextIndex;
			continue;
		}

		if (/^\s{0,3}>/.test(line)) {
			const { block, nextIndex } = parseQuoteBlock(lines, index);
			blocks.push(block);
			index = nextIndex;
			continue;
		}

		if (isRule(line)) {
			blocks.push({ kind: 'rule' });
			index += 1;
			continue;
		}

		const { block, nextIndex } = parseParagraphBlock(lines, index);
		blocks.push(block);
		index = nextIndex;
	}

	return blocks;
}

function renderBlock(block: MarkdownBlock): string {
	if (block.kind === 'heading') {
		const headingProperties = block.level <= 3
			? `<w:pStyle w:val="Heading${block.level}"/>`
			: `<w:keepNext/><w:keepLines/><w:spacing w:before="120" w:after="0"/><w:outlineLvl w:val="${block.level - 1}"/>`;
		const inlineFormatting = block.level > 3 ? { bold: true } : {};
		return `<w:p><w:pPr>${headingProperties}</w:pPr>${renderInline(block.text, inlineFormatting)}</w:p>`;
	}
	if (block.kind === 'list') {
		return `<w:p><w:pPr><w:numPr><w:ilvl w:val="${block.level}"/><w:numId w:val="${block.ordered ? 2 : 1}"/></w:numPr></w:pPr>${renderInline(block.text)}</w:p>`;
	}
	if (block.kind === 'quote') {
		return `<w:p><w:pPr><w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="B7B7B7"/></w:pBdr><w:ind w:left="480"/></w:pPr>${renderInline(block.text, { italic: true })}</w:p>`;
	}
	if (block.kind === 'code') {
		return block.lines.map(line => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:left="240"/><w:shd w:val="clear" w:color="auto" w:fill="F3F3F3"/></w:pPr>${renderRun(line, { code: true })}</w:p>`).join('\n');
	}
	if (block.kind === 'rule') {
		return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="B7B7B7"/></w:pBdr><w:spacing w:after="120"/></w:pPr></w:p>';
	}
	return `<w:p>${renderInline(block.text)}</w:p>`;
}

function buildNumberingXml(): string {
	const levels = (format: 'bullet' | 'decimal') => Array.from({ length: 9 }, (_, level) => {
		const left = 720 + (level * 360);
		const text = format === 'bullet' ? '•' : `%${level + 1}.`;
		return `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="${left}"/></w:tabs><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;
	}).join('');

	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${WORD_NS}">
  <w:abstractNum w:abstractNumId="0">${levels('bullet')}</w:abstractNum>
  <w:abstractNum w:abstractNumId="1">${levels('decimal')}</w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;
}

export function buildMarkdownDocxCandidatePath(sourcePath: string): string {
	if (!/\.md$/i.test(sourcePath)) {
		throw new Error('Markdown source path must end in .md');
	}
	return `${sourcePath.slice(0, -3)}.docx`;
}

export function resolveMarkdownDocxOutputPath(sourcePath: string, exists: (path: string) => boolean): string {
	const candidatePath = buildMarkdownDocxCandidatePath(sourcePath);
	return exists(candidatePath) ? getAvailableNumberedPath(candidatePath, exists) : candidatePath;
}

export async function buildMarkdownDocxArrayBuffer(markdown: string): Promise<ArrayBuffer> {
	const blocks = parseMarkdownBlocks(markdown);
	const body = (blocks.length ? blocks : [{ kind: 'paragraph', text: '' } satisfies MarkdownBlock])
		.map(renderBlock)
		.join('\n');
	const zip = new JSZip();

	zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`);
	zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`);
	zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_RELS_NS}">
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${OFFICE_RELS_NS}/numbering" Target="numbering.xml"/>
</Relationships>`);
	zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${WORD_NS}">
  <w:body>
    ${body}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`);
	zip.file('word/styles.xml', DEFAULT_DOCX_STYLES_XML);
	zip.file('word/numbering.xml', buildNumberingXml());

	return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

/** Convert one vault Markdown file into a collision-safe DOCX sibling and open it. */
export async function convertMarkdownFileToDocx(app: App, sourceFile: TFile): Promise<TFile> {
	const outputPath = resolveMarkdownDocxOutputPath(
		sourceFile.path,
		path => app.vault.getAbstractFileByPath(path) != null,
	);
	const markdown = await app.vault.read(sourceFile);
	const buffer = await buildMarkdownDocxArrayBuffer(markdown);

	infoLog('file', 'Converting Markdown to DOCX', { sourcePath: sourceFile.path, outputPath });
	const outputFile = await app.vault.createBinary(outputPath, buffer);
	try {
		await app.workspace.getLeaf('tab').openFile(outputFile, { active: true });
	} catch (error) {
		errorLog('file', 'Converted Markdown but failed to open DOCX', {
			sourcePath: sourceFile.path,
			outputPath,
			error: error instanceof Error ? error.message : String(error),
		});
		throw error;
	}
	return outputFile;
}
