import type JSZip from 'jszip';
import type { DocxPartKind, DocxStableLocation } from './docxStableIds';

export const DOCX_DOCUMENT_PATH = 'word/document.xml';
export const DOCX_CORE_PROPERTIES_PATH = 'docProps/core.xml';
export const DOCX_FOOTNOTES_PATH = 'word/footnotes.xml';
export const DOCX_ENDNOTES_PATH = 'word/endnotes.xml';
export const DOCX_COMMENTS_PATH = 'word/comments.xml';

const HEADER_PATH = /^word\/headers\/header(\d+)\.xml$/;
const FOOTER_PATH = /^word\/footers\/footer(\d+)\.xml$/;

const REVIEW_MARKUP_PATTERN = /<w:(?:ins|del|moveFrom|moveTo|commentRangeStart|commentReference|rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange|sectPrChange|numberingChange|cellIns|cellDel|cellMerge)\b/;

export type DocxWrapperTag = 'body' | 'hdr' | 'ftr';

export interface DocxListedPart {
	path: string;
	part: DocxPartKind;
	partNumber: number | null;
}

export function listDocxDescribeParts(zip: JSZip): DocxListedPart[] {
	const parts: DocxListedPart[] = [{ path: DOCX_DOCUMENT_PATH, part: 'body', partNumber: null }];

	for (const path of Object.keys(zip.files).sort()) {
		const headerMatch = HEADER_PATH.exec(path);
		if (headerMatch) {
			parts.push({ path, part: 'header', partNumber: Number(headerMatch[1]) });
			continue;
		}
		const footerMatch = FOOTER_PATH.exec(path);
		if (footerMatch) {
			parts.push({ path, part: 'footer', partNumber: Number(footerMatch[1]) });
		}
	}

	if (zip.file(DOCX_FOOTNOTES_PATH)) {
		parts.push({ path: DOCX_FOOTNOTES_PATH, part: 'footnotes', partNumber: null });
	}
	if (zip.file(DOCX_ENDNOTES_PATH)) {
		parts.push({ path: DOCX_ENDNOTES_PATH, part: 'endnotes', partNumber: null });
	}

	return parts;
}

export function resolvePartPath(location: Pick<DocxStableLocation, 'part' | 'partNumber'>): string {
	switch (location.part) {
		case 'body':
			return DOCX_DOCUMENT_PATH;
		case 'header':
			return `word/headers/header${location.partNumber ?? 1}.xml`;
		case 'footer':
			return `word/footers/footer${location.partNumber ?? 1}.xml`;
		case 'footnotes':
			return DOCX_FOOTNOTES_PATH;
		case 'endnotes':
			return DOCX_ENDNOTES_PATH;
	}
}

export function wrapperTagForPart(part: DocxPartKind): DocxWrapperTag | 'footnotes' | 'endnotes' {
	switch (part) {
		case 'body':
			return 'body';
		case 'header':
			return 'hdr';
		case 'footer':
			return 'ftr';
		case 'footnotes':
			return 'footnotes';
		case 'endnotes':
			return 'endnotes';
	}
}

export function listReplaceTextPartPaths(zip: JSZip): string[] {
	return listDocxDescribeParts(zip)
		.map((entry) => entry.path)
		.filter((path) => path !== DOCX_COMMENTS_PATH);
}

export function hasTrackChangesMarkup(xml: string): boolean {
	return REVIEW_MARKUP_PATTERN.test(xml);
}

export async function scanDocxReviewState(zip: JSZip): Promise<{ hasTrackChanges: boolean; hasComments: boolean }> {
	let hasTrackChanges = false;
	let hasComments = false;

	for (const listed of listDocxDescribeParts(zip)) {
		const xml = await zip.file(listed.path)?.async('string');
		if (!xml) continue;
		if (hasTrackChangesMarkup(xml)) {
			hasTrackChanges = true;
		}
	}

	const commentsXml = await zip.file(DOCX_COMMENTS_PATH)?.async('string');
	if (commentsXml && /<w:comment\b/.test(commentsXml)) {
		hasComments = true;
	}

	return { hasTrackChanges, hasComments };
}
