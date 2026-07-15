import type JSZip from 'jszip';
import { listDocxDescribeParts } from './docxParts';
import type { DocxPatchSession } from './docxPatchSession';

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const DOCUMENT_RELS_PATH = 'word/_rels/document.xml.rels';

const COMMENT_PARTS = [
	{
		path: 'word/comments.xml',
		target: 'comments.xml',
		relationshipType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
	},
	{
		path: 'word/commentsExtended.xml',
		target: 'commentsExtended.xml',
		relationshipType: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
	},
	{
		path: 'word/commentsIds.xml',
		target: 'commentsIds.xml',
		relationshipType: 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
	},
	{
		path: 'word/commentsExtensible.xml',
		target: 'commentsExtensible.xml',
		relationshipType: 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
	},
] as const;

const COMMENT_RANGE_START_PATTERN = /<w:commentRangeStart\b[^>]*(?:\/\s*>|>\s*<\/w:commentRangeStart\s*>)/g;
const COMMENT_RANGE_END_PATTERN = /<w:commentRangeEnd\b[^>]*(?:\/\s*>|>\s*<\/w:commentRangeEnd\s*>)/g;
const COMMENT_REFERENCE_PATTERN = /<w:commentReference\b[^>]*(?:\/\s*>|>\s*<\/w:commentReference\s*>)/g;
const XML_EMPTY_ELEMENT_PATTERN = /<(Override|Relationship)\b[^>]*(?:\/\s*>|>\s*<\/\1\s*>)/g;

export interface DocxCommentRemovalResult {
	changedPartPaths: string[];
	commentCount: number;
	documentXml: string;
}

function commentRelationshipPartPath(partPath: string): string {
	const filename = partPath.split('/').pop();
	return filename ? `word/_rels/${filename}.rels` : partPath;
}

function xmlElementHasAttributeValue(elementXml: string, attribute: string, value: string): boolean {
	const escapedAttribute = attribute.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	const escapedValue = value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	return new RegExp(`\\b${escapedAttribute}\\s*=\\s*(["'])${escapedValue}\\1`).test(elementXml);
}

function removeCommentContentTypeOverrides(contentTypesXml: string): string {
	return contentTypesXml.replace(XML_EMPTY_ELEMENT_PATTERN, (elementXml) => {
		if (!elementXml.startsWith('<Override')) return elementXml;
		return COMMENT_PARTS.some(({ path }) => xmlElementHasAttributeValue(elementXml, 'PartName', `/${path}`))
			? ''
			: elementXml;
	});
}

function removeCommentRelationships(relsXml: string): string {
	return relsXml.replace(XML_EMPTY_ELEMENT_PATTERN, (elementXml) => {
		if (!elementXml.startsWith('<Relationship')) return elementXml;
		return COMMENT_PARTS.some(({ target, relationshipType, path }) =>
			xmlElementHasAttributeValue(elementXml, 'Type', relationshipType)
			|| xmlElementHasAttributeValue(elementXml, 'Target', target)
			|| xmlElementHasAttributeValue(elementXml, 'Target', path)
			|| xmlElementHasAttributeValue(elementXml, 'Target', `/${path}`),
		)
			? ''
			: elementXml;
	});
}

/**
 * Remove inline anchors while preserving surrounding runs and all non-comment
 * review markup. An empty CommentReference run is valid OOXML and is safer to
 * retain than deleting an unfamiliar run that happens to carry a reference.
 */
export function removeDocxCommentAnchors(xml: string): string {
	return xml
		.replace(COMMENT_RANGE_START_PATTERN, '')
		.replace(COMMENT_RANGE_END_PATTERN, '')
		.replace(COMMENT_REFERENCE_PATTERN, '');
}

async function removeCommentPartPackageEntries(zip: JSZip, changedPartPaths: Set<string>): Promise<void> {
	for (const { path } of COMMENT_PARTS) {
		if (zip.file(path)) {
			zip.remove(path);
			changedPartPaths.add(path);
		}

		const relationshipPartPath = commentRelationshipPartPath(path);
		if (zip.file(relationshipPartPath)) {
			zip.remove(relationshipPartPath);
			changedPartPaths.add(relationshipPartPath);
		}
	}

	const contentTypesFile = zip.file(CONTENT_TYPES_PATH);
	if (contentTypesFile) {
		const contentTypesXml = await contentTypesFile.async('string');
		const nextContentTypesXml = removeCommentContentTypeOverrides(contentTypesXml);
		if (nextContentTypesXml !== contentTypesXml) {
			zip.file(CONTENT_TYPES_PATH, nextContentTypesXml);
			changedPartPaths.add(CONTENT_TYPES_PATH);
		}
	}

	const relsFile = zip.file(DOCUMENT_RELS_PATH);
	if (relsFile) {
		const relsXml = await relsFile.async('string');
		const nextRelsXml = removeCommentRelationships(relsXml);
		if (nextRelsXml !== relsXml) {
			zip.file(DOCUMENT_RELS_PATH, nextRelsXml);
			changedPartPaths.add(DOCUMENT_RELS_PATH);
		}
	}
}

/**
 * Remove every DOCX comment annotation and its package metadata without
 * reserializing unrelated document content.
 */
export async function removeAllDocxComments(session: DocxPatchSession): Promise<DocxCommentRemovalResult> {
	const zip = session.getZip();
	const commentsXml = await zip.file('word/comments.xml')?.async('string');
	const commentCount = commentsXml?.match(/<w:comment\b/g)?.length ?? 0;
	const changedPartPaths = new Set<string>();

	for (const { path } of listDocxDescribeParts(zip)) {
		if (!session.hasPart(path)) continue;
		const partXml = session.getPartXml(path);
		const nextPartXml = removeDocxCommentAnchors(partXml);
		if (nextPartXml !== partXml) {
			session.setPartXml(path, nextPartXml);
			changedPartPaths.add(path);
		}
	}

	await removeCommentPartPackageEntries(zip, changedPartPaths);

	return {
		changedPartPaths: [...changedPartPaths],
		commentCount,
		documentXml: session.getDocumentXml(),
	};
}
