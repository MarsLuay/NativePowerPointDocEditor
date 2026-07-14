import JSZip from 'jszip';
import { findTopLevelBlock, insertBlockAfter } from './docxBlockResolver';
import { getDocumentBodyInner } from './docxOoxml';
import { buildInlineImageParagraphXml, extractBlipEmbedId, paragraphContainsDrawing } from './docxOoxmlWrite';
import { AI_ERROR_CODES, createAiError } from './errors';

const RELS_PATH = 'word/_rels/document.xml.rels';
const CONTENT_TYPES_PATH = '[Content_Types].xml';
const IMAGE_CONTENT_TYPE: Record<string, string> = {
	png: 'image/png',
	jpg: 'image/jpeg',
	jpeg: 'image/jpeg',
	gif: 'image/gif',
	webp: 'image/webp',
	bmp: 'image/bmp',
};

function nextRelationshipId(relsXml: string): string {
	const matches = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]));
	const max = matches.length > 0 ? Math.max(...matches) : 0;
	return `rId${max + 1}`;
}

function nextImagePartName(zip: JSZip, extension: string): string {
	let index = 1;
	while (zip.file(`word/media/image${index}.${extension}`)) {
		index++;
	}
	return `image${index}.${extension}`;
}

function appendRelationship(relsXml: string, relationshipId: string, target: string): string {
	const relationship = `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`;
	if (relsXml.includes('</Relationships>')) {
		return relsXml.replace('</Relationships>', `${relationship}</Relationships>`);
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship}</Relationships>`;
}

function normalizeExtension(extension: string): string {
	const lowered = extension.toLowerCase();
	if (lowered === 'jpeg') return 'jpg';
	return lowered;
}

export async function addInlineImage(
	zip: JSZip,
	documentXml: string,
	afterBlockId: string,
	imageBytes: Uint8Array,
	extension: string,
): Promise<string> {
	const normalizedExtension = normalizeExtension(extension);
	const mediaFileName = nextImagePartName(zip, normalizedExtension);
	const mediaPath = `word/media/${mediaFileName}`;
	zip.file(mediaPath, imageBytes);

	let relsXml = (await zip.file(RELS_PATH)?.async('string')) ?? '';
	const relationshipId = nextRelationshipId(relsXml);
	relsXml = appendRelationship(relsXml, relationshipId, `media/${mediaFileName}`);
	zip.file(RELS_PATH, relsXml);

	let contentTypesXml = (await zip.file(CONTENT_TYPES_PATH)?.async('string')) ?? '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
	if (!contentTypesXml.includes(`Extension="${normalizedExtension}"`)) {
		const contentType = IMAGE_CONTENT_TYPE[normalizedExtension] ?? 'image/png';
		contentTypesXml = contentTypesXml.replace(
			'</Types>',
			`<Default Extension="${normalizedExtension}" ContentType="${contentType}"/></Types>`,
		);
	}
	zip.file(CONTENT_TYPES_PATH, contentTypesXml);

	const paragraphXml = buildInlineImageParagraphXml(relationshipId);
	return insertBlockAfter(documentXml, afterBlockId, paragraphXml);
}

export async function replaceInlineImage(
	zip: JSZip,
	documentXml: string,
	blockId: string,
	imageBytes: Uint8Array,
	extension: string,
): Promise<string> {
	const bodyInner = getDocumentBodyInner(documentXml);
	if (bodyInner === null) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Missing w:body in document.xml.');
	}

	const block = findTopLevelBlock(bodyInner, blockId);
	if (block.kind !== 'paragraph' || !paragraphContainsDrawing(block.xml)) {
		throw createAiError(
			AI_ERROR_CODES.BLOCK_NOT_FOUND,
			`Image block ${blockId} was not found. Use a describe() image block (body/p[N] with an embedded drawing).`,
			{ field: 'blockId' },
		);
	}

	const relationshipId = extractBlipEmbedId(block.xml);
	if (!relationshipId) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Image relationship for ${blockId} was not found.`, { field: 'blockId' });
	}

	const relsXml = await zip.file(RELS_PATH)?.async('string');
	if (!relsXml) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Missing word/_rels/document.xml.rels.');
	}

	const relationshipMatch = new RegExp(
		`<Relationship\\b[^>]*Id="${relationshipId}"[^>]*Target="([^"]+)"`,
	).exec(relsXml);
	if (!relationshipMatch?.[1]) {
		throw createAiError(AI_ERROR_CODES.BLOCK_NOT_FOUND, `Image media for ${blockId} was not found.`, { field: 'blockId' });
	}

	const normalizedExtension = normalizeExtension(extension);
	void normalizedExtension;
	const mediaTarget = relationshipMatch[1].replace(/^media\//, '');
	const mediaPath = mediaTarget.startsWith('word/') ? mediaTarget : `word/media/${mediaTarget}`;
	zip.file(mediaPath, imageBytes);

	return documentXml;
}
