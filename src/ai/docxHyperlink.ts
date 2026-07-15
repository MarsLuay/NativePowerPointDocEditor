import JSZip from 'jszip';
import { resolvePartPath } from './docxParts';
import type { DocxStableLocation } from './docxStableIds';
import { AI_ERROR_CODES, createAiError } from './errors';

const HYPERLINK_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

function encodeXmlAttribute(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');
}

export function relsPathForPart(partPath: string): string {
	const fileName = partPath.split('/').pop();
	if (!fileName) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, `Invalid part path: ${partPath}.`);
	}
	return `word/_rels/${fileName}.rels`;
}

function nextRelationshipId(relsXml: string): string {
	const matches = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]));
	const max = matches.length > 0 ? Math.max(...matches) : 0;
	return `rId${max + 1}`;
}

function appendRelationship(
	relsXml: string,
	relationshipId: string,
	target: string,
	type: string,
	targetMode?: string,
): string {
	const targetModeAttr = targetMode ? ` TargetMode="${targetMode}"` : '';
	const relationship = `<Relationship Id="${relationshipId}" Type="${type}" Target="${encodeXmlAttribute(target)}"${targetModeAttr}/>`;
	if (relsXml.includes('</Relationships>')) {
		return relsXml.replace('</Relationships>', `${relationship}</Relationships>`);
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship}</Relationships>`;
}

export async function registerExternalHyperlink(
	zip: JSZip,
	location: Pick<DocxStableLocation, 'part' | 'partNumber'>,
	url: string,
): Promise<string> {
	if (!/^https?:\/\//i.test(url) && !/^mailto:/i.test(url)) {
		throw createAiError(
			AI_ERROR_CODES.VALIDATION_FAILED,
			'Hyperlink url must use http, https, or mailto scheme.',
			{ field: 'url' },
		);
	}

	const partPath = resolvePartPath(location);
	const relsPath = relsPathForPart(partPath);
	let relsXml = (await zip.file(relsPath)?.async('string')) ?? '';
	const relationshipId = nextRelationshipId(relsXml);
	relsXml = appendRelationship(relsXml, relationshipId, url, HYPERLINK_REL_TYPE, 'External');
	zip.file(relsPath, relsXml);
	return relationshipId;
}

export function buildHyperlinkXml(
	relationshipId: string,
	innerXml: string,
	tooltip?: string,
): string {
	const attrs = [`r:id="${relationshipId}"`];
	if (tooltip) {
		attrs.push(`w:tooltip="${encodeXmlAttribute(tooltip)}"`);
	}
	return `<w:hyperlink ${attrs.join(' ')}>${innerXml}</w:hyperlink>`;
}
