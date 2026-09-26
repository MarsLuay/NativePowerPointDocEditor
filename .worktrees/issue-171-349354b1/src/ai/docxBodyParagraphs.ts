import { getDocumentBodyInner, replaceDocumentBodyInner } from './docxOoxml';
import { AI_ERROR_CODES, createAiError } from './errors';

function encodeXmlText(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function extractTrailingSectPr(bodyInner: string): { content: string; sectPr: string } {
	const match = /(<w:sectPr\b[\s\S]*?<\/w:sectPr>)\s*$/.exec(bodyInner);
	if (!match?.[1]) {
		return { content: bodyInner, sectPr: '' };
	}
	return {
		content: bodyInner.slice(0, match.index),
		sectPr: match[1],
	};
}

function paragraphXmlForText(text: string): string {
	return `<w:p><w:r><w:t xml:space="preserve">${encodeXmlText(text)}</w:t></w:r></w:p>`;
}

/**
 * Replace all top-level body blocks with plain paragraphs.
 * Preserves trailing `w:sectPr`. Tables/images in the body are removed.
 */
export function applyReplaceBodyParagraphs(documentXml: string, paragraphs: string[]): string {
	if (!Array.isArray(paragraphs)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'paragraphs must be an array of strings.', {
			field: 'paragraphs',
		});
	}
	const texts = paragraphs.map((entry, index) => {
		if (typeof entry !== 'string') {
			throw createAiError(
				AI_ERROR_CODES.SCHEMA_INVALID,
				`paragraphs[${index}] must be a string.`,
				{ field: 'paragraphs' },
			);
		}
		return entry;
	});
	const normalized = texts.length > 0 ? texts : [''];

	const bodyInner = getDocumentBodyInner(documentXml);
	if (bodyInner === null) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Missing w:body wrapper in document.xml.');
	}

	const { sectPr } = extractTrailingSectPr(bodyInner);
	const nextInner = `${normalized.map(paragraphXmlForText).join('')}${sectPr}`;
	return replaceDocumentBodyInner(documentXml, nextInner);
}
