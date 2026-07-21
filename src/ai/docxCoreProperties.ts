import { AI_ERROR_CODES, createAiError } from './errors';

export interface DocxCorePropertiesPatch {
	creator: string;
	lastModifiedBy: string;
}

const CORE_PROPERTIES_CLOSE_TAG = '</cp:coreProperties>';

function encodeXmlText(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function upsertCoreProperty(corePropertiesXml: string, qualifiedName: string, value: string): string {
	const encodedValue = encodeXmlText(value);
	const elementPattern = new RegExp(`(<${qualifiedName}\\b[^<>]*>)[^<]*(</${qualifiedName}>)`);
	if (elementPattern.test(corePropertiesXml)) {
		return corePropertiesXml.replace(elementPattern, `$1${encodedValue}$2`);
	}

	const closeTagIndex = corePropertiesXml.indexOf(CORE_PROPERTIES_CLOSE_TAG);
	if (closeTagIndex === -1) {
		throw createAiError(
			AI_ERROR_CODES.VALIDATION_FAILED,
			'Malformed docProps/core.xml: missing cp:coreProperties closing tag.',
		);
	}

	const element = `<${qualifiedName}>${encodedValue}</${qualifiedName}>`;
	return `${corePropertiesXml.slice(0, closeTagIndex)}${element}${corePropertiesXml.slice(closeTagIndex)}`;
}

/**
 * Update author attribution without rewriting unrelated core-property fields.
 */
export function patchDocxCoreProperties(
	corePropertiesXml: string,
	patch: DocxCorePropertiesPatch,
): string {
	return upsertCoreProperty(
		upsertCoreProperty(corePropertiesXml, 'dc:creator', patch.creator),
		'cp:lastModifiedBy',
		patch.lastModifiedBy,
	);
}
