import JSZip from 'jszip';

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const DOCUMENT_RELS_PATH = 'word/_rels/document.xml.rels';
const STYLES_PATH = 'word/styles.xml';
const STYLES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const STYLES_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const FALLBACK_PARAGRAPH_STYLE_IDS = ['Normal', 'Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3'] as const;

export interface DocxDefaultStylesResult {
	buffer: ArrayBuffer;
	addedDefaultStyles: boolean;
}

export const DEFAULT_DOCX_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:docDefaults>
		<w:rPrDefault>
			<w:rPr>
				<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
				<w:sz w:val="22"/>
				<w:szCs w:val="22"/>
				<w:lang w:val="en-US"/>
			</w:rPr>
		</w:rPrDefault>
		<w:pPrDefault>
			<w:pPr>
				<w:spacing w:after="160" w:line="276" w:lineRule="auto"/>
			</w:pPr>
		</w:pPrDefault>
	</w:docDefaults>
	<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
		<w:name w:val="Normal"/>
		<w:qFormat/>
		<w:rsid w:val="00000000"/>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Title">
		<w:name w:val="Title"/>
		<w:basedOn w:val="Normal"/>
		<w:next w:val="Normal"/>
		<w:uiPriority w:val="1"/>
		<w:qFormat/>
		<w:pPr>
			<w:spacing w:before="240" w:after="120"/>
		</w:pPr>
		<w:rPr>
			<w:b/>
			<w:bCs/>
			<w:sz w:val="52"/>
			<w:szCs w:val="52"/>
		</w:rPr>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Subtitle">
		<w:name w:val="Subtitle"/>
		<w:basedOn w:val="Normal"/>
		<w:next w:val="Normal"/>
		<w:uiPriority w:val="2"/>
		<w:qFormat/>
		<w:pPr>
			<w:spacing w:after="160"/>
		</w:pPr>
		<w:rPr>
			<w:color w:val="666666"/>
			<w:sz w:val="30"/>
			<w:szCs w:val="30"/>
		</w:rPr>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Heading1">
		<w:name w:val="Heading 1"/>
		<w:basedOn w:val="Normal"/>
		<w:next w:val="Normal"/>
		<w:uiPriority w:val="3"/>
		<w:qFormat/>
		<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="240" w:after="0"/>
			<w:outlineLvl w:val="0"/>
		</w:pPr>
		<w:rPr>
			<w:b/>
			<w:bCs/>
			<w:color w:val="4A6C8C"/>
			<w:sz w:val="40"/>
			<w:szCs w:val="40"/>
		</w:rPr>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Heading2">
		<w:name w:val="Heading 2"/>
		<w:basedOn w:val="Normal"/>
		<w:next w:val="Normal"/>
		<w:uiPriority w:val="4"/>
		<w:qFormat/>
		<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="200" w:after="0"/>
			<w:outlineLvl w:val="1"/>
		</w:pPr>
		<w:rPr>
			<w:b/>
			<w:bCs/>
			<w:color w:val="4A6C8C"/>
			<w:sz w:val="32"/>
			<w:szCs w:val="32"/>
		</w:rPr>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Heading3">
		<w:name w:val="Heading 3"/>
		<w:basedOn w:val="Normal"/>
		<w:next w:val="Normal"/>
		<w:uiPriority w:val="5"/>
		<w:qFormat/>
		<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="160" w:after="0"/>
			<w:outlineLvl w:val="2"/>
		</w:pPr>
		<w:rPr>
			<w:b/>
			<w:bCs/>
			<w:color w:val="4A6C8C"/>
			<w:sz w:val="28"/>
			<w:szCs w:val="28"/>
		</w:rPr>
	</w:style>
</w:styles>`;

function escapeRegExp(value: string) {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function createParagraphStyleOpeningTagPattern(styleId: string) {
	const escapedStyleId = escapeRegExp(styleId);
	return new RegExp(`<w:style\\b(?=[^>]*\\bw:type=["']paragraph["'])(?=[^>]*\\bw:styleId=["']${escapedStyleId}["'])[^>]*>`, 'i');
}

function hasParagraphStyle(stylesXml: string, styleId: string) {
	return createParagraphStyleOpeningTagPattern(styleId).test(stylesXml);
}

function getDefaultParagraphStyleXml(styleId: string) {
	const openingTagPattern = createParagraphStyleOpeningTagPattern(styleId);
	const openingTagMatch = DEFAULT_DOCX_STYLES_XML.match(openingTagPattern);
	if (openingTagMatch?.index === undefined) {
		return null;
	}

	const startIndex = openingTagMatch.index;
	const endIndex = DEFAULT_DOCX_STYLES_XML.indexOf('</w:style>', startIndex);
	if (endIndex === -1) {
		return null;
	}

	return DEFAULT_DOCX_STYLES_XML.slice(startIndex, endIndex + '</w:style>'.length);
}

function addMissingDefaultParagraphStyles(stylesXml: string) {
	let nextStylesXml = stylesXml;
	let addedDefaultStyles = false;

	for (const styleId of FALLBACK_PARAGRAPH_STYLE_IDS) {
		if (hasParagraphStyle(nextStylesXml, styleId)) {
			continue;
		}

		const styleXml = getDefaultParagraphStyleXml(styleId);
		if (!styleXml) {
			continue;
		}

		const updatedStylesXml = nextStylesXml.replace(/<\/w:styles>/i, `${styleXml}</w:styles>`);
		if (updatedStylesXml === nextStylesXml) {
			continue;
		}

		nextStylesXml = updatedStylesXml;
		addedDefaultStyles = true;
	}

	return { stylesXml: nextStylesXml, addedDefaultStyles };
}

function createDefaultContentTypesXml() {
	return [
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
		`<Override PartName="/${STYLES_PATH}" ContentType="${STYLES_CONTENT_TYPE}"/>`,
		'</Types>',
	].join('');
}

function addStylesContentType(contentTypesXml: string | undefined) {
	if (!contentTypesXml) {
		return createDefaultContentTypesXml();
	}

	if (/PartName=["']\/word\/styles\.xml["']/i.test(contentTypesXml)) {
		return contentTypesXml;
	}

	const overrideXml = `<Override PartName="/${STYLES_PATH}" ContentType="${STYLES_CONTENT_TYPE}"/>`;
	if (/<Types\b[^>]*\/>/i.test(contentTypesXml)) {
		return contentTypesXml.replace(/<Types\b([^>]*)\/>/i, `<Types$1>${overrideXml}</Types>`);
	}

	return contentTypesXml.replace(/<\/Types>/i, `${overrideXml}</Types>`);
}

function getNextRelationshipId(relsXml: string) {
	const usedIds = new Set<string>();
	for (const match of relsXml.matchAll(/\bId=["']([^"']+)["']/g)) {
		const id = match[1];
		if (id) {
			usedIds.add(id);
		}
	}

	for (let index = 1; index < 1000; index += 1) {
		const candidate = `rId${index}`;
		if (!usedIds.has(candidate)) {
			return candidate;
		}
	}

	return `rIdNativePowerPointDocEditorStyles${Date.now()}`;
}

function createDefaultDocumentRelsXml() {
	return [
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
		`<Relationship Id="rId1" Type="${STYLES_RELATIONSHIP_TYPE}" Target="styles.xml"/>`,
		'</Relationships>',
	].join('');
}

function addStylesRelationship(relsXml: string | undefined) {
	if (!relsXml) {
		return createDefaultDocumentRelsXml();
	}

	if (relsXml.includes(STYLES_RELATIONSHIP_TYPE) || /\bTarget=["']styles\.xml["']/i.test(relsXml)) {
		return relsXml;
	}

	const relationshipXml = `<Relationship Id="${getNextRelationshipId(relsXml)}" Type="${STYLES_RELATIONSHIP_TYPE}" Target="styles.xml"/>`;
	if (/<Relationships\b[^>]*\/>/i.test(relsXml)) {
		return relsXml.replace(/<Relationships\b([^>]*)\/>/i, `<Relationships$1>${relationshipXml}</Relationships>`);
	}

	return relsXml.replace(/<\/Relationships>/i, `${relationshipXml}</Relationships>`);
}

export async function ensureDocxDefaultStyles(buffer: ArrayBuffer): Promise<DocxDefaultStylesResult> {
	const zip = await JSZip.loadAsync(buffer.slice(0));
	const stylesFile = zip.file(STYLES_PATH);
	const existingStylesXml = await stylesFile?.async('string');
	const contentTypesXml = await zip.file(CONTENT_TYPES_PATH)?.async('string');
	const relsXml = await zip.file(DOCUMENT_RELS_PATH)?.async('string');
	const nextContentTypesXml = addStylesContentType(contentTypesXml);
	const nextRelsXml = addStylesRelationship(relsXml);
	const stylesResult = existingStylesXml
		? addMissingDefaultParagraphStyles(existingStylesXml)
		: { stylesXml: DEFAULT_DOCX_STYLES_XML, addedDefaultStyles: true };
	const packageChanged =
		stylesResult.addedDefaultStyles
		|| nextContentTypesXml !== contentTypesXml
		|| nextRelsXml !== relsXml;

	if (!packageChanged) {
		return { buffer, addedDefaultStyles: false };
	}

	zip.file(STYLES_PATH, stylesResult.stylesXml);
	zip.file(CONTENT_TYPES_PATH, nextContentTypesXml);
	zip.file(DOCUMENT_RELS_PATH, nextRelsXml);

	return {
		buffer: await zip.generateAsync({ type: 'arraybuffer' }),
		addedDefaultStyles: packageChanged,
	};
}
