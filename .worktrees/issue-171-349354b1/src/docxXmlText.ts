const TEXT_TOKEN_PATTERN = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>|<w:br\b[^>]*\/>|<w:cr\b[^>]*\/>|<\/w:p>/g;
const RUN_TOKEN_PATTERN = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>|<w:br\b[^>]*\/>|<w:cr\b[^>]*\/>/g;

export function decodeDocxXmlEntities(value: string): string {
	return value
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&amp;/g, '&')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&#(\d+);/g, (_match, codePoint: string) => {
			const numericCodePoint = Number(codePoint);
			return Number.isFinite(numericCodePoint) ? String.fromCodePoint(numericCodePoint) : '';
		})
		.replace(/&#x([0-9a-fA-F]+);/g, (_match, codePoint: string) => {
			const numericCodePoint = Number.parseInt(codePoint, 16);
			return Number.isFinite(numericCodePoint) ? String.fromCodePoint(numericCodePoint) : '';
		});
}

export function normalizeDocxExtractedText(value: string): string {
	return value
		.replace(/\r/g, '\n')
		.replace(/[ \t]+\n/g, '\n')
		.replace(/\n[ \t]+/g, '\n')
		.replace(/\n{3,}/g, '\n\n')
		.trim();
}

export function extractDocxTextFromXml(xml: string): string {
	const pieces: string[] = [];
	let match: RegExpExecArray | null;
	TEXT_TOKEN_PATTERN.lastIndex = 0;

	while ((match = TEXT_TOKEN_PATTERN.exec(xml)) !== null) {
		const [token, text] = match;
		if (text !== undefined) {
			pieces.push(decodeDocxXmlEntities(text));
		} else if (token.startsWith('<w:tab')) {
			pieces.push('\t');
		} else {
			pieces.push('\n');
		}
	}

	return normalizeDocxExtractedText(pieces.join(''));
}

export function extractDocxRunText(runXml: string): string {
	let result = '';
	let match: RegExpExecArray | null;
	RUN_TOKEN_PATTERN.lastIndex = 0;

	while ((match = RUN_TOKEN_PATTERN.exec(runXml)) !== null) {
		const [token, text] = match;
		if (text !== undefined) {
			result += decodeDocxXmlEntities(text);
		} else if (token.startsWith('<w:tab')) {
			result += '\t';
		} else {
			result += '\n';
		}
	}

	return result;
}
