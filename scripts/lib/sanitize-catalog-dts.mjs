/**
 * Rewrite package declaration emit so Obsidian catalog ESLint accepts it.
 * Keeps runtime JS unchanged; only adjusts public `.d.ts` / `.d.mts` text.
 */

/**
 * @param {string} source
 * @returns {string}
 */
export function sanitizeCatalogDts(source) {
	let out = source;

	// Popout-window DOM type / value: never leave a `globalThis` token.
	out = out.replace(/\bglobalThis\.Document\b/g, "Window['document']");
	out = out.replace(/\bglobalThis\.document\b/g, 'window.document');
	out = out.replace(/\bglobalThis\b/g, 'window');

	// TS private-field emit stub in .d.ts (own line or inline in class body)
	out = out.replace(/#private;\s*/g, '');

	// OOXML/model fields named `document` → quoted property name (not Identifier)
	out = out.replace(/^(\s*)document(\s*\??\s*:)/gm, "$1['document']$2");
	out = out.replace(/([{;,]\s*)document(\s*\??\s*:)/g, "$1['document']$2");

	// Params named `document` → docxDocument (declaration surface only)
	out = out.replace(/\((\s*)document(\s*\??\s*:)/g, '($1docxDocument$2');
	out = out.replace(/(,\s*)document(\s*\??\s*:)/g, '$1docxDocument$2');

	// Relationship URI literals overridden by `| string`
	out = out.replace(
		/type RelationshipType\s*=\s*(?:'[^']+'\s*\|\s*)+string\s*;/g,
		'type RelationshipType = string;',
	);

	// Value used only as a type
	out = out.replace(/:\s*typeof\s+cssColorToHex\b/g, ': typeof rgbToHex');

	// unknown overrides other constituents
	out = out.replace(/\bunknown\s*\|\s*null\b/g, 'unknown');
	out = out.replace(/\bnull\s*\|\s*unknown\b/g, 'unknown');

	// Explicit any (catalog @typescript-eslint/no-explicit-any, including any[])
	out = out.replace(/=\s*any\b/g, '= unknown');
	out = out.replace(/:\s*any\b(?!\[)/g, ': unknown');
	out = out.replace(/\bany\[\]/g, 'unknown[]');
	out = out.replace(/<any,\s*any>/g, '<string, string>');
	out = out.replace(/Record<string,\s*any>/g, 'Record<string, unknown>');

	return out;
}
