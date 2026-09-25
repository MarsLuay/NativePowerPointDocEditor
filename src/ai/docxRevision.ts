export interface DocxRevisionEntry {
	path: string;
	xml: string;
}

/**
 * Deterministic, session-local DOCX epoch. It intentionally hashes only the
 * loaded editable XML parts so describe/apply agree without exposing content.
 */
export function computeDocxRevision(entries: readonly DocxRevisionEntry[]): string {
	let hash = 0x811c9dc5;
	for (const entry of [...entries].sort((left, right) => left.path.localeCompare(right.path))) {
		const value = `${entry.path}\u0000${entry.xml}\u0000`;
		for (let index = 0; index < value.length; index += 1) {
			hash ^= value.charCodeAt(index);
			hash = Math.imul(hash, 0x01000193) >>> 0;
		}
	}
	return `docx-${hash.toString(16).padStart(8, '0')}`;
}
