export interface DocumentWordCount {
	totalWords: number;
	selectedWords: number | null;
}

/**
 * Mirrors the familiar document-editor definition of a word: a contiguous
 * run of non-whitespace characters. This retains punctuation, hyphenated
 * terms, identifiers, and non-Latin text as the author entered them.
 */
export function countDocumentWords(text: string): number {
	return text.match(/[^\s]+/gu)?.length ?? 0;
}

export function formatDocumentWordCount({ totalWords, selectedWords }: DocumentWordCount): string {
	const count = selectedWords ?? totalWords;
	const label = count === 1 ? 'word' : 'words';
	return selectedWords === null ? `${count.toLocaleString()} ${label}` : `${count.toLocaleString()} selected ${label}`;
}
