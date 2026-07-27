/**
 * Creates a detached HTML node for DOCX editor chrome.
 *
 * Obsidian augments Document#createEl/createDiv to append the result to the
 * document. Chrome fallbacks must stay detached until their caller inserts
 * them, because appending a second document root throws HierarchyRequestError.
 */
export function createDetachedDocxEditorChromeElement<K extends keyof HTMLElementTagNameMap>(
	parent: HTMLElement,
	tagName: K,
): HTMLElementTagNameMap[K] {
	return parent.ownerDocument.createElement(tagName);
}
