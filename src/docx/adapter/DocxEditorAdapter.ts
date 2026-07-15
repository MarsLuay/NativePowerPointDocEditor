import type { DocxEditorRef, EditorMode } from '../runtime';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

const MAX_INSERTED_IMAGE_WIDTH = 612;

export interface Disposable {
	dispose(): void;
}

export type ChromeListener = (records: MutationRecord[]) => void;

export interface DocxFindMatch {
	from: number;
	to: number;
	text: string;
}

export interface DocxFindOptions {
	matchCase: boolean;
	wholeWord: boolean;
}

export interface DocxEditorAdapter {
	/**
	 * Serializes the current document through the editor's supported ref API.
	 * `null` means the editor has not mounted yet.
	 */
	serialize(): Promise<ArrayBuffer | null>;
	observeChrome(listener: ChromeListener): Disposable;
	setMode(mode: EditorMode): void;
	getView(): EditorView | null;
	getSelectedText(): string;
	find(searchText: string, options: DocxFindOptions): DocxFindMatch[];
	select(match: DocxFindMatch): boolean;
	replace(match: DocxFindMatch, replacement: string): boolean;
	replaceAll(matches: readonly DocxFindMatch[], replacement: string): boolean;
	insertImage(file: File): Promise<{ width: number; height: number } | null>;
}

/**
 * Transitional binding used by the React host. Keep editor-package calls here
 * while callers migrate onto the stable adapter surface.
 */
export interface DocxEditorAdapterController extends DocxEditorAdapter {
	bindEditor(getEditor: () => DocxEditorRef | null): void;
	bindMode(onSetMode: (mode: EditorMode) => void): void;
}

export function createDocxEditorAdapter(chromeTarget: HTMLElement): DocxEditorAdapterController {
	let getEditor: (() => DocxEditorRef | null) | null = null;
	let onSetMode: ((mode: EditorMode) => void) | null = null;

	const getView = () => getEditor?.()?.getEditorRef()?.getView() ?? null;

	return {
		bindEditor(nextGetEditor) {
			getEditor = nextGetEditor;
		},
		bindMode(nextOnSetMode) {
			onSetMode = nextOnSetMode;
		},
		async serialize() {
			const editor = getEditor?.() ?? null;
			return editor ? await editor.save({ selective: false }) : null;
		},
		observeChrome(listener) {
			const observer = new MutationObserver(listener);
			observer.observe(chromeTarget, {
				childList: true,
				subtree: true,
				attributes: true,
				attributeFilter: ['title', 'aria-label'],
			});
			return { dispose: () => observer.disconnect() };
		},
		setMode(mode) {
			onSetMode?.(mode);
		},
		getView,
		getSelectedText() {
			return getEditor?.()?.getSelectionInfo()?.selectedText?.trim() ?? '';
		},
		find(searchText, { matchCase, wholeWord }) {
			const view = getView();
			const normalizedSearchText = searchText.trim();
			if (!view || !normalizedSearchText) return [];

			const escaped = normalizedSearchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
			const pattern = new RegExp(wholeWord ? `\\b${escaped}\\b` : escaped, matchCase ? 'g' : 'gi');
			const matches: DocxFindMatch[] = [];
			view.state.doc.descendants((node, pos) => {
				if (!node.isTextblock) return true;
				for (const match of node.textContent.matchAll(pattern)) {
					const index = match.index ?? -1;
					if (index >= 0) {
						matches.push({
							from: pos + 1 + index,
							to: pos + 1 + index + match[0].length,
							text: match[0],
						});
					}
				}
				return false;
			});
			return matches;
		},
		select(match) {
			const view = getView();
			if (!view) return false;
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, match.from, match.to)).scrollIntoView());
			getEditor?.()?.scrollToPosition(match.from);
			return true;
		},
		replace(match, replacement) {
			const view = getView();
			if (!view) return false;
			const textNode = replacement ? view.state.schema.text(replacement) : null;
			view.dispatch(view.state.tr.replaceWith(match.from, match.to, textNode ? [textNode] : []).scrollIntoView());
			return true;
		},
		replaceAll(matches, replacement) {
			const view = getView();
			if (!view || matches.length === 0) return false;
			let transaction = view.state.tr;
			for (const match of [...matches].sort((a, b) => b.from - a.from)) {
				const textNode = replacement ? view.state.schema.text(replacement) : null;
				transaction = transaction.replaceWith(match.from, match.to, textNode ? [textNode] : []);
			}
			view.dispatch(transaction.scrollIntoView());
			return true;
		},
		async insertImage(file) {
			const view = getView();
			const imageNodeType = view?.state.schema.nodes.image;
			if (!view || !imageNodeType) return null;

			const src = await new Promise<string>((resolve, reject) => {
				const reader = new FileReader();
				reader.addEventListener('error', () => reject(reader.error ?? new Error('Could not read image file')), { once: true });
				reader.addEventListener('load', () => typeof reader.result === 'string'
					? resolve(reader.result)
					: reject(new Error('Image file did not produce a data URL')), { once: true });
				reader.readAsDataURL(file);
			});
			const naturalDimensions = await new Promise<{ width: number; height: number }>((resolve, reject) => {
				const image = new Image();
				image.addEventListener('error', () => reject(new Error('Could not load image')), { once: true });
				image.addEventListener('load', () => resolve({ width: image.naturalWidth, height: image.naturalHeight }), { once: true });
				image.src = src;
			});
			const dimensions = naturalDimensions.width > MAX_INSERTED_IMAGE_WIDTH
				? {
					width: MAX_INSERTED_IMAGE_WIDTH,
					height: Math.round(naturalDimensions.height * (MAX_INSERTED_IMAGE_WIDTH / naturalDimensions.width)),
				}
				: naturalDimensions;
			const imageNode = imageNodeType.create({
				src,
				alt: file.name,
				...dimensions,
				rId: `rId_img_${Date.now()}`,
				wrapType: 'inline',
				displayMode: 'inline',
			});
			view.dispatch(view.state.tr.insert(view.state.selection.from, imageNode).scrollIntoView());
			view.focus();
			return dimensions;
		},
	};
}
