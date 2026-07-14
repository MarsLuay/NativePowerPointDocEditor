import type { EditorMode } from '@eigenpal/docx-editor-react';
import type { EditorView } from 'prosemirror-view';
import type {
	ChromeListener,
	Disposable,
	DocxEditorAdapter,
	DocxFindMatch,
	DocxFindOptions,
} from './DocxEditorAdapter';

export interface FakeDocxEditorAdapterOptions {
	serialize?: () => Promise<ArrayBuffer | null>;
	getView?: () => EditorView | null;
	getSelectedText?: () => string;
	find?: (searchText: string, options: DocxFindOptions) => DocxFindMatch[];
	select?: (match: DocxFindMatch) => boolean;
	replace?: (match: DocxFindMatch, replacement: string) => boolean;
	replaceAll?: (matches: readonly DocxFindMatch[], replacement: string) => boolean;
	insertImage?: (file: File) => Promise<{ width: number; height: number } | null>;
}

/** In-memory adapter for session and view tests. */
export class FakeDocxEditorAdapter implements DocxEditorAdapter {
	readonly modes: EditorMode[] = [];
	private readonly chromeListeners = new Set<ChromeListener>();

	constructor(private readonly options: FakeDocxEditorAdapterOptions = {}) {}

	serialize(): Promise<ArrayBuffer | null> {
		return this.options.serialize?.() ?? Promise.resolve(null);
	}

	observeChrome(listener: ChromeListener): Disposable {
		this.chromeListeners.add(listener);
		return { dispose: () => this.chromeListeners.delete(listener) };
	}

	emitChrome(records: MutationRecord[] = []): void {
		for (const listener of this.chromeListeners) listener(records);
	}

	setMode(mode: EditorMode): void {
		this.modes.push(mode);
	}

	getView(): EditorView | null {
		return this.options.getView?.() ?? null;
	}

	getSelectedText(): string {
		return this.options.getSelectedText?.() ?? '';
	}

	find(searchText: string, options: DocxFindOptions): DocxFindMatch[] {
		return this.options.find?.(searchText, options) ?? [];
	}

	select(match: DocxFindMatch): boolean {
		return this.options.select?.(match) ?? false;
	}

	replace(match: DocxFindMatch, replacement: string): boolean {
		return this.options.replace?.(match, replacement) ?? false;
	}

	replaceAll(matches: readonly DocxFindMatch[], replacement: string): boolean {
		return this.options.replaceAll?.(matches, replacement) ?? false;
	}

	insertImage(file: File): Promise<{ width: number; height: number } | null> {
		return this.options.insertImage?.(file) ?? Promise.resolve(null);
	}
}
