// eslint-disable-next-line import/no-extraneous-dependencies -- provided by Obsidian; esbuild keeps this external
import { EditorSelection, EditorState } from '@codemirror/state';
// eslint-disable-next-line import/no-extraneous-dependencies -- provided by Obsidian; esbuild keeps this external
import { EditorView } from '@codemirror/view';
import type { App } from 'obsidian';
import type { EditorView as ProseMirrorView } from 'prosemirror-view';
import { TextSelection } from 'prosemirror-state';

import { compileEquation } from '../equation/compileEquation';
import { insertMathNodeAtSelection } from '../equation/insertMathNode';
import { debugLog, errorLog } from '../logger';
import { getLatexSuitePlugin } from './detect';
import { findClosedMathSpans } from './mathSpans';

export interface LatexSuiteSession {
	getText(): string;
	getSelection(): { from: number; to: number };
	setText(text: string, selection: { from: number; to: number }): void;
}

const HOST_CLASS = 'native-powerpoint-latex-suite-host';

export class LatexSuiteHost {
	private view: EditorView | null = null;
	private parent: HTMLElement | null = null;
	private enabled = false;

	constructor(
		private readonly app: App,
		private readonly isSettingOn: () => boolean,
	) {}

	setEnabled(enabled: boolean): void {
		this.enabled = enabled;
		this.destroyView();
	}

	isActive(): boolean {
		return this.enabled && this.isSettingOn() && getLatexSuitePlugin(this.app) !== null;
	}

	handleKeydown(event: KeyboardEvent, session: LatexSuiteSession): boolean {
		if (!this.isActive() || event.isComposing || event.metaKey || event.ctrlKey || event.altKey) {
			return false;
		}

		if (!this.ensureView()) {
			return false;
		}

		const view = this.view;
		if (!view) {
			return false;
		}

		const latex = getLatexSuitePlugin(this.app);
		const keyboardPlugin = latex?.api?.plugin?.keyboardEventPlugin;
		const handler = keyboardPlugin ? view.plugin(keyboardPlugin) : null;
		if (!handler?.onKeydown) {
			return false;
		}

		const before = session.getText();
		const selection = session.getSelection();
		const from = clamp(selection.from, 0, before.length);
		const to = clamp(selection.to, 0, before.length);
		try {
			view.dispatch({
				changes: { from: 0, to: view.state.doc.length, insert: before },
				selection: EditorSelection.single(from, to),
			});
			handler.onKeydown(event, view);
		} catch (error) {
			errorLog('editor', 'LaTeX Suite key dispatch failed', error);
			return false;
		}

		if (!event.defaultPrevented) {
			return false;
		}

		const next = view.state.doc.toString();
		const nextSelection = view.state.selection.main;
		session.setText(next, {
			from: clamp(nextSelection.from, 0, next.length),
			to: clamp(nextSelection.to, 0, next.length),
		});
		debugLog('editor', 'LaTeX Suite handled a key in an Office editor', {
			key: event.key,
			beforeLength: before.length,
			afterLength: next.length,
		});
		return true;
	}

	handleDocxKeydown(view: ProseMirrorView, event: KeyboardEvent): boolean {
		if (!this.isActive() || view.composing) {
			return false;
		}

		const { $from, $to } = view.state.selection;
		if (!$from.parent.isTextblock || $from.parent !== $to.parent) {
			return false;
		}
		for (let index = 0; index < $from.parent.childCount; index += 1) {
			if (!$from.parent.child(index).isText) {
				return false;
			}
		}

		const handled = this.handleKeydown(event, {
			getText: () => $from.parent.textContent,
			getSelection: () => ({ from: $from.parentOffset, to: $to.parentOffset }),
			setText: (text, selection) => replaceTextblockText(view, text, selection.from, selection.to),
		});
		if (!handled) {
			return false;
		}

		convertClosedDocxMath(view);
		return true;
	}

	handleTextareaKeydown(event: KeyboardEvent, editor: HTMLTextAreaElement, afterChange?: () => void): boolean {
		const handled = this.handleKeydown(event, {
			getText: () => editor.value,
			getSelection: () => ({
				from: editor.selectionStart ?? 0,
				to: editor.selectionEnd ?? 0,
			}),
			setText: (text, selection) => {
				editor.value = text;
				editor.setSelectionRange(selection.from, selection.to);
				afterChange?.();
				editor.dispatchEvent(new Event('input', { bubbles: true }));
			},
		});
		if (!handled) {
			return false;
		}

		const cursor = editor.selectionStart ?? 0;
		const converted = applyClosedMathPlainText(editor.value, cursor);
		if (converted !== editor.value) {
			editor.value = converted;
			const caret = Math.min(cursor, converted.length);
			editor.setSelectionRange(caret, caret);
			afterChange?.();
			editor.dispatchEvent(new Event('input', { bubbles: true }));
		}
		return true;
	}

	destroy(): void {
		this.enabled = false;
		this.destroyView();
	}

	private ensureView(): boolean {
		const latex = getLatexSuitePlugin(this.app);
		const extensions = latex?.editorExtensions;
		const keyboardPlugin = latex?.api?.plugin?.keyboardEventPlugin;
		if (!extensions?.length || !keyboardPlugin) {
			this.destroyView();
			return false;
		}
		if (this.view?.plugin(keyboardPlugin)?.onKeydown) {
			return true;
		}
		this.destroyView();

		const parent = this.app.workspace.containerEl.createDiv({
			cls: HOST_CLASS,
			attr: {
				'data-native-powerpoint-latex-suite-host': 'true',
				'aria-hidden': 'true',
			},
		});

		try {
			this.view = new EditorView({
				parent,
				state: EditorState.create({
					doc: '',
					extensions: [...extensions],
				}),
			});
			this.parent = parent;
			if (!this.view.plugin(keyboardPlugin)?.onKeydown) {
				throw new Error('LaTeX Suite keyboard plugin did not attach');
			}
			return true;
		} catch (error) {
			parent.remove();
			errorLog('editor', 'Could not host LaTeX Suite in the Office editors', error);
			this.view = null;
			this.parent = null;
			return false;
		}
	}

	private destroyView(): void {
		this.view?.destroy();
		this.view = null;
		this.parent?.remove();
		this.parent = null;
	}
}

let boundHost: LatexSuiteHost | null = null;

export function bindLatexSuiteHost(host: LatexSuiteHost | null): void {
	boundHost = host;
}

export function getLatexSuiteHost(): LatexSuiteHost | null {
	return boundHost;
}

function clamp(value: number, min: number, max: number): number {
	return Math.min(max, Math.max(min, value));
}

function replaceTextblockText(
	view: ProseMirrorView,
	text: string,
	anchor: number,
	head: number,
): void {
	const { $from } = view.state.selection;
	const start = $from.start();
	const end = $from.end();
	const marks = $from.marks();
	const node = text ? view.state.schema.text(text, marks) : null;
	let transaction = view.state.tr.replaceWith(start, end, node ? [node] : []);
	const resolvedAnchor = start + clamp(anchor, 0, text.length);
	const resolvedHead = start + clamp(head, 0, text.length);
	transaction = transaction.setSelection(TextSelection.create(transaction.doc, resolvedAnchor, resolvedHead));
	view.dispatch(transaction.scrollIntoView());
}

function convertClosedDocxMath(view: ProseMirrorView): void {
	const { $from } = view.state.selection;
	if (!$from.parent.isTextblock) {
		return;
	}
	for (let index = 0; index < $from.parent.childCount; index += 1) {
		if (!$from.parent.child(index).isText) {
			return;
		}
	}

	const text = $from.parent.textContent;
	const cursor = $from.parentOffset;
	const spans = findClosedMathSpans(text, cursor);
	if (spans.length === 0) {
		return;
	}

	const start = $from.start();
	for (const span of [...spans].sort((left, right) => right.from - left.from)) {
		try {
			const equation = compileEquation(span.inner, span.display);
			const from = start + span.from;
			const to = start + span.to;
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
			if (!insertMathNodeAtSelection(view, equation)) {
				view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, start + cursor)));
			}
		} catch {
			// Leave the LaTeX source in the paragraph.
		}
	}
}

function applyClosedMathPlainText(text: string, cursor: number): string {
	const spans = findClosedMathSpans(text, cursor);
	if (spans.length === 0) {
		return text;
	}

	let next = text;
	for (const span of [...spans].sort((left, right) => right.from - left.from)) {
		try {
			next = `${next.slice(0, span.from)}${compileEquation(span.inner, span.display).plainText}${next.slice(span.to)}`;
		} catch {
			// Keep the original span.
		}
	}
	return next;
}
