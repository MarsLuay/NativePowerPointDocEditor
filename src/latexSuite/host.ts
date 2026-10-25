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
import {
	findMathSpansToCommit,
	unwrapLatexSuiteInner,
	type MathCommitOptions,
} from './mathSpans';

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
	private lastDocxBlockStart: number | null = null;

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
		const finish = isLatexFinishKey(event);
		const converted = convertClosedDocxMath(view, {
			commitReadyAtCaret: finish,
		});
		this.rememberDocxBlock(view);
		if (handled) {
			return true;
		}
		if (converted && event.key === 'Tab') {
			event.preventDefault();
			return true;
		}
		return false;
	}

	handleDocxCaretSettled(view: ProseMirrorView): void {
		if (!this.isActive()) {
			return;
		}
		this.commitPreviousDocxBlock(view);
		convertClosedDocxMath(view);
		this.rememberDocxBlock(view);
	}

	handleDocxBlur(view: ProseMirrorView): void {
		if (!this.isActive()) {
			return;
		}
		this.commitPreviousDocxBlock(view);
		convertClosedDocxMath(view, { commitNonEmptyAtCaret: true });
		this.rememberDocxBlock(view);
	}

	private rememberDocxBlock(view: ProseMirrorView): void {
		const { $from } = view.state.selection;
		this.lastDocxBlockStart = $from.parent.isTextblock ? $from.start() : null;
	}

	private commitPreviousDocxBlock(view: ProseMirrorView): void {
		const current = view.state.selection.$from.parent.isTextblock
			? view.state.selection.$from.start()
			: null;
		const previous = this.lastDocxBlockStart;
		if (previous == null || previous === current) {
			return;
		}
		convertDocxMathInBlock(view, previous, {}, view.state.selection);
	}

	handleTextareaKeydown(event: KeyboardEvent, editor: HTMLTextAreaElement, afterChange?: () => void): boolean {
		if (!this.isActive()) {
			return false;
		}

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
		const finish = isLatexFinishKey(event);
		const converted = commitTextareaMath(editor, afterChange, {
			commitReadyAtCaret: finish,
		});
		if (handled) {
			return true;
		}
		if (converted && event.key === 'Tab') {
			event.preventDefault();
			return true;
		}
		return false;
	}

	handleTextareaBlur(editor: HTMLTextAreaElement, afterChange?: () => void): void {
		if (!this.isActive()) {
			return;
		}
		commitTextareaMath(editor, afterChange, { commitNonEmptyAtCaret: true });
	}

	handleTextareaCaretSettled(editor: HTMLTextAreaElement, afterChange?: () => void): void {
		if (!this.isActive()) {
			return;
		}
		commitTextareaMath(editor, afterChange);
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
		this.lastDocxBlockStart = null;
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

function isLatexFinishKey(event: KeyboardEvent): boolean {
	return event.key === 'Tab' || event.key === 'Enter';
}

function convertClosedDocxMath(view: ProseMirrorView, options: MathCommitOptions = {}): boolean {
	const { $from } = view.state.selection;
	if (!$from.parent.isTextblock) {
		return false;
	}
	return convertDocxMathInBlock(view, $from.start(), options, null);
}

function convertDocxMathInBlock(
	view: ProseMirrorView,
	blockStart: number,
	options: MathCommitOptions,
	restore: { from: number; to: number } | null,
): boolean {
	if (blockStart < 1 || blockStart > view.state.doc.content.size) {
		return false;
	}
	let $pos: ReturnType<ProseMirrorView['state']['doc']['resolve']>;
	try {
		$pos = view.state.doc.resolve(blockStart);
	} catch {
		return false;
	}
	const node = $pos.parent;
	if (!node.isTextblock) {
		return false;
	}
	for (let index = 0; index < node.childCount; index += 1) {
		if (!node.child(index).isText) {
			return false;
		}
	}

	const start = $pos.start();
	const text = node.textContent;
	const end = $pos.end();
	const caret = restore?.from ?? view.state.selection.from;
	const cursorInBlock = caret >= start && caret <= end;
	const cursor = cursorInBlock ? caret - start : text.length;
	const spans = findMathSpansToCommit(text, cursor, cursorInBlock ? options : {});
	if (spans.length === 0) {
		return false;
	}

	let converted = false;
	let mapFrom = restore?.from ?? view.state.selection.from;
	let mapTo = restore?.to ?? view.state.selection.to;
	for (const span of [...spans].sort((left, right) => right.from - left.from)) {
		try {
			const inner = unwrapLatexSuiteInner(span.inner);
			if (!inner) {
				continue;
			}
			const equation = compileEquation(inner, span.display);
			const from = start + span.from;
			const to = start + span.to;
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
			if (!insertMathNodeAtSelection(view, equation)) {
				continue;
			}
			converted = true;
			const nodeSize = view.state.selection.from - from;
			const delta = nodeSize - (to - from);
			if (mapFrom >= to) {
				mapFrom += delta;
			} else if (mapFrom > from) {
				mapFrom = from + nodeSize;
			}
			if (mapTo >= to) {
				mapTo += delta;
			} else if (mapTo > from) {
				mapTo = from + nodeSize;
			}
		} catch {
			// Leave the LaTeX source in the paragraph.
		}
	}
	if (converted && restore) {
		try {
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, mapFrom, mapTo)));
		} catch {
			// Selection may no longer resolve after the replacement.
		}
	}
	return converted;
}

function commitTextareaMath(
	editor: HTMLTextAreaElement,
	afterChange: (() => void) | undefined,
	options: MathCommitOptions = {},
): boolean {
	const cursor = editor.selectionStart ?? 0;
	const converted = applyClosedMathPlainText(editor.value, cursor, options);
	if (converted.text === editor.value) {
		return false;
	}
	editor.value = converted.text;
	const caret = clamp(converted.caret, 0, converted.text.length);
	editor.setSelectionRange(caret, caret);
	afterChange?.();
	editor.dispatchEvent(new Event('input', { bubbles: true }));
	return true;
}

function applyClosedMathPlainText(
	text: string,
	cursor: number,
	options: MathCommitOptions = {},
): { text: string; caret: number } {
	const spans = findMathSpansToCommit(text, cursor, options);
	if (spans.length === 0) {
		return { text, caret: cursor };
	}

	let next = text;
	let caret = cursor;
	for (const span of [...spans].sort((left, right) => right.from - left.from)) {
		try {
			const inner = unwrapLatexSuiteInner(span.inner);
			if (!inner) {
				continue;
			}
			const replacement = compileEquation(inner, span.display).plainText;
			next = `${next.slice(0, span.from)}${replacement}${next.slice(span.to)}`;
			const delta = replacement.length - (span.to - span.from);
			if (caret >= span.to) {
				caret += delta;
			} else if (caret > span.from) {
				caret = span.from + replacement.length;
			}
		} catch {
			// Keep the original span.
		}
	}
	return { text: next, caret };
}
