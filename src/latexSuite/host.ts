// eslint-disable-next-line import/no-extraneous-dependencies -- provided by Obsidian; esbuild keeps this external
import { EditorSelection, EditorState } from '@codemirror/state';
// eslint-disable-next-line import/no-extraneous-dependencies -- provided by Obsidian; esbuild keeps this external
import { EditorView } from '@codemirror/view';
import type { App } from 'obsidian';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import type { EditorView as ProseMirrorView } from 'prosemirror-view';

import { compileEquation } from '../equation/compileEquation';
import { insertMathNodeAtSelection } from '../equation/insertMathNode';
import { debugLog, errorLog } from '../logger';
import { getLatexSuitePlugin } from './detect';
import {
	findMathSpansToCommit,
	unwrapLatexSuiteInner,
	type MathCommitOptions,
	type MathSpanDisplay,
} from './mathSpans';
import {
	encodeHiddenMath,
	expandPlainTextMath,
	latexCaretOffsets,
	mathInnerForEdit,
	wrapLatexForEdit,
} from './mathSource';

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
		if (!this.isActive() || view.composing || !view.editable) {
			return false;
		}

		if (expandDocxMathForArrow(view, event)) {
			event.preventDefault();
			this.rememberDocxBlock(view);
			return true;
		}

		const { $from, $to } = view.state.selection;
		if (!$from.parent.isTextblock || $from.parent !== $to.parent) {
			return false;
		}

		let handled = false;
		if (textblockIsPlainText($from.parent)) {
			handled = this.handleKeydown(event, {
				getText: () => $from.parent.textContent,
				getSelection: () => ({ from: $from.parentOffset, to: $to.parentOffset }),
				setText: (text, selection) => replaceTextblockText(view, text, selection.from, selection.to),
			});
		}
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

	handleDocxPointerUp(view: ProseMirrorView, event: Event): void {
		if (!this.isActive() || !view.editable) {
			return;
		}
		if ('button' in event && (event as PointerEvent).button !== 0) {
			return;
		}
		if (view.state.selection.from !== view.state.selection.to) {
			if (view.hasFocus()) {
				this.handleDocxCaretSettled(view);
			}
			return;
		}
		const mathPos = mathPosFromPointerEvent(view, event);
		if (mathPos != null && expandMathNode(view, mathPos, 'start')) {
			this.rememberDocxBlock(view);
			return;
		}
		if (!view.hasFocus()) {
			return;
		}
		this.handleDocxCaretSettled(view);
	}

	handleDocxCaretSettled(view: ProseMirrorView): void {
		if (!this.isActive()) {
			return;
		}
		if (expandDocxMathAtSelection(view)) {
			this.rememberDocxBlock(view);
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

		if (expandTextareaMath(editor, afterChange, { key: event.key })) {
			event.preventDefault();
			return true;
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

	handleTextareaCaretSettled(
		editor: HTMLTextAreaElement,
		afterChange?: () => void,
		options: { fromPointer?: boolean } = {},
	): void {
		if (!this.isActive()) {
			return;
		}
		if (expandTextareaMath(editor, afterChange, { fromPointer: options.fromPointer })) {
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

	const start = $pos.start();
	const ranges = collectTextblockRanges(node, start);
	if (ranges.length === 0) {
		return false;
	}

	const caret = restore?.from ?? view.state.selection.from;
	const jobs: Array<{ from: number; to: number; inner: string; display: MathSpanDisplay }> = [];
	for (const range of ranges) {
		const inRange = caret >= range.from && caret <= range.from + range.text.length;
		const cursor = inRange ? caret - range.from : range.text.length;
		const spans = findMathSpansToCommit(range.text, cursor, inRange ? options : {});
		for (const span of spans) {
			jobs.push({
				from: range.from + span.from,
				to: range.from + span.to,
				inner: span.inner,
				display: span.display,
			});
		}
	}
	if (jobs.length === 0) {
		return false;
	}

	let converted = false;
	let mapFrom = restore?.from ?? view.state.selection.from;
	let mapTo = restore?.to ?? view.state.selection.to;
	for (const span of [...jobs].sort((left, right) => right.from - left.from)) {
		try {
			const inner = unwrapLatexSuiteInner(span.inner);
			if (!inner) {
				continue;
			}
			const equation = compileEquation(inner, span.display);
			const from = span.from;
			const to = span.to;
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
	const converted = commitPlainTextMath(editor.value, cursor, options);
	if (converted.text === editor.value) {
		return false;
	}
	applyTextareaUpdate(editor, converted.text, converted.caret, afterChange);
	return true;
}

function commitPlainTextMath(
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
			const equation = compileEquation(inner, span.display);
			const replacement = `${equation.plainText}${encodeHiddenMath(span.display, inner, equation.plainText)}`;
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

function expandTextareaMath(
	editor: HTMLTextAreaElement,
	afterChange: (() => void) | undefined,
	options: { key?: string; fromPointer?: boolean },
): boolean {
	if (editor.selectionStart !== editor.selectionEnd) {
		return false;
	}
	const expanded = expandPlainTextMath(editor.value, editor.selectionStart ?? 0, options);
	if (!expanded) {
		return false;
	}
	applyTextareaUpdate(editor, expanded.text, expanded.caret, afterChange);
	return true;
}

function applyTextareaUpdate(
	editor: HTMLTextAreaElement,
	text: string,
	caret: number,
	afterChange: (() => void) | undefined,
): void {
	editor.value = text;
	const next = clamp(caret, 0, text.length);
	editor.setSelectionRange(next, next);
	afterChange?.();
	editor.dispatchEvent(new Event('input', { bubbles: true }));
}

function textblockIsPlainText(node: ProseMirrorNode): boolean {
	for (let index = 0; index < node.childCount; index += 1) {
		if (!node.child(index).isText) {
			return false;
		}
	}
	return true;
}

function collectTextblockRanges(
	node: ProseMirrorNode,
	blockStart: number,
): Array<{ from: number; text: string }> {
	const ranges: Array<{ from: number; text: string }> = [];
	node.forEach((child, offset) => {
		if (!child.isText || !child.text) {
			return;
		}
		const from = blockStart + offset;
		const last = ranges[ranges.length - 1];
		if (last && last.from + last.text.length === from) {
			last.text += child.text;
			return;
		}
		ranges.push({ from, text: child.text });
	});
	return ranges;
}

function expandDocxMathForArrow(view: ProseMirrorView, event: KeyboardEvent): boolean {
	if (event.shiftKey || event.metaKey || event.ctrlKey || event.altKey) {
		return false;
	}
	if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') {
		return false;
	}
	const selection = view.state.selection;
	if (!selection.empty && !(selection instanceof NodeSelection)) {
		return false;
	}
	if (selection instanceof NodeSelection && selection.node.type.name === 'math') {
		return expandMathNode(view, selection.from, event.key === 'ArrowLeft' ? 'end' : 'start');
	}
	const $from = selection.$from;
	if (event.key === 'ArrowRight' && $from.nodeAfter?.type.name === 'math') {
		return expandMathNode(view, $from.pos, 'start');
	}
	if (event.key === 'ArrowLeft' && $from.nodeBefore?.type.name === 'math') {
		return expandMathNode(view, $from.pos - $from.nodeBefore.nodeSize, 'end');
	}
	return false;
}

function expandDocxMathAtSelection(view: ProseMirrorView): boolean {
	const selection = view.state.selection;
	if (selection instanceof NodeSelection && selection.node.type.name === 'math') {
		return expandMathNode(view, selection.from, 'start');
	}
	return false;
}

function mathPosFromPointerEvent(view: ProseMirrorView, event: Event): number | null {
	const target = eventTargetElement(event.target);
	const span = target?.closest('[data-pm-start][data-pm-end]');
	if (!span || !('dataset' in span)) {
		return null;
	}
	const pos = Number((span as HTMLElement).dataset.pmStart);
	if (!Number.isFinite(pos) || pos < 0 || pos > view.state.doc.content.size) {
		return null;
	}
	try {
		return view.state.doc.nodeAt(pos)?.type.name === 'math' ? pos : null;
	} catch {
		return null;
	}
}

function eventTargetElement(raw: EventTarget | null | undefined): { closest: (selector: string) => Element | null } | null {
	if (!raw || typeof raw !== 'object') {
		return null;
	}
	if ('nodeType' in raw && (raw as Node).nodeType === 3) {
		const parent = (raw as Node).parentElement;
		return parent && typeof parent.closest === 'function' ? parent : null;
	}
	if ('closest' in raw && typeof (raw as Element).closest === 'function') {
		return raw as Element;
	}
	return null;
}

function expandMathNode(
	view: ProseMirrorView,
	pos: number,
	caret: 'start' | 'end',
): boolean {
	const node = view.state.doc.nodeAt(pos);
	if (!node || node.type.name !== 'math') {
		return false;
	}
	const display: MathSpanDisplay = node.attrs.display === 'block' ? 'block' : 'inline';
	const source = mathInnerForEdit(
		String(node.attrs.ommlXml ?? ''),
		String(node.attrs.plainText ?? ''),
		display,
	);
	const wrapped = wrapLatexForEdit(source.inner, source.display);
	const marks = view.state.storedMarks ?? view.state.selection.$from.marks();
	const textNode = view.state.schema.text(wrapped, marks);
	let transaction = view.state.tr.replaceWith(pos, pos + node.nodeSize, textNode);
	const offsets = latexCaretOffsets(wrapped, source.display);
	const caretPos = pos + (caret === 'end' ? offsets.end : offsets.start);
	try {
		transaction = transaction.setSelection(TextSelection.create(transaction.doc, caretPos));
	} catch {
		return false;
	}
	view.dispatch(transaction.scrollIntoView());
	debugLog('editor', 'Expanded a math node to editable LaTeX', {
		pos,
		display: source.display,
		length: wrapped.length,
	});
	return true;
}
