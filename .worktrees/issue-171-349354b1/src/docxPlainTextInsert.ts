import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export type PlainTextInsertRange = {
	from: number;
	to: number;
	collapsedCrossBlock: boolean;
};

const INHERITED_PLAIN_PASTE_ATTRS = [
	'defaultTextFormatting',
	'styleId',
	'numPr',
	'listIsBullet',
	'listNumFmt',
	'listMarker',
	'listMarkerHidden',
	'listMarkerFontFamily',
	'listMarkerFontSize',
	'listMarkerSuffix',
	'listLevelNumFmts',
	'listAbstractNumId',
	'listStartOverride',
	'lineSpacing',
	'lineSpacingRule',
	'spaceAfter',
	'spaceBefore',
	'contextualSpacing',
	'indentLeft',
	'indentRight',
	'indentFirstLine',
	'hangingIndent',
] as const;

/**
 * Plain `tr.insertText(text, from, to)` deletes every node between `from` and `to`.
 * Across textblocks that includes paragraph breaks, so a one-line paste over an
 * accidental multi-block selection (list markers, drag select) silently drops
 * adjacent skill lines. Collapse cross-block ranges to a caret at `from`.
 */
export function resolveSafePlainTextInsertRange(
	state: EditorState,
	from: number,
	to: number,
): PlainTextInsertRange {
	const safeFrom = Math.max(0, Math.min(from, state.doc.content.size));
	const safeTo = Math.max(0, Math.min(to, state.doc.content.size));
	const start = Math.min(safeFrom, safeTo);
	const end = Math.max(safeFrom, safeTo);

	if (start === end) {
		return { from: start, to: end, collapsedCrossBlock: false };
	}

	const $from = state.doc.resolve(start);
	const $to = state.doc.resolve(end);
	if ($from.sameParent($to) && $from.parent.isTextblock) {
		return { from: start, to: end, collapsedCrossBlock: false };
	}

	return { from: start, to: start, collapsedCrossBlock: true };
}

/** Keep empty strings so `\n\n` becomes a real blank paragraph (PM default collapses them). */
export function splitPlainTextLines(text: string): string[] {
	return text.split(/\r\n|\n|\r/);
}

export function marksToDefaultTextFormatting(marks: readonly Mark[]): Record<string, unknown> | null {
	const formatting: Record<string, unknown> = {};

	for (const mark of marks) {
		const attrs = mark.attrs as Record<string, unknown>;
		switch (mark.type.name) {
			case 'bold':
				formatting.bold = true;
				break;
			case 'italic':
				formatting.italic = true;
				break;
			case 'underline':
				formatting.underline = { style: attrs.style || 'single' };
				break;
			case 'strike':
				formatting.strike = true;
				break;
			case 'textColor':
				formatting.color = attrs;
				break;
			case 'highlight':
				formatting.highlight = attrs.color;
				break;
			case 'fontSize':
				formatting.fontSize = attrs.size ?? attrs.sizeCs;
				if (attrs.sizeCs != null) {
					formatting.fontSizeCs = attrs.sizeCs;
				}
				break;
			case 'fontFamily':
				formatting.fontFamily = {
					ascii: attrs.ascii,
					hAnsi: attrs.hAnsi,
				};
				break;
			case 'superscript':
				formatting.vertAlign = 'superscript';
				break;
			case 'subscript':
				formatting.vertAlign = 'subscript';
				break;
			case 'rtl':
				formatting.rtl = true;
				break;
			default:
				break;
		}
	}

	return Object.keys(formatting).length > 0 ? formatting : null;
}

function inheritPlainPasteParagraphAttrs(
	source: ProseMirrorNode | null,
	defaultTextFormatting: Record<string, unknown> | null,
): Record<string, unknown> {
	const attrs: Record<string, unknown> = {};
	if (source?.type.name === 'paragraph') {
		const sourceAttrs = source.attrs as Record<string, unknown>;
		for (const key of INHERITED_PLAIN_PASTE_ATTRS) {
			const value = sourceAttrs[key];
			if (value != null) {
				attrs[key] = value;
			}
		}
	}

	if (defaultTextFormatting) {
		attrs.defaultTextFormatting = defaultTextFormatting;
		const sourceAttrs = source?.attrs;
		const original = sourceAttrs?._originalFormatting as Record<string, unknown> | null | undefined;
		if (original && typeof original === 'object') {
			attrs._originalFormatting = {
				...original,
				runProperties: defaultTextFormatting,
			};
		}
	}

	return attrs;
}

export function buildPlainTextInsertTransaction(
	state: EditorState,
	text: string,
	from = state.selection.from,
	to = state.selection.to,
): { transaction: Transaction; range: PlainTextInsertRange } {
	const range = resolveSafePlainTextInsertRange(state, from, to);
	const transaction = state.tr.insertText(text, range.from, range.to);
	const insertionEnd = transaction.mapping.map(range.to, 1);
	transaction
		.setSelection(TextSelection.near(transaction.doc.resolve(insertionEnd), -1))
		.scrollIntoView();
	return { transaction, range };
}

export function insertPlainTypedText(
	view: EditorView,
	text: string,
	from = view.state.selection.from,
	to = view.state.selection.to,
): PlainTextInsertRange {
	const { transaction, range } = buildPlainTextInsertTransaction(view.state, text, from, to);
	view.dispatch(transaction);
	return range;
}

/**
 * Paste multi-line plain text as real paragraphs, preserving blank lines and
 * empty-paragraph font defaults (storedMarks / defaultTextFormatting).
 * Stock PM clipboard parsing collapses `\n+` and drops empty section gaps.
 */
export function buildPlainTextParagraphSlice(
	state: EditorState,
	text: string,
	from = state.selection.from,
): { slice: Slice; rangeHint: { marks: readonly Mark[]; source: ProseMirrorNode | null } } {
	const $from = state.doc.resolve(Math.max(0, Math.min(from, state.doc.content.size)));
	const source = $from.parent.type.name === 'paragraph' ? $from.parent : null;
	const marks = state.storedMarks ?? $from.marks();
	const fromMarks = marksToDefaultTextFormatting(marks);
	const defaultTextFormatting = fromMarks
		?? (source?.attrs.defaultTextFormatting as Record<string, unknown> | null | undefined)
		?? null;
	const baseAttrs = inheritPlainPasteParagraphAttrs(source, defaultTextFormatting);
	const paragraphType = state.schema.nodes.paragraph;
	if (!paragraphType) {
		throw new Error('Schema is missing paragraph node');
	}

	const lines = splitPlainTextLines(text);
	const nodes = lines.map((line) => {
		if (line.length === 0) {
			return paragraphType.create(baseAttrs);
		}
		return paragraphType.create(baseAttrs, state.schema.text(line, marks));
	});

	// openStart/openEnd 1 merges first/last into surrounding textblocks like Word paste.
	return {
		slice: new Slice(Fragment.from(nodes), 1, 1),
		rangeHint: { marks, source },
	};
}

export function insertPlainTextAsParagraphs(
	view: EditorView,
	text: string,
	from = view.state.selection.from,
	to = view.state.selection.to,
): PlainTextInsertRange {
	if (!/[\r\n]/.test(text)) {
		return insertPlainTypedText(view, text, from, to);
	}

	const safeFrom = Math.max(0, Math.min(from, to, view.state.doc.content.size));
	const safeTo = Math.max(0, Math.min(Math.max(from, to), view.state.doc.content.size));
	const $from = view.state.doc.resolve(safeFrom);
	const $to = view.state.doc.resolve(safeTo);
	const crossBlock = safeFrom !== safeTo && !$from.sameParent($to);

	// Multi-line paste must replace the real selection via replaceRange + a
	// paragraph Slice. Collapsing cross-block to a caret (safe for insertText)
	// left the old selection in place; a later delete of that shifted range
	// looked like "paste vanished after autosave".
	const { slice } = buildPlainTextParagraphSlice(view.state, text, safeFrom);
	const transaction = view.state.tr.replaceRange(safeFrom, safeTo, slice);
	// `replaceRange` maps the old range selection over the inserted content.
	// Paste must leave a caret after the inserted paragraphs, like Word. Keeping
	// the pasted block selected made the next history operation target it.
	const insertionEnd = transaction.mapping.map(safeTo, 1);
	transaction
		.setSelection(TextSelection.near(transaction.doc.resolve(insertionEnd), -1))
		.scrollIntoView();
	view.dispatch(transaction);
	return { from: safeFrom, to: safeTo, collapsedCrossBlock: crossBlock };
}

/**
 * Replace a selection with ProseMirror's already-parsed rich clipboard slice.
 * Unlike plain text insertion, the slice retains distinct marks for every run
 * (for example, a bold section title followed by unbolded skill details).
 */
export function insertRichClipboardSlice(
	view: EditorView,
	slice: Slice,
	from = view.state.selection.from,
	to = view.state.selection.to,
): PlainTextInsertRange {
	const safeFrom = Math.max(0, Math.min(from, to, view.state.doc.content.size));
	const safeTo = Math.max(0, Math.min(Math.max(from, to), view.state.doc.content.size));
	const $from = view.state.doc.resolve(safeFrom);
	const $to = view.state.doc.resolve(safeTo);
	const crossBlock = safeFrom !== safeTo && !$from.sameParent($to);
	// A parsed clipboard Slice normally opens its first/last paragraphs so a
	// one-line paste can merge with surrounding text. Structured paste instead
	// replaces a block range: keep those paragraphs closed, otherwise the last
	// pasted bullet (or empty paragraph defaults) merges into the following
	// paragraph and its attrs are discarded.
	const structuredSlice = slice.openStart || slice.openEnd
		? new Slice(slice.content, 0, 0)
		: slice;
	const transaction = view.state.tr.replaceRange(safeFrom, safeTo, structuredSlice);
	const insertionEnd = transaction.mapping.map(safeTo, 1);
	transaction
		.setSelection(TextSelection.near(transaction.doc.resolve(insertionEnd), -1))
		.setMeta('paste', true)
		.setMeta('uiEvent', 'paste')
		.scrollIntoView();
	view.dispatch(transaction);
	return { from: safeFrom, to: safeTo, collapsedCrossBlock: crossBlock };
}

export function summarizeRichClipboardSlice(slice: Slice): {
	listParagraphs: number;
	paragraphsWithBorders: number;
	emptyParagraphs: number;
	emptyParagraphsWithFontSize: number;
} {
	let listParagraphs = 0;
	let paragraphsWithBorders = 0;
	let emptyParagraphs = 0;
	let emptyParagraphsWithFontSize = 0;
	slice.content.nodesBetween(0, slice.content.size, (node) => {
		if (node.type.name !== 'paragraph') return true;
		if (node.attrs.numPr) {
			listParagraphs += 1;
		}
		if (node.attrs.borders) {
			paragraphsWithBorders += 1;
		}
		if (node.content.size === 0) {
			emptyParagraphs += 1;
			const defaults = node.attrs.defaultTextFormatting as { fontSize?: unknown } | null | undefined;
			if (typeof defaults?.fontSize === 'number') {
				emptyParagraphsWithFontSize += 1;
			}
		}
		return true;
	});
	return { listParagraphs, paragraphsWithBorders, emptyParagraphs, emptyParagraphsWithFontSize };
}

export function countDocTextblocks(doc: { descendants: (f: (node: { isTextblock: boolean }) => boolean | void) => void }): number {
	let count = 0;
	doc.descendants((node) => {
		if (node.isTextblock) {
			count += 1;
		}
		return true;
	});
	return count;
}

export function summarizeTransactionSteps(transaction: Transaction): string[] {
	const names: string[] = [];
	transaction.steps.forEach((step) => {
		const name = step.constructor?.name ?? 'Step';
		try {
			const json: unknown = typeof step.toJSON === 'function' ? step.toJSON() : null;
			if (json && typeof json === 'object') {
				const stepType = typeof (json as { stepType?: unknown }).stepType === 'string'
					? (json as { stepType: string }).stepType
					: null;
				const from = (json as { from?: unknown }).from;
				const to = (json as { to?: unknown }).to;
				names.push(
					stepType
						? `${name}/${stepType}:${String(from)}->${String(to)}`
						: `${name}:${JSON.stringify(json).slice(0, 160)}`,
				);
				return;
			}
		} catch {
			// Fall through to the constructor name.
		}
		names.push(name);
	});
	return names;
}

export function summarizeTransactionMeta(transaction: Transaction): string[] {
	const keys: string[] = [];
	const metaBag = (transaction as unknown as { meta?: Map<string, unknown> | Record<string, unknown> }).meta;
	if (!metaBag) {
		return keys;
	}
	if (metaBag instanceof Map) {
		for (const key of metaBag.keys()) {
			keys.push(String(key));
		}
		return keys;
	}
	if (typeof metaBag === 'object') {
		for (const key of Object.keys(metaBag)) {
			keys.push(key);
		}
	}
	return keys;
}
