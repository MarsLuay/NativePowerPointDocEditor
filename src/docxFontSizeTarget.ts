import type { DocxTextRange } from './docx/runtime/contract';

export type DocxFormattingTargetSource =
	| 'selection'
	| 'preserved-selection'
	| 'caret-paragraph'
	| 'rendered-paragraph'
	| 'empty-paragraph';

export interface ResolveDocxFormattingTargetInput {
	selection: DocxTextRange;
	preservedSelection?: DocxTextRange | null;
	caretParagraph?: DocxTextRange | null;
	renderedParagraph?: DocxTextRange | null;
	preferCurrentSelection?: boolean;
}

export interface ResolvedDocxFormattingTarget {
	range: DocxTextRange;
	source: DocxFormattingTargetSource;
}

function hasText(range: DocxTextRange | null | undefined): range is DocxTextRange {
	return Boolean(range && range.to > range.from);
}

function isCollapsed(range: DocxTextRange | null | undefined): range is DocxTextRange {
	return Boolean(range && range.to === range.from);
}

function overlaps(left: DocxTextRange, right: DocxTextRange): boolean {
	return left.from < right.to && right.from < left.to;
}

/**
 * Resolve the text range a text-formatting command should format.
 *
 * DOCX pagination paints a visible, read-only page over a hidden ProseMirror
 * editor. A toolbar focus transition can therefore leave the PM caret on a
 * zero-height spacer even though the last visible click was on a real line.
 * Keyboard navigation explicitly owns the current ProseMirror selection.
 * Otherwise prefer an explicit selection, a selection captured before toolbar
 * focus, the current paragraph, and finally the last painted paragraph.
 */
export function resolveDocxFormattingTarget(
	input: ResolveDocxFormattingTargetInput,
): ResolvedDocxFormattingTarget {
	if (input.preferCurrentSelection) {
		if (hasText(input.selection)) {
			return { range: input.selection, source: 'selection' };
		}
		if (hasText(input.caretParagraph)) {
			return { range: input.caretParagraph, source: 'caret-paragraph' };
		}
		return { range: input.selection, source: 'empty-paragraph' };
	}
	if (isCollapsed(input.renderedParagraph)) {
		return { range: input.renderedParagraph, source: 'empty-paragraph' };
	}
	const renderedParagraph = hasText(input.renderedParagraph) ? input.renderedParagraph : null;
	if (hasText(input.selection) && (!renderedParagraph || overlaps(input.selection, renderedParagraph))) {
		return { range: input.selection, source: 'selection' };
	}
	if (hasText(input.preservedSelection) && (!renderedParagraph || overlaps(input.preservedSelection, renderedParagraph))) {
		return { range: input.preservedSelection, source: 'preserved-selection' };
	}
	if (hasText(input.caretParagraph) && (!renderedParagraph || overlaps(input.caretParagraph, renderedParagraph))) {
		return { range: input.caretParagraph, source: 'caret-paragraph' };
	}
	if (renderedParagraph) {
		return { range: renderedParagraph, source: 'rendered-paragraph' };
	}
	return { range: input.selection, source: 'empty-paragraph' };
}

export type DocxFontSizeTargetSource = DocxFormattingTargetSource;
export type ResolveDocxFontSizeTargetInput = ResolveDocxFormattingTargetInput;
export type ResolvedDocxFontSizeTarget = ResolvedDocxFormattingTarget;

export function resolveDocxFontSizeTarget(
	input: ResolveDocxFontSizeTargetInput,
): ResolvedDocxFontSizeTarget {
	return resolveDocxFormattingTarget(input);
}

export function resolveDocxFontSizeStepBase(
	selectionFontSizePoints: readonly number[],
	controlFontSizePoints: number,
): number {
	const sizes = new Set(selectionFontSizePoints.filter((value) => Number.isFinite(value)));
	return sizes.size === 1 ? [...sizes][0]! : controlFontSizePoints;
}
