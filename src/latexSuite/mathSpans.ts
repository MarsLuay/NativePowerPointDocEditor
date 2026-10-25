export type MathSpanDisplay = 'inline' | 'block';

export interface MathSpan {
	from: number;
	to: number;
	inner: string;
	display: MathSpanDisplay;
}

export interface MathCommitOptions {
	commitReadyAtCaret?: boolean;
	commitNonEmptyAtCaret?: boolean;
}

export function isEmptyMath(inner: string): boolean {
	return inner.replace(/[{}\s]/g, '').length === 0;
}

/**
 * Drop the empty `{ }` shells LaTeX Suite's `mk` snippet wraps around the body.
 */
export function unwrapLatexSuiteInner(inner: string): string {
	let next = inner.trim();
	if (next.startsWith('{}')) {
		next = next.slice(2);
	}
	if (next.endsWith('{}')) {
		next = next.slice(0, -2);
	}
	return next.trim();
}

/**
 * Closed `$...$` / `$$...$$` / `\(...\)` / `\[...\]` spans.
 * Empty `{ }` shells are skipped so `mk` stays editable.
 */
export function findMathSpans(text: string): MathSpan[] {
	const spans: MathSpan[] = [];
	let index = 0;
	while (index < text.length) {
		if (text.startsWith('$$', index)) {
			const end = text.indexOf('$$', index + 2);
			if (end < 0) break;
			pushSpan(spans, index, end + 2, text.slice(index + 2, end), 'block');
			index = end + 2;
			continue;
		}
		if (text.startsWith('\\[', index)) {
			const end = text.indexOf('\\]', index + 2);
			if (end < 0) break;
			pushSpan(spans, index, end + 2, text.slice(index + 2, end), 'block');
			index = end + 2;
			continue;
		}
		if (text.startsWith('\\(', index)) {
			const end = text.indexOf('\\)', index + 2);
			if (end < 0) break;
			pushSpan(spans, index, end + 2, text.slice(index + 2, end), 'inline');
			index = end + 2;
			continue;
		}
		if (text[index] === '$') {
			const end = text.indexOf('$', index + 1);
			if (end < 0) break;
			pushSpan(spans, index, end + 1, text.slice(index + 1, end), 'inline');
			index = end + 1;
			continue;
		}
		index += 1;
	}
	return spans;
}

/**
 * Closed math spans the caret is not inside.
 * Used to turn finished LaTeX Suite math into OMML or Cambria Math.
 */
export function findClosedMathSpans(text: string, cursor: number): MathSpan[] {
	return findMathSpans(text).filter((span) => cursor <= span.from || cursor >= span.to);
}

/**
 * True when leftover inner text is only the `mk` trailing `{ }` and the closer.
 * Tab then commits `${}\cdot{}$` without waiting to leave the `$...$` span.
 */
export function isMathReadyToCommit(text: string, span: MathSpan, cursor: number): boolean {
	if (isEmptyMath(span.inner)) {
		return false;
	}
	if (cursor <= span.from || cursor >= span.to) {
		return true;
	}
	let remainder = text.slice(cursor, span.to);
	remainder = remainder.replace(/\{\s*\}(\$\$|\\\)|\\\]|\$)$/, '$1');
	remainder = remainder.replace(/\s+/g, '');
	return /^(?:\$\$|\\\)|\\\]|\$)?$/.test(remainder);
}

export function findMathSpansToCommit(
	text: string,
	cursor: number,
	options: MathCommitOptions = {},
): MathSpan[] {
	const spans = findClosedMathSpans(text, cursor);
	const current = findMathSpans(text).find((span) => cursor > span.from && cursor < span.to);
	if (!current) {
		return spans;
	}
	const include = options.commitNonEmptyAtCaret
		|| (Boolean(options.commitReadyAtCaret) && isMathReadyToCommit(text, current, cursor));
	if (include && !spans.some((span) => span.from === current.from && span.to === current.to)) {
		return [...spans, current];
	}
	return spans;
}

function pushSpan(
	spans: MathSpan[],
	from: number,
	to: number,
	inner: string,
	display: MathSpanDisplay,
): void {
	if (isEmptyMath(inner)) return;
	spans.push({ from, to, inner, display });
}
