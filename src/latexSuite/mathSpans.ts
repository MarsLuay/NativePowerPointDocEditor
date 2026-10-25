export type MathSpanDisplay = 'inline' | 'block';

export interface MathSpan {
	from: number;
	to: number;
	inner: string;
	display: MathSpanDisplay;
}

function isEmptyMath(inner: string): boolean {
	return inner.replace(/[{}\s]/g, '').length === 0;
}

/**
 * Closed `$...$` / `$$...$$` / `\(...\)` / `\[...\]` spans the caret is not inside.
 * Used to turn finished LaTeX Suite math into OMML or Cambria Math.
 */
export function findClosedMathSpans(text: string, cursor: number): MathSpan[] {
	const spans: MathSpan[] = [];
	let index = 0;
	while (index < text.length) {
		if (text.startsWith('$$', index)) {
			const end = text.indexOf('$$', index + 2);
			if (end < 0) break;
			pushSpan(spans, cursor, index, end + 2, text.slice(index + 2, end), 'block');
			index = end + 2;
			continue;
		}
		if (text.startsWith('\\[', index)) {
			const end = text.indexOf('\\]', index + 2);
			if (end < 0) break;
			pushSpan(spans, cursor, index, end + 2, text.slice(index + 2, end), 'block');
			index = end + 2;
			continue;
		}
		if (text.startsWith('\\(', index)) {
			const end = text.indexOf('\\)', index + 2);
			if (end < 0) break;
			pushSpan(spans, cursor, index, end + 2, text.slice(index + 2, end), 'inline');
			index = end + 2;
			continue;
		}
		if (text[index] === '$') {
			const end = text.indexOf('$', index + 1);
			if (end < 0) break;
			pushSpan(spans, cursor, index, end + 1, text.slice(index + 1, end), 'inline');
			index = end + 1;
			continue;
		}
		index += 1;
	}
	return spans;
}

function pushSpan(
	spans: MathSpan[],
	cursor: number,
	from: number,
	to: number,
	inner: string,
	display: MathSpanDisplay,
): void {
	if (isEmptyMath(inner)) return;
	if (cursor > from && cursor < to) return;
	spans.push({ from, to, inner, display });
}
