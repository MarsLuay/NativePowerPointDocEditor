import type { MathSpanDisplay } from './mathSpans';

const LATEX_MARKER_RE = /<!--npde-latex:(inline|block):([\s\S]*?)-->/;
const LATEX_MARKER_GLOBAL_RE = /<!--npde-latex:(?:inline|block):[\s\S]*?-->/g;

const HIDDEN_START = '\u2064\u2064';
const HIDDEN_END = '\u2062\u2062';
const HIDDEN_ZERO = '\u2060';
const HIDDEN_ONE = '\u2063';

export interface HiddenMathRegion {
	from: number;
	to: number;
	source: string;
	display: MathSpanDisplay;
}

export function latexBodyFromSource(source: string): string {
	const trimmed = source.trim();
	if (trimmed.startsWith('$$') && trimmed.endsWith('$$') && trimmed.length >= 4) {
		return trimmed.slice(2, -2);
	}
	if (trimmed.startsWith('\\[') && trimmed.endsWith('\\]') && trimmed.length >= 4) {
		return trimmed.slice(2, -2);
	}
	if (trimmed.startsWith('\\(') && trimmed.endsWith('\\)') && trimmed.length >= 4) {
		return trimmed.slice(2, -2);
	}
	if (trimmed.startsWith('$') && trimmed.endsWith('$') && trimmed.length >= 2) {
		return trimmed.slice(1, -1);
	}
	return trimmed;
}

export function stripLatexSource(ommlXml: string): string {
	return ommlXml.replace(LATEX_MARKER_GLOBAL_RE, '');
}

export function embedLatexSource(ommlXml: string, source: string, display: MathSpanDisplay): string {
	const body = latexBodyFromSource(source);
	const cleaned = stripLatexSource(ommlXml);
	if (!body) {
		return cleaned;
	}
	const marker = `<!--npde-latex:${display}:${encodeURIComponent(body)}-->`;
	const gt = cleaned.indexOf('>');
	if (gt < 0) {
		return `${marker}${cleaned}`;
	}
	return `${cleaned.slice(0, gt + 1)}${marker}${cleaned.slice(gt + 1)}`;
}

export function extractLatexSource(ommlXml: string): { source: string; display: MathSpanDisplay } | null {
	const match = LATEX_MARKER_RE.exec(ommlXml);
	if (!match) {
		return null;
	}
	try {
		return {
			source: decodeURIComponent(match[2] ?? ''),
			display: match[1] === 'block' ? 'block' : 'inline',
		};
	} catch {
		return null;
	}
}

export function mathInnerForEdit(
	ommlXml: string,
	plainText: string,
	display: MathSpanDisplay,
): { inner: string; display: MathSpanDisplay } {
	const extracted = extractLatexSource(ommlXml);
	if (extracted?.source) {
		return { inner: extracted.source, display: extracted.display };
	}
	return { inner: plainText || '', display };
}

export function wrapLatexForEdit(inner: string, display: MathSpanDisplay): string {
	if (display === 'block') {
		return `$$${inner}$$`;
	}
	return inner ? `$${inner}$` : '$ $';
}

export function latexCaretOffsets(
	wrapped: string,
	display: MathSpanDisplay,
): { start: number; end: number } {
	const open = display === 'block' ? 2 : 1;
	return { start: open, end: Math.max(open, wrapped.length - open) };
}

export function encodeHiddenMath(
	display: MathSpanDisplay,
	source: string,
	plainText: string,
): string {
	const payload = `${display === 'block' ? 'b' : 'i'}:${plainText.length}:${encodeURIComponent(source)}`;
	const bytes = new TextEncoder().encode(payload);
	let bits = '';
	for (const byte of bytes) {
		for (let bit = 7; bit >= 0; bit -= 1) {
			bits += (byte >> bit) & 1 ? HIDDEN_ONE : HIDDEN_ZERO;
		}
	}
	return `${HIDDEN_START}${bits}${HIDDEN_END}`;
}

export function findHiddenMathRegions(text: string): HiddenMathRegion[] {
	const regions: HiddenMathRegion[] = [];
	let index = 0;
	while (index < text.length) {
		const start = text.indexOf(HIDDEN_START, index);
		if (start < 0) {
			break;
		}
		const end = text.indexOf(HIDDEN_END, start + HIDDEN_START.length);
		if (end < 0) {
			break;
		}
		const payload = decodeHiddenBits(text.slice(start + HIDDEN_START.length, end));
		index = end + HIDDEN_END.length;
		if (!payload) {
			continue;
		}
		const parsed = parseHiddenPayload(payload);
		if (!parsed) {
			continue;
		}
		const from = start - parsed.visibleLen;
		if (from < 0) {
			continue;
		}
		regions.push({
			from,
			to: index,
			source: parsed.source,
			display: parsed.display,
		});
	}
	return regions;
}

export function expandPlainTextMath(
	text: string,
	caret: number,
	options: { key?: string; fromPointer?: boolean } = {},
): { text: string; caret: number } | null {
	const region = findHiddenMathRegions(text).find((item) => {
		if (options.key === 'ArrowRight') {
			return caret === item.from;
		}
		if (options.key === 'ArrowLeft') {
			return caret === item.to;
		}
		if (options.fromPointer) {
			return caret >= item.from && caret < item.to;
		}
		return caret > item.from && caret < item.to;
	});
	if (!region) {
		return null;
	}
	const wrapped = wrapLatexForEdit(region.source, region.display);
	const offsets = latexCaretOffsets(wrapped, region.display);
	const caretAtEnd = options.key === 'ArrowLeft';
	return {
		text: `${text.slice(0, region.from)}${wrapped}${text.slice(region.to)}`,
		caret: region.from + (caretAtEnd ? offsets.end : offsets.start),
	};
}

function decodeHiddenBits(bits: string): string | null {
	if (!bits || bits.length % 8 !== 0) {
		return null;
	}
	const bytes = new Uint8Array(bits.length / 8);
	for (let index = 0; index < bytes.length; index += 1) {
		let value = 0;
		for (let bit = 0; bit < 8; bit += 1) {
			const ch = bits[index * 8 + bit];
			if (ch === HIDDEN_ONE) {
				value = (value << 1) | 1;
			} else if (ch === HIDDEN_ZERO) {
				value <<= 1;
			} else {
				return null;
			}
		}
		bytes[index] = value;
	}
	try {
		return new TextDecoder().decode(bytes);
	} catch {
		return null;
	}
}

function parseHiddenPayload(
	payload: string,
): { display: MathSpanDisplay; visibleLen: number; source: string } | null {
	const match = payload.match(/^([ib]):(\d+):(.*)$/);
	if (!match) {
		return null;
	}
	const visibleLen = Number(match[2]);
	if (!Number.isFinite(visibleLen) || visibleLen < 0) {
		return null;
	}
	try {
		return {
			display: match[1] === 'b' ? 'block' : 'inline',
			visibleLen,
			source: decodeURIComponent(match[3] ?? ''),
		};
	} catch {
		return null;
	}
}
