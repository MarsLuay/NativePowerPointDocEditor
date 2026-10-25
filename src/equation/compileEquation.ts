export type EquationDisplay = 'inline' | 'block';

export interface CompiledEquation {
	display: EquationDisplay;
	ommlXml: string;
	plainText: string;
	source: string;
}

type MathNode =
	| { type: 'row'; children: MathNode[] }
	| { type: 'text'; value: string; roman?: boolean }
	| { type: 'frac'; num: MathNode; den: MathNode }
	| { type: 'sup'; base: MathNode; exp: MathNode }
	| { type: 'sub'; base: MathNode; sub: MathNode }
	| { type: 'subsup'; base: MathNode; sub: MathNode; exp: MathNode }
	| { type: 'sqrt'; radicand: MathNode; index?: MathNode }
	| { type: 'nary'; chr: string; sub?: MathNode; exp?: MathNode; body: MathNode }
	| { type: 'delim'; left: string; right: string; inner: MathNode };

const GREEK: Record<string, string> = {
	alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε',
	zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ',
	lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ',
	sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'ϕ', chi: 'χ',
	psi: 'ψ', omega: 'ω',
	Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π',
	Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
};

const SYMBOLS: Record<string, string> = {
	cdot: '⋅', times: '×', div: '÷', pm: '±', mp: '∓',
	leq: '≤', geq: '≥', neq: '≠', ne: '≠', approx: '≈', equiv: '≡',
	sim: '∼', infty: '∞', partial: '∂', nabla: '∇',
	land: '∧', lor: '∨', neg: '¬', lnot: '¬',
	to: '→', rightarrow: '→', leftarrow: '←', leftrightarrow: '↔',
	Rightarrow: '⇒', Leftarrow: '⇐', Leftrightarrow: '⇔',
	forall: '∀', exists: '∃', in: '∈', notin: '∉',
	subset: '⊂', supset: '⊃', subseteq: '⊆', supseteq: '⊇',
	cup: '∪', cap: '∩', emptyset: '∅', setminus: '∖',
	cdotp: '·', dots: '…', cdots: '⋯', ldots: '…',
	ell: 'ℓ', hbar: 'ℏ', wp: '℘', Re: 'ℜ', Im: 'ℑ',
	angle: '∠', degree: '°',
	quad: '  ', qquad: '    ',
	colon: ':', comma: ',',
};

const NARY: Record<string, string> = {
	sum: '∑', prod: '∏', int: '∫', oint: '∮',
	bigcup: '⋃', bigcap: '⋂', bigvee: '⋁', bigwedge: '⋀',
};

const FUNCTIONS = new Set([
	'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
	'sinh', 'cosh', 'tanh',
	'arcsin', 'arccos', 'arctan',
	'log', 'ln', 'lg', 'exp',
	'max', 'min', 'sup', 'inf', 'lim',
	'dim', 'ker', 'det', 'deg', 'gcd', 'lcm',
	'Pr', 'arg', 'hom',
]);

const BLACKBOARD: Record<string, string> = {
	N: 'ℕ', Z: 'ℤ', Q: 'ℚ', R: 'ℝ', C: 'ℂ',
};

const SUPER: Record<string, string> = {
	'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
	'+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
	n: 'ⁿ', i: 'ⁱ',
};

const SUB: Record<string, string> = {
	'0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
	'+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
	a: 'ₐ', e: 'ₑ', h: 'ₕ', i: 'ᵢ', j: 'ⱼ', k: 'ₖ', l: 'ₗ', m: 'ₘ', n: 'ₙ',
	o: 'ₒ', p: 'ₚ', r: 'ᵣ', s: 'ₛ', t: 'ₜ', u: 'ᵤ', v: 'ᵥ', x: 'ₓ',
};

class ParseError extends Error {}

class LatexParser {
	private readonly source: string;
	private pos = 0;

	constructor(source: string) {
		this.source = source;
	}

	parse(): MathNode {
		const node = this.parseExpr();
		this.skipSpace();
		if (this.pos < this.source.length) {
			throw new ParseError(`Unexpected "${this.source.slice(this.pos, this.pos + 8)}"`);
		}
		return flatten(node);
	}

	private parseExpr(stopRight = false): MathNode {
		const children: MathNode[] = [];
		while (this.pos < this.source.length) {
			const spaced = this.consumeSpace();
			if (this.pos >= this.source.length) break;
			const ch = this.source[this.pos]!;
			if (ch === '}') break;
			if (ch === '&') break;
			if (stopRight && this.peekCommand() === 'right') break;
			if (spaced && children.length > 0) {
				children.push({ type: 'text', value: ' ', roman: true });
			}
			children.push(this.parseScripts(this.parseAtom()));
		}
		if (children.length === 0) return { type: 'text', value: '' };
		if (children.length === 1) return children[0]!;
		return { type: 'row', children };
	}

	private parseAtom(): MathNode {
		this.skipSpace();
		if (this.pos >= this.source.length) return { type: 'text', value: '' };
		const ch = this.source[this.pos]!;
		if (ch === '{') {
			this.pos++;
			const inner = this.parseExpr();
			this.expect('}');
			return inner;
		}
		if (ch === '\\') {
			return this.parseCommand();
		}
		this.pos++;
		if (ch === '/') return { type: 'text', value: '/' };
		return { type: 'text', value: ch };
	}

	private parseCommand(): MathNode {
		this.pos++;
		if (this.pos >= this.source.length) return { type: 'text', value: '\\' };
		const escaped = this.source[this.pos]!;
		if (!/[A-Za-z]/.test(escaped)) {
			this.pos++;
			if (escaped === '\\') return { type: 'text', value: '\n' };
			if (escaped === ' ') return { type: 'text', value: ' ' };
			if (escaped === ',') return { type: 'text', value: ' ' };
			if (escaped === ';') return { type: 'text', value: '  ' };
			if (escaped === '!') return { type: 'text', value: '' };
			return { type: 'text', value: escaped };
		}
		const name = this.readCommandName();
		if (name === 'frac') {
			return { type: 'frac', num: this.parseGroup(), den: this.parseGroup() };
		}
		if (name === 'sqrt') {
			let index: MathNode | undefined;
			this.skipSpace();
			if (this.source[this.pos] === '[') {
				this.pos++;
				index = this.parseExpr();
				this.expect(']');
			}
			return { type: 'sqrt', radicand: this.parseGroup(), ...(index ? { index } : {}) };
		}
		if (name === 'text' || name === 'mathrm' || name === 'textrm' || name === 'operatorname') {
			return romanize(this.parseGroup());
		}
		if (name === 'mathbf' || name === 'boldsymbol') {
			return this.parseGroup();
		}
		if (name === 'mathbb') {
			const inner = serializeUnicode(this.parseGroup());
			return { type: 'text', value: [...inner].map((letter) => BLACKBOARD[letter] ?? letter).join(''), roman: true };
		}
		if (name === 'left') {
			const left = this.readDelimiter();
			const inner = this.parseExpr(true);
			this.skipSpace();
			if (this.peekCommand() === 'right') {
				this.pos++;
				this.readCommandName();
			}
			const right = this.readDelimiter();
			return { type: 'delim', left, right, inner };
		}
		if (name === 'right') {
			return { type: 'text', value: this.readDelimiter() };
		}
		if (name in NARY) {
			const { sub, exp } = this.parseLimitScripts();
			this.skipSpace();
			const body = this.pos < this.source.length && this.source[this.pos] !== '}' && this.peekCommand() !== 'right'
				? this.parseScripts(this.parseAtom())
				: { type: 'text' as const, value: '' };
			return { type: 'nary', chr: NARY[name]!, body, ...(sub ? { sub } : {}), ...(exp ? { exp } : {}) };
		}
		if (FUNCTIONS.has(name)) {
			return { type: 'text', value: name, roman: true };
		}
		if (name in GREEK) {
			return { type: 'text', value: GREEK[name]! };
		}
		if (name in SYMBOLS) {
			return { type: 'text', value: SYMBOLS[name]!, roman: true };
		}
		if (name === 'overline') {
			return { type: 'row', children: [this.parseGroup(), { type: 'text', value: '̅' }] };
		}
		if (name === 'underline') {
			return this.parseGroup();
		}
		return { type: 'text', value: `\\${name}` };
	}

	private parseScripts(base: MathNode): MathNode {
		let sub: MathNode | undefined;
		let exp: MathNode | undefined;
		for (;;) {
			const saved = this.pos;
			this.skipSpace();
			const ch = this.source[this.pos];
			if (ch === '_') {
				this.pos++;
				sub = this.parseAtom();
				continue;
			}
			if (ch === '^') {
				this.pos++;
				exp = this.parseAtom();
				continue;
			}
			this.pos = saved;
			break;
		}
		if (sub && exp) return { type: 'subsup', base, sub, exp };
		if (sub) return { type: 'sub', base, sub };
		if (exp) return { type: 'sup', base, exp };
		return base;
	}

	private parseLimitScripts(): { sub?: MathNode; exp?: MathNode } {
		let sub: MathNode | undefined;
		let exp: MathNode | undefined;
		for (;;) {
			const saved = this.pos;
			this.skipSpace();
			const ch = this.source[this.pos];
			if (ch === '_') {
				this.pos++;
				sub = this.parseAtom();
				continue;
			}
			if (ch === '^') {
				this.pos++;
				exp = this.parseAtom();
				continue;
			}
			this.pos = saved;
			break;
		}
		return { ...(sub ? { sub } : {}), ...(exp ? { exp } : {}) };
	}

	private parseGroup(): MathNode {
		this.skipSpace();
		if (this.source[this.pos] === '{') {
			this.pos++;
			const inner = this.parseExpr();
			this.expect('}');
			return inner;
		}
		return this.parseAtom();
	}

	private readCommandName(): string {
		const start = this.pos;
		while (this.pos < this.source.length && /[A-Za-z]/.test(this.source[this.pos]!)) {
			this.pos++;
		}
		return this.source.slice(start, this.pos);
	}

	private peekCommand(): string | null {
		this.skipSpace();
		if (this.source[this.pos] !== '\\') return null;
		let index = this.pos + 1;
		if (index >= this.source.length || !/[A-Za-z]/.test(this.source[index]!)) return null;
		const start = index;
		while (index < this.source.length && /[A-Za-z]/.test(this.source[index]!)) index++;
		return this.source.slice(start, index);
	}

	private readDelimiter(): string {
		this.skipSpace();
		if (this.pos >= this.source.length) return '';
		if (this.source[this.pos] === '\\') {
			this.pos++;
			const name = this.readCommandName();
			if (name === 'lvert' || name === '|') return '|';
			if (name === 'lVert' || name === 'rVert') return '∥';
			if (name === 'langle') return '⟨';
			if (name === 'rangle') return '⟩';
			if (name === 'lfloor') return '⌊';
			if (name === 'rfloor') return '⌋';
			if (name === 'lceil') return '⌈';
			if (name === 'rceil') return '⌉';
			if (name === '{') return '{';
			if (name === '}') return '}';
			if (!name) {
				const ch = this.source[this.pos] ?? '';
				this.pos++;
				return ch === '.' ? '' : ch;
			}
			return name;
		}
		const ch = this.source[this.pos]!;
		this.pos++;
		return ch === '.' ? '' : ch;
	}

	private expect(ch: string): void {
		this.skipSpace();
		if (this.source[this.pos] !== ch) {
			throw new ParseError(`Expected "${ch}"`);
		}
		this.pos++;
	}

	private skipSpace(): void {
		this.consumeSpace();
	}

	private consumeSpace(): boolean {
		const start = this.pos;
		while (this.pos < this.source.length && /\s/.test(this.source[this.pos]!)) {
			this.pos++;
		}
		return this.pos > start;
	}
}

function flatten(node: MathNode): MathNode {
	if (node.type !== 'row') return node;
	const children: MathNode[] = [];
	for (const child of node.children) {
		const flat = flatten(child);
		if (flat.type === 'row') children.push(...flat.children);
		else children.push(flat);
	}
	if (children.length === 1) return children[0]!;
	return { type: 'row', children };
}

function romanize(node: MathNode): MathNode {
	if (node.type === 'text') return { ...node, roman: true };
	if (node.type === 'row') return { type: 'row', children: node.children.map(romanize) };
	return node;
}

function looksLikeLatex(source: string): boolean {
	return /\\[A-Za-z]+|[_^{}]/.test(source);
}

function unwrapMathDelimiters(source: string): { body: string; display: EquationDisplay } {
	const displayMatch = source.match(/^\$\$([\s\S]*)\$\$$/) ?? source.match(/^\\\[([\s\S]*)\\\]$/);
	if (displayMatch) {
		return { body: displayMatch[1]!.trim(), display: 'block' };
	}
	const inlineMatch = source.match(/^\$([\s\S]*)\$$/) ?? source.match(/^\\\(([\s\S]*)\\\)$/);
	if (inlineMatch) {
		return { body: inlineMatch[1]!.trim(), display: 'inline' };
	}
	return { body: source, display: 'inline' };
}

function escapeXml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;');
}

function ommlText(value: string, roman = false): string {
	const style = roman ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : '';
	const space = /\s/.test(value) ? ' xml:space="preserve"' : '';
	return `<m:r>${style}<m:t${space}>${escapeXml(value)}</m:t></m:r>`;
}

function ommlOf(node: MathNode): string {
	switch (node.type) {
		case 'text':
			return node.value ? ommlText(node.value, node.roman) : '';
		case 'row':
			return node.children.map(ommlOf).join('');
		case 'frac':
			return `<m:f><m:num>${ommlOf(node.num)}</m:num><m:den>${ommlOf(node.den)}</m:den></m:f>`;
		case 'sup':
			return `<m:sSup><m:e>${ommlOf(node.base)}</m:e><m:sup>${ommlOf(node.exp)}</m:sup></m:sSup>`;
		case 'sub':
			return `<m:sSub><m:e>${ommlOf(node.base)}</m:e><m:sub>${ommlOf(node.sub)}</m:sub></m:sSub>`;
		case 'subsup':
			return `<m:sSubSup><m:e>${ommlOf(node.base)}</m:e><m:sub>${ommlOf(node.sub)}</m:sub><m:sup>${ommlOf(node.exp)}</m:sup></m:sSubSup>`;
		case 'sqrt':
			return node.index
				? `<m:rad><m:deg>${ommlOf(node.index)}</m:deg><m:e>${ommlOf(node.radicand)}</m:e></m:rad>`
				: `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${ommlOf(node.radicand)}</m:e></m:rad>`;
		case 'nary':
			return `<m:nary><m:naryPr><m:chr m:val="${escapeXml(node.chr)}"/><m:limLoc m:val="undOvr"/></m:naryPr><m:sub>${node.sub ? ommlOf(node.sub) : ''}</m:sub><m:sup>${node.exp ? ommlOf(node.exp) : ''}</m:sup><m:e>${ommlOf(node.body)}</m:e></m:nary>`;
		case 'delim':
			return `<m:d><m:dPr><m:begChr m:val="${escapeXml(node.left)}"/><m:endChr m:val="${escapeXml(node.right)}"/></m:dPr><m:e>${ommlOf(node.inner)}</m:e></m:d>`;
	}
}

function mapScripts(value: string, table: Record<string, string>): string | null {
	let mapped = '';
	for (const letter of value) {
		const next = table[letter];
		if (!next) return null;
		mapped += next;
	}
	return mapped;
}

function serializeUnicode(node: MathNode): string {
	switch (node.type) {
		case 'text':
			return node.value;
		case 'row':
			return node.children.map(serializeUnicode).join('');
		case 'frac':
			return `(${serializeUnicode(node.num)})/(${serializeUnicode(node.den)})`;
		case 'sup': {
			const exp = serializeUnicode(node.exp);
			return `${serializeUnicode(node.base)}${mapScripts(exp, SUPER) ?? `^${exp}`}`;
		}
		case 'sub': {
			const sub = serializeUnicode(node.sub);
			return `${serializeUnicode(node.base)}${mapScripts(sub, SUB) ?? `_${sub}`}`;
		}
		case 'subsup': {
			const sub = serializeUnicode(node.sub);
			const exp = serializeUnicode(node.exp);
			return `${serializeUnicode(node.base)}${mapScripts(sub, SUB) ?? `_${sub}`}${mapScripts(exp, SUPER) ?? `^${exp}`}`;
		}
		case 'sqrt':
			return node.index
				? `√[${serializeUnicode(node.index)}](${serializeUnicode(node.radicand)})`
				: `√(${serializeUnicode(node.radicand)})`;
		case 'nary': {
			const sub = node.sub ? `_${serializeUnicode(node.sub)}` : '';
			const exp = node.exp ? `^${serializeUnicode(node.exp)}` : '';
			return `${node.chr}${sub}${exp} ${serializeUnicode(node.body)}`.trim();
		}
		case 'delim':
			return `${node.left}${serializeUnicode(node.inner)}${node.right}`;
	}
}

function wrapOmml(inner: string, display: EquationDisplay): string {
	const math = `<m:oMath>${inner}</m:oMath>`;
	if (display === 'block') {
		return `<m:oMathPara>${math}</m:oMathPara>`;
	}
	return math;
}

export function compileEquation(source: string, display?: EquationDisplay): CompiledEquation {
	const trimmed = source.trim();
	if (!trimmed) {
		throw new Error('Equation is empty.');
	}

	const normalized = unwrapMathDelimiters(trimmed);
	const resolvedDisplay = display ?? normalized.display;
	let ast: MathNode;
	if (looksLikeLatex(normalized.body)) {
		try {
			ast = new LatexParser(normalized.body).parse();
		} catch {
			ast = { type: 'text', value: normalized.body };
		}
	} else {
		ast = { type: 'text', value: normalized.body };
	}

	const inner = ommlOf(ast) || ommlText(normalized.body);
	const plainText = serializeUnicode(ast) || normalized.body;
	return {
		display: resolvedDisplay,
		ommlXml: wrapOmml(inner, resolvedDisplay),
		plainText,
		source: trimmed,
	};
}
