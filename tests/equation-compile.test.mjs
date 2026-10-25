import assert from 'node:assert/strict';
import test from 'node:test';

import { compileEquation } from '../src/equation/compileEquation.ts';

test('compiles a LaTeX fraction to OMML and Unicode', () => {
	const equation = compileEquation('\\frac{a}{b}');
	assert.equal(equation.display, 'inline');
	assert.match(equation.ommlXml, /<m:oMath>/);
	assert.match(equation.ommlXml, /<m:f>/);
	assert.match(equation.ommlXml, /<m:num>.*a.*<\/m:num>/s);
	assert.match(equation.ommlXml, /<m:den>.*b.*<\/m:den>/s);
	assert.doesNotMatch(equation.ommlXml, /<m:oMathPara>/);
	assert.equal(equation.plainText, '(a)/(b)');
});

test('wraps $$...$$ as a display equation', () => {
	const equation = compileEquation('$$x^2 + y^2 = z^2$$');
	assert.equal(equation.display, 'block');
	assert.match(equation.ommlXml, /<m:oMathPara>/);
	assert.match(equation.ommlXml, /<m:sSup>/);
	assert.equal(equation.plainText, 'x² + y² = z²');
	assert.match(equation.ommlXml, /xml:space="preserve"/);
});

test('compiles logic operators used in discrete math', () => {
	const equation = compileEquation('\\forall x \\in \\mathbb{N} (P(x) \\rightarrow Q(x))');
	assert.match(equation.plainText, /∀/);
	assert.match(equation.plainText, /∈/);
	assert.match(equation.plainText, /ℕ/);
	assert.match(equation.plainText, /→/);
	assert.match(equation.ommlXml, /<m:t[^>]*>∀<\/m:t>/);
});

test('compiles sums, roots, and delimiters', () => {
	const equation = compileEquation('\\sum_{i=1}^{n} \\sqrt{i}');
	assert.match(equation.ommlXml, /<m:nary>/);
	assert.match(equation.ommlXml, /m:val="∑"/);
	assert.match(equation.ommlXml, /<m:rad>/);
	assert.match(equation.plainText, /∑/);
	assert.match(equation.plainText, /√/);
});

test('plain Unicode math stays a single OMML run', () => {
	const equation = compileEquation('∀x∈ℕ', 'block');
	assert.equal(equation.display, 'block');
	assert.equal(equation.plainText, '∀x∈ℕ');
	assert.match(equation.ommlXml, /<m:oMathPara>/);
	assert.match(equation.ommlXml, /∀x∈ℕ/);
});

test('rejects empty input', () => {
	assert.throws(() => compileEquation('   '), /empty/i);
});
