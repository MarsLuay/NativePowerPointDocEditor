import assert from 'node:assert/strict';
import test from 'node:test';

import { compileEquation } from '../src/equation/compileEquation.ts';
import {
	embedLatexSource,
	encodeHiddenMath,
	expandPlainTextMath,
	extractLatexSource,
	findHiddenMathRegions,
	latexBodyFromSource,
	wrapLatexForEdit,
} from '../src/latexSuite/mathSource.ts';

function commitVisible(source, display = 'inline') {
	const equation = compileEquation(source, display);
	const inner = latexBodyFromSource(equation.source);
	return `${equation.plainText}${encodeHiddenMath(display, inner, equation.plainText)}`;
}

test('stores LaTeX in OMML so a math node can expand back to source', () => {
	const equation = compileEquation('\\cdot', 'inline');
	const omml = embedLatexSource(equation.ommlXml, equation.source, equation.display);
	assert.match(omml, /<!--npde-latex:inline:/);
	assert.equal(extractLatexSource(omml)?.source, '\\cdot');
	assert.equal(wrapLatexForEdit('\\cdot', 'inline'), '$\\cdot$');
	assert.equal(latexBodyFromSource('$\\cdot$'), '\\cdot');
	assert.equal(wrapLatexForEdit('x^2 + y^2', 'block'), '$$x^2 + y^2$$');
});

test('PowerPoint commit hides source and click/arrow restores $...$', () => {
	const committed = commitVisible('\\cdot');
	assert.match(committed, /⋅/);
	assert.equal(committed.includes('\\cdot'), false);
	assert.equal(findHiddenMathRegions(committed).length, 1);
	assert.equal(findHiddenMathRegions(committed)[0]?.source, '\\cdot');

	const fromRight = expandPlainTextMath(committed, committed.length, { key: 'ArrowLeft' });
	assert.equal(fromRight?.text, '$\\cdot$');
	assert.equal(fromRight?.caret, '$\\cdot$'.length - 1);

	const fromLeft = expandPlainTextMath(committed, committed.indexOf('⋅'), { key: 'ArrowRight' });
	assert.equal(fromLeft?.text, '$\\cdot$');
	assert.equal(fromLeft?.caret, 1);

	const clicked = expandPlainTextMath(committed, committed.indexOf('⋅') + 1, { fromPointer: true });
	assert.equal(clicked?.text, '$\\cdot$');
});

test('caret after a committed glyph does not expand until arrow or click', () => {
	const committed = commitVisible('\\cdot');
	assert.equal(expandPlainTextMath(committed, committed.length), null);
	const insideHidden = committed.indexOf('⋅') + 1;
	assert.equal(expandPlainTextMath(committed, insideHidden)?.text, '$\\cdot$');
});

test('frac source round-trips through the hidden payload', () => {
	const committed = commitVisible('\\frac{a}{b}');
	const expanded = expandPlainTextMath(committed, committed.length, { key: 'ArrowLeft' });
	assert.equal(expanded?.text, '$\\frac{a}{b}$');
});
