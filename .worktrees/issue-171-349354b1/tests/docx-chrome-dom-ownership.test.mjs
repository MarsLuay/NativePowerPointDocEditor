import assert from 'node:assert/strict';
import test from 'node:test';

import { createDetachedDocxEditorChromeElement } from '../src/docxEditorChromeDom.ts';

test('DOCX chrome fallback creation avoids Obsidian Document helpers', () => {
	const created = { tagName: 'DIV' };
	let createElementCalls = 0;
	let createDivCalls = 0;
	const ownerDocument = {
		createElement(tagName) {
			createElementCalls += 1;
			assert.equal(tagName, 'div');
			return created;
		},
		createDiv() {
			createDivCalls += 1;
			throw new Error('Only one element on document allowed.');
		},
	};

	const element = createDetachedDocxEditorChromeElement({ ownerDocument }, 'div');

	assert.equal(element, created);
	assert.equal(createElementCalls, 1);
	assert.equal(createDivCalls, 0);
});
