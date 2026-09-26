import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import path from 'node:path';

const styles = readFileSync(path.resolve(import.meta.dirname, '..', 'styles.css'), 'utf8');
const scrollContainer = '[data-native-powerpoint-doc-editor-scroll-container]';

test('DOCX scroll chrome hides only the vertical WebKit scrollbar', () => {
	const scrollbarRule = new RegExp(`${escapeRegExp(scrollContainer)}::-webkit-scrollbar\\s*\\{[^}]*\\}`, 's').exec(styles)?.[0] ?? '';
	assert.match(scrollbarRule, /width:\s*0/);
	assert.match(scrollbarRule, /height:\s*6px/);
	assert.match(styles, new RegExp(`${escapeRegExp(scrollContainer)}::-webkit-scrollbar:vertical\\s*\\{[^}]*width:\\s*0`, 's'));
	assert.match(styles, new RegExp(`${escapeRegExp(scrollContainer)}::-webkit-scrollbar:horizontal\\s*\\{[^}]*height:\\s*6px`, 's'));
	assert.match(styles, new RegExp(`${escapeRegExp(scrollContainer)}::-webkit-scrollbar-thumb:horizontal\\s*\\{[^}]*--npde-docx-document-scrollbar-thumb`, 's'));
	assert.doesNotMatch(styles, new RegExp(`${escapeRegExp(scrollContainer)}[^{}]*\\{[^}]*overflow-y\\s*:\s*hidden`, 's'));
});

function escapeRegExp(value) {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
