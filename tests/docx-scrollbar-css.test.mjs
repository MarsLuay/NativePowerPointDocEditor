import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import path from 'node:path';

const styles = readFileSync(path.resolve(import.meta.dirname, '..', 'styles.css'), 'utf8');
const docxReactView = readFileSync(path.resolve(import.meta.dirname, '..', 'src', 'DocxReactView.tsx'), 'utf8');
const scrollContainer = '[data-native-powerpoint-doc-editor-scroll-container]';
const touchOnlyScrollbarClass = '.native-powerpoint-doc-editor-touch-only-scrollbar';

test('DOCX scroll chrome hides only the vertical WebKit scrollbar', () => {
	const scrollbarRule = new RegExp(`${escapeRegExp(scrollContainer)}::-webkit-scrollbar\\s*\\{[^}]*\\}`, 's').exec(styles)?.[0] ?? '';
	assert.match(scrollbarRule, /width:\s*0/);
	assert.match(scrollbarRule, /height:\s*6px/);
	assert.match(styles, new RegExp(`${escapeRegExp(scrollContainer)}::-webkit-scrollbar:vertical\\s*\\{[^}]*width:\\s*0`, 's'));
	assert.match(styles, new RegExp(`${escapeRegExp(scrollContainer)}::-webkit-scrollbar:horizontal\\s*\\{[^}]*height:\\s*6px`, 's'));
	assert.match(styles, new RegExp(`${escapeRegExp(scrollContainer)}::-webkit-scrollbar-thumb:horizontal\\s*\\{[^}]*--npde-docx-document-scrollbar-thumb`, 's'));
	assert.doesNotMatch(styles, new RegExp(`${escapeRegExp(scrollContainer)}[^{}]*\\{[^}]*overflow-y\\s*:\s*hidden`, 's'));
});

test('touch-only mobile styling hides only the DOCX horizontal track', () => {
	const trackSelector = `${escapeRegExp(touchOnlyScrollbarClass)}[^{}]*${escapeRegExp(scrollContainer)}::-webkit-scrollbar-track:horizontal`;
	assert.match(styles, new RegExp(`${trackSelector}\\s*\\{[^}]*background-color:\\s*transparent`, 's'));
	assert.match(styles, new RegExp(`${escapeRegExp(touchOnlyScrollbarClass)}[^{}]*${escapeRegExp(scrollContainer)}::-webkit-scrollbar:horizontal`));
	assert.match(styles, new RegExp(`${escapeRegExp(touchOnlyScrollbarClass)}[^{}]*${escapeRegExp(scrollContainer)}::-webkit-scrollbar-track:horizontal`));
	assert.doesNotMatch(styles, new RegExp(`${escapeRegExp(touchOnlyScrollbarClass)}[^{}]*formatting-bar`));
	assert.doesNotMatch(styles, new RegExp(`${escapeRegExp(touchOnlyScrollbarClass)}[^{}]*native-powerpoint-canvas-pane`));
});

test('touch-only scrollbar styling requires mobile platform and live input capability checks', () => {
	assert.match(docxReactView, /Platform\.isMobile/);
	assert.match(docxReactView, /Platform\.isMobileApp/);
	assert.match(docxReactView, /maxTouchPoints >= 1/);
	assert.match(docxReactView, /matchMedia\('\(hover: none\)'\)/);
	assert.match(docxReactView, /matchMedia\('\(pointer: coarse\)'\)/);
	assert.match(docxReactView, /addEventListener\('change'/);
	assert.match(docxReactView, /removeEventListener\('change'/);
	assert.match(docxReactView, /!Platform\.isMobile && !Platform\.isMobileApp/);
});

function escapeRegExp(value) {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
