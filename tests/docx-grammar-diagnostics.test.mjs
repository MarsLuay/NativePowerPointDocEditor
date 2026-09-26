import assert from 'node:assert/strict';
import test from 'node:test';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { createRequire } from 'node:module';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
const grammar = require(await bundleSource('src/harper/docxGrammarDiagnostics.ts', 'docx-grammar-diagnostics.cjs'));

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'text*', group: 'block' },
		text: { group: 'inline' },
	},
	marks: {
		bold: {},
	},
});

function paragraph(text, marks) {
	return schema.node('paragraph', null, text ? [schema.text(text, marks)] : []);
}

test('only the changed paragraph is linted, and a long paragraph stays inside the window', () => {
	const before = schema.node('doc', null, [
		paragraph('Alpha'),
		paragraph('Bravo stays'),
	]);
	const after = schema.node('doc', null, [
		paragraph('Alpha!'),
		paragraph('Bravo stays'),
	]);
	const windows = grammar.collectChangedTextWindows(before, after);
	assert.equal(windows.length, 1);
	assert.equal(windows[0].text, 'Alpha!');

	const long = `${'a'.repeat(2500)}changed`;
	const longBefore = schema.node('doc', null, [paragraph('a'.repeat(2500))]);
	const longAfter = schema.node('doc', null, [paragraph(long)]);
	const [window] = grammar.collectChangedTextWindows(longBefore, longAfter);
	assert.ok(window.text.length <= grammar.DOCX_GRAMMAR_TEXT_WINDOW);
	assert.ok(window.text.includes('changed'));
	assert.ok(window.textOffset > 0);
});

test('Harper spans map onto marked text and a suggestion keeps formatting', () => {
	const bold = schema.marks.bold.create();
	const doc = schema.node('doc', null, [
		schema.node('paragraph', null, [
			schema.text('Hello '),
			schema.text('wrld', [bold]),
		]),
	]);
	const block = doc.firstChild;
	const window = { blockPos: 0, textOffset: 0, text: 'Hello wrld' };
	const mapped = grammar.mapWindowSpan(doc, window, { start: 6, end: 10 });
	assert.equal(doc.textBetween(mapped.from, mapped.to), 'wrld');
	const state = EditorState.create({ schema, doc });
	const transaction = grammar.applyDocxGrammarSuggestion(state, mapped.from, mapped.to, {
		kind: 'replace',
		replacement: 'world',
	});
	const next = state.apply(transaction);
	assert.equal(next.doc.textContent, 'Hello world');
	let boldText = '';
	next.doc.descendants((node) => {
		if (node.isText && bold.isInSet(node.marks)) boldText += node.text;
	});
	assert.equal(boldText, 'world');
	assert.equal(block.type.name, 'paragraph');
});

test('stale grammar results and composition do not update diagnostics', () => {
	assert.equal(grammar.acceptGrammarResult(2, 2), true);
	assert.equal(grammar.acceptGrammarResult(2, 3), false);
	assert.equal(grammar.grammarLintAllowed({ enabled: true, composing: true }), false);
	assert.equal(grammar.grammarLintAllowed({ enabled: false, composing: false }), false);
	const doc = schema.node('doc', null, [paragraph('Hello wrld')]);
	const window = { blockPos: 0, textOffset: 0, text: 'Hello wrld' };
	const current = grammar.diagnosticsFromLints(doc, window, [{
		span: { start: 6, end: 10 },
		message: 'Use world',
		suggestions: [{ kind: 'replace', replacement: 'world' }],
	}], 4);
	assert.equal(current.length, 1);
	assert.equal(grammar.acceptGrammarResult(current[0].version, 5), false);
});

test('ignore and dictionary actions identify a diagnostic without logging its text', async () => {
	const plugin = require(await bundleSource('src/harper/docxGrammarPlugin.ts', 'docx-grammar-plugin.cjs'));
	const doc = schema.node('doc', null, [paragraph('Hello wrld')]);
	const ignored = [];
	const dictionary = [];
	let pluginInstance;
	const state = EditorState.create({
		schema,
		doc,
		plugins: [
			pluginInstance = plugin.createDocxGrammarPlugin({
				getEnabled: () => true,
				requestLint: async () => [{
					span: { start: 6, end: 10 },
					message: 'Use world',
					suggestions: [{ kind: 'replace', replacement: 'world' }],
				}],
				actions: {
					ignore: (diagnostic) => ignored.push(diagnostic.id),
					addToDictionary: (word) => dictionary.push(word),
				},
				log: (data) => {
					assert.equal(JSON.stringify(data).includes('wrld'), false);
				},
			}),
		],
	});
	const dirty = state.tr.insertText('!', 7);
	let next = state.apply(dirty);
	const appended = pluginInstance.spec.appendTransaction([dirty], state, next);
	assert.ok(appended);
	next = next.apply(appended);
	const windows = plugin.docxGrammarPluginKey.getState(next).windows;
	assert.equal(windows.length, 1);
	assert.equal(windows[0].text.includes('Hello'), true);
	const diagnostics = grammar.diagnosticsFromLints(next.doc, windows[0], [{
		span: { start: windows[0].text.indexOf('wrld'), end: windows[0].text.indexOf('wrld') + 4 },
		message: 'Use world',
		suggestions: [{ kind: 'replace', replacement: 'world' }],
	}], plugin.docxGrammarPluginKey.getState(next).version);
	const withResults = next.apply(next.tr.setMeta(plugin.docxGrammarPluginKey, {
		type: 'results',
		version: diagnostics[0].version,
		diagnostics,
	}));
	const word = plugin.addDocxGrammarDictionaryWord(withResults, diagnostics[0].id, {
		addToDictionary: (value) => dictionary.push(value),
	});
	const ignoredTransaction = plugin.ignoreDocxGrammarDiagnostic(withResults, diagnostics[0].id, {
		ignore: (diagnostic) => ignored.push(diagnostic.id),
	});
	const cleared = withResults.apply(ignoredTransaction);
	assert.equal(word, 'wrld');
	assert.equal(dictionary[0], 'wrld');
	assert.equal(ignored[0], diagnostics[0].id);
	assert.equal(plugin.docxGrammarPluginKey.getState(cleared).diagnostics.length, 0);
	assert.equal(plugin.docxGrammarPluginKey.getState(cleared).decorations.find().length, 0);
});
