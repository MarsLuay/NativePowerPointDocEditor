import assert from 'node:assert/strict';
import { test } from 'node:test';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

let modulePromise;
async function loadDiagnosticsModule() {
	modulePromise ??= bundleSource('src/docxInputDiagnostics.ts', 'docx-input-diagnostics.cjs').then((outfile) => import(`file://${outfile}`));
	return modulePromise;
}

function target(id) {
	return {
		tagName: 'DIV',
		id,
		className: 'ProseMirror editor-surface',
		getAttribute(name) {
			return name === 'role' ? 'textbox' : null;
		},
	};
}

function keyboardEvent(overrides = {}) {
	return {
		key: 'Enter',
		code: 'Enter',
		repeat: false,
		timeStamp: 10,
		isTrusted: true,
		location: 0,
		eventPhase: 2,
		defaultPrevented: false,
		isComposing: false,
		target: target('docx-editor-a'),
		currentTarget: target('docx-editor-a'),
		...overrides,
	};
}

function relevantLogs(logs, message) {
	return logs.filter((entry) => entry.message === message).map((entry) => entry.data);
}

test('keyboard diagnostics correlate repeat, handler, beforeinput, keyup, and transaction routes', async () => {
	const { createDocxInputDiagnostics } = await loadDiagnosticsModule();
	const logs = [];
	const diagnostics = createDocxInputDiagnostics({
		viewId: 'editor-view-a',
		emit: (message, data) => logs.push({ message, data }),
	});
	const firstEnter = keyboardEvent();
	diagnostics.observeKeyDown(firstEnter);
	const firstHandler = diagnostics.beginHandler(firstEnter, 'test.keydown.handler');
	diagnostics.finishHandler(firstHandler, false, false);

	const repeatedEnter = keyboardEvent({ repeat: true, timeStamp: 20 });
	diagnostics.observeKeyDown(repeatedEnter);
	const repeatedHandler = diagnostics.beginHandler(repeatedEnter, 'test.keydown.handler');
	diagnostics.finishHandler(repeatedHandler, false, false);

	const beforeInput = keyboardEvent({
		inputType: 'insertParagraph',
		key: undefined,
		code: undefined,
		timeStamp: 21,
	});
	diagnostics.observeBeforeInput(beforeInput);
	diagnostics.recordTransaction({
		paragraphsBefore: 1,
		paragraphsAfter: 2,
		selectionBefore: { from: 5, to: 5, empty: true },
		selectionAfter: { from: 7, to: 7, empty: true },
		transactionCount: 1,
		docChangedCount: 1,
		steps: ['ReplaceStep/replace:5->5'],
		meta: ['uiEvent'],
	});
	diagnostics.observeKeyUp(keyboardEvent({ timeStamp: 22 }));

	const observed = relevantLogs(logs, 'DOCX input event observed');
	const keydowns = observed.filter((entry) => entry.eventType === 'keydown');
	assert.equal(keydowns.length, 2);
	assert.notEqual(keydowns[0].eventId, keydowns[1].eventId);
	assert.equal(keydowns[0].correlationId, keydowns[1].correlationId);
	assert.equal(keydowns[0].repeat, false);
	assert.equal(keydowns[1].repeat, true);
	assert.equal(keydowns[1].repeatIndex, 1);
	assert.equal(keydowns[0].target, 'div#docx-editor-a.ProseMirror.editor-surface[role=textbox]');
	assert.equal(keydowns[0].currentTarget, keydowns[0].target);
	assert.equal(keydowns[0].isTrusted, true);
	assert.equal(keydowns[0].eventPhase, 2);
	assert.equal(keydowns[0].isComposing, false);

	const before = observed.find((entry) => entry.eventType === 'beforeinput');
	assert.equal(before.inputType, 'insertParagraph');
	assert.equal(before.correlationId, keydowns[0].correlationId);
	assert.equal(before.keySequenceId, keydowns[0].keySequenceId);
	assert.equal(before.origin, 'dom.beforeinput');

	const keyup = observed.find((entry) => entry.eventType === 'keyup');
	assert.equal(keyup.correlationId, keydowns[0].correlationId);
	assert.equal(keyup.keySequenceId, keydowns[0].keySequenceId);

	const transaction = relevantLogs(logs, 'DOCX input transaction applied')[0];
	assert.equal(transaction.correlationId, before.correlationId);
	assert.deepEqual(transaction.selectionBefore, { from: 5, to: 5, empty: true });
	assert.deepEqual(transaction.selectionAfter, { from: 7, to: 7, empty: true });
	assert.equal(transaction.paragraphsBefore, 1);
	assert.equal(transaction.paragraphsAfter, 2);
	assert.equal(transaction.transactionCount, 1);
	assert.deepEqual(transaction.steps, ['ReplaceStep/replace:5->5']);
	assert.deepEqual(transaction.meta, ['uiEvent']);

	for (const entry of logs) {
		assert.equal('data' in entry.data, false, 'diagnostics must not capture document text');
		assert.equal(entry.data.viewId, 'editor-view-a');
		assert.match(entry.data.pluginId, /^docx-input-plugin-/);
	}
});

test('diagnostic IDs distinguish independent view and DOM input routes', async () => {
	const { createDocxInputDiagnostics } = await loadDiagnosticsModule();
	const firstLogs = [];
	const secondLogs = [];
	const first = createDocxInputDiagnostics({ viewId: 'editor-view-a', emit: (message, data) => firstLogs.push({ message, data }) });
	const second = createDocxInputDiagnostics({ viewId: 'editor-view-b', emit: (message, data) => secondLogs.push({ message, data }) });
	const firstBackspace = keyboardEvent({ key: 'Backspace', code: 'Backspace', target: target('editor-a') });
	const secondBackspace = keyboardEvent({ key: 'Backspace', code: 'Backspace', target: target('editor-b') });
	first.observeKeyDown(firstBackspace);
	second.observeKeyDown(secondBackspace);
	first.observeBeforeInput({ ...keyboardEvent(), inputType: 'deleteContentBackward', target: target('editor-a') });
	second.observeBeforeInput({ ...keyboardEvent(), inputType: 'deleteContentBackward', target: target('editor-b') });

	assert.notEqual(first.ids.pluginId, second.ids.pluginId);
	assert.notEqual(first.ids.handlerIds.keydown, second.ids.handlerIds.keydown);
	assert.equal(relevantLogs(firstLogs, 'DOCX input event observed').at(-1).origin, 'dom.beforeinput');
	assert.equal(relevantLogs(secondLogs, 'DOCX input event observed').at(-1).origin, 'dom.beforeinput');
	assert.equal(relevantLogs(firstLogs, 'DOCX input event observed').at(-1).target, 'div#editor-a.ProseMirror.editor-surface[role=textbox]');
	assert.equal(relevantLogs(secondLogs, 'DOCX input event observed').at(-1).target, 'div#editor-b.ProseMirror.editor-surface[role=textbox]');
});
