import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { docxEditorAliases } from './helpers/docx-esbuild-aliases.mjs';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

let cachedCommandProtocolModule;

async function loadCommandProtocolModule() {
	if (cachedCommandProtocolModule) return cachedCommandProtocolModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-cmd-test-'));
	const outfile = path.join(outputDirectory, 'command-protocol.cjs');
	await build({
		alias: docxEditorAliases,
		entryPoints: [path.join(projectRoot, 'src/ai/commandProtocol.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cachedCommandProtocolModule = require(outfile);
	return cachedCommandProtocolModule;
}

test('command protocol resolves active path fallback', async () => {
	const {
		resolveDocumentPath,
		getActiveDocumentPath,
		parseApplyRequest,
		parseUndoRequest,
		parseRedoRequest,
		parseValidateRequest,
	} = await loadCommandProtocolModule();

	assert.equal(resolveDocumentPath('deck.pptx', 'active.docx'), 'deck.pptx');
	assert.equal(resolveDocumentPath(undefined, 'active.pptx'), 'active.pptx');
	assert.equal(resolveDocumentPath('', null), null);
	assert.equal(getActiveDocumentPath({ path: 'notes/a.pptx', extension: 'pptx' }), 'notes/a.pptx');
	assert.equal(getActiveDocumentPath({ path: 'notes/a.md', extension: 'md' }), null);

	const apply = parseApplyRequest({
		ops: [{ op: 'pptx.updateShapeText', slideIndex: 0, shapeIndex: 0, text: 'Hi' }],
		dryRun: true,
	});
	assert.equal(apply.path, undefined);
	assert.equal(apply.dryRun, true);
	assert.equal(apply.ops.length, 1);

	const validate = parseValidateRequest({ ops: [{ op: 'pptx.updateShapeText', slideIndex: 0, shapeIndex: 0, text: 'x' }] });
	assert.equal(validate.ops?.length, 1);

	const undo = parseUndoRequest({ path: 'notes/a.docx' });
	assert.equal(undo.path, 'notes/a.docx');
	const redo = parseRedoRequest({});
	assert.equal(redo.path, undefined);
});
