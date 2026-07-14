import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

let cachedErrorsModule;
let cachedOpRegistryModule;
let cachedAiCoreModule;

async function loadErrorsModule() {
	if (cachedErrorsModule) return cachedErrorsModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-test-'));
	const outfile = path.join(outputDirectory, 'ai-errors.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/errors.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cachedErrorsModule = require(outfile);
	return cachedErrorsModule;
}

async function loadOpRegistryModule() {
	if (cachedOpRegistryModule) return cachedOpRegistryModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-test-'));
	const outfile = path.join(outputDirectory, 'op-registry.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/opRegistry.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cachedOpRegistryModule = require(outfile);
	return cachedOpRegistryModule;
}

async function loadAiTestModule() {
	if (cachedAiCoreModule) return cachedAiCoreModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-test-'));
	const outfile = path.join(outputDirectory, 'ai-core.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/aiCore.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		external: ['./pptxDocumentService', './docxDocumentService', '../PresentationEngine', '../PowerPointPackage'],
	});
	cachedAiCoreModule = require(outfile);
	return cachedAiCoreModule;
}

test('AI op catalog validates known pptx.updateShapeText payload', async () => {
	const { validateDocumentOps } = await loadOpRegistryModule();
	const { AI_ERROR_CODES } = await loadErrorsModule();

	const result = validateDocumentOps([
		{
			op: 'pptx.updateShapeText',
			slideIndex: 2,
			shapeIndex: 3,
			text: 'Hello',
		},
	]);

	assert.equal(result.length, 0);

	const bad = validateDocumentOps([{ op: 'pptx.updateShapeText', slideIndex: 2 }]);
	assert.ok(bad.some((issue) => issue.code === AI_ERROR_CODES.SCHEMA_INVALID));
});

test('AI core gates apply when disabled and blocks without runtime when enabled', async () => {
	const { AiCore } = await loadAiTestModule();
	const { AI_ERROR_CODES } = await loadErrorsModule();

	let enabled = false;
	const core = new AiCore({
		getEnabled: () => enabled,
		pluginVersion: '1.0.33',
		runtime: null,
	});

	const disabled = await core.apply('deck.pptx', [
		{ op: 'pptx.updateShapeText', slideIndex: 0, shapeIndex: 0, text: 'x' },
	]);
	assert.equal(disabled.ok, false);
	assert.equal(disabled.errors[0]?.code, AI_ERROR_CODES.AI_DISABLED);

	enabled = true;
	const missingRuntime = await core.apply('deck.pptx', [
		{ op: 'pptx.updateShapeText', slideIndex: 0, shapeIndex: 0, text: 'x' },
	]);
	assert.equal(missingRuntime.ok, false);
	assert.equal(missingRuntime.errors[0]?.code, AI_ERROR_CODES.NOT_IMPLEMENTED);

	const docxBlocked = await core.apply('doc.docx', [
		{ op: 'docx.setRunText', blockId: 'body/p[0]', runId: 'body/p[0]/r[0]', text: 'x' },
	]);
	assert.equal(docxBlocked.ok, false);
	assert.equal(docxBlocked.errors[0]?.code, AI_ERROR_CODES.NOT_IMPLEMENTED);

	const manifest = core.buildManifest(true);
	assert.ok(manifest.operations.length >= 30);
	assert.equal(manifest.enabled, true);
	const pptxOps = manifest.operations.filter((op) => op.namespace === 'pptx');
	assert.ok(pptxOps.every((op) => op.status === 'implemented'));
	const docxOps = manifest.operations.filter((op) => op.namespace === 'docx');
	assert.ok(docxOps.every((op) => op.status === 'implemented'));
});
