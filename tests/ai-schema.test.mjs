import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

let cachedModules;

async function loadAiSchemaModules() {
	if (cachedModules) return cachedModules;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-schema-test-'));
	const registryOut = path.join(outputDirectory, 'op-registry.cjs');
	const examplesOut = path.join(outputDirectory, 'op-examples.cjs');
	const errorsOut = path.join(outputDirectory, 'errors.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/opRegistry.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile: registryOut,
		platform: 'node',
		target: 'node22',
	});
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/opExamples.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile: examplesOut,
		platform: 'node',
		target: 'node22',
	});
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/errors.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile: errorsOut,
		platform: 'node',
		target: 'node22',
	});
	cachedModules = {
		...require(registryOut),
		...require(examplesOut),
		...require(errorsOut),
	};
	return cachedModules;
}

test('every catalog operation has an object schema and example payload', async () => {
	const { listOpDefinitions, OP_EXAMPLES } = await loadAiSchemaModules();
	const operations = listOpDefinitions();

	assert.ok(operations.length > 0);
	assert.ok(operations.some((operation) => operation.id === 'pptx.setShapeFillColor'));
	for (const operation of operations) {
		assert.equal(operation.parameters.type, 'object', `${operation.id} parameters must be an object schema`);
		assert.ok(Array.isArray(operation.parameters.required), `${operation.id} must declare required fields`);
		assert.ok(OP_EXAMPLES[operation.id], `${operation.id} must have OP_EXAMPLES entry`);
		assert.equal(operation.status, 'implemented', `${operation.id} must be implemented`);
	}
});

test('OP_EXAMPLES payloads pass schema validation', async () => {
	const { validateDocumentOps, listOpDefinitions, OP_EXAMPLES } = await loadAiSchemaModules();

	for (const operation of listOpDefinitions()) {
		const example = OP_EXAMPLES[operation.id];
		const errors = validateDocumentOps([example]);
		assert.equal(errors.length, 0, `${operation.id} example failed validation: ${JSON.stringify(errors)}`);
	}
});

test('PPTX image replacement defaults to aspect-ratio-preserving cover', async () => {
	const { validateDocumentOps } = await loadAiSchemaModules();
	const errors = validateDocumentOps([{
		op: 'pptx.replaceImage',
		slideIndex: 0,
		shapeIndex: 0,
		vaultImagePath: 'assets/example.png',
	}]);
	assert.deepEqual(errors, []);
});

test('OP_EXAMPLES reject missing required fields', async () => {
	const { validateDocumentOps, OP_EXAMPLES, AI_ERROR_CODES } = await loadAiSchemaModules();

	for (const [opId, example] of Object.entries(OP_EXAMPLES)) {
		const broken = { ...example };
		const firstRequired = Object.keys(example).find((key) => key !== 'op');
		if (!firstRequired) continue;
		delete broken[firstRequired];
		const errors = validateDocumentOps([broken]);
		assert.ok(errors.length > 0, `${opId} should reject missing ${firstRequired}`);
		assert.equal(errors[0]?.code, AI_ERROR_CODES.SCHEMA_INVALID);
	}
});

test('generated capabilities.json includes per-op schemas and examples', async () => {
	const { listOpDefinitions } = await loadAiSchemaModules();
	const capabilities = JSON.parse(
		readFileSync(path.join(projectRoot, 'ai/capabilities.json'), 'utf8'),
	);
	assert.equal(capabilities.schemaVersion, 2);
	assert.ok(capabilities.limitations?.pptxFormats);
	assert.deepEqual(capabilities.limitations.pptxFormats.unsupported, ['ppt', 'pps', 'pot']);
	assert.ok(capabilities.limitations.pptxRuntime?.fallbackLimits?.length > 0);
	assert.deepEqual(
		capabilities.operations.map((operation) => operation.id),
		listOpDefinitions().map((operation) => operation.id),
	);
	for (const operation of capabilities.operations) {
		assert.equal(operation.parameters.type, 'object');
		assert.ok(operation.example, `${operation.id} missing example in capabilities.json`);
		assert.equal(operation.example.op, operation.id);
	}
	const replaceImage = capabilities.operations.find((operation) => operation.id === 'pptx.replaceImage');
	assert.ok(replaceImage, 'pptx.replaceImage must be in capabilities.json');
	assert.deepEqual(replaceImage.parameters.required, ['slideIndex', 'shapeIndex', 'vaultImagePath']);
});
