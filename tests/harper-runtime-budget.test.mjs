import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { gzipSync } from 'node:zlib';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const {
	NPDE_MAIN_BUNDLE_LIMIT_BYTES,
	NPDE_RUNTIME_FILE_LIMIT_BYTES,
	assessHarperRuntimeContribution,
} = require(await bundleSource('src/harper/harperRuntimeBudget.ts', 'harper-runtime-budget.cjs'));

test('a budget-sized Harper artifact may be embedded and an oversized one may not', () => {
	const fits = assessHarperRuntimeContribution({
		rawBytes: 100_000,
		gzipBytes: 40_000,
		mainBytes: 1_000_000,
	});
	assert.equal(fits.withinFileLimit, true);
	assert.equal(fits.embedFitsMain, true);

	const oversized = assessHarperRuntimeContribution({
		rawBytes: NPDE_RUNTIME_FILE_LIMIT_BYTES + 1,
		gzipBytes: NPDE_MAIN_BUNDLE_LIMIT_BYTES,
		mainBytes: NPDE_MAIN_BUNDLE_LIMIT_BYTES,
	});
	assert.equal(oversized.withinFileLimit, false);
	assert.equal(oversized.embedFitsMain, false);
});

test('pinned Harper slim WASM stays out of the 5 MB runtime and main.js budgets', async () => {
	const packageJson = JSON.parse(await readFile(path.join(projectRoot, 'package.json'), 'utf8'));
	assert.equal(packageJson.dependencies['harper.js'], '2.10.0');

	const runtime = await readFile(path.join(projectRoot, 'src/harper/harperGrammarRuntime.mjs'), 'utf8');
	assert.match(runtime, /WorkerLinter/);
	assert.match(runtime, /slimBinary/);

	const wasmPath = path.join(projectRoot, 'node_modules/harper.js/dist/harper_wasm_slim_bg.wasm');
	const raw = await readFile(wasmPath);
	const gzipBytes = gzipSync(raw, { level: 9 }).length;
	const mainPath = path.join(projectRoot, 'main.js');
	let mainBytes = 4_989_999;
	try {
		mainBytes = (await readFile(mainPath)).length;
	} catch {
		mainBytes = NPDE_MAIN_BUNDLE_LIMIT_BYTES - 20_000;
	}
	const contribution = assessHarperRuntimeContribution({
		rawBytes: raw.length,
		gzipBytes,
		mainBytes,
	});
	assert.equal(contribution.withinFileLimit, false);
	assert.equal(contribution.embedFitsMain, false);
	assert.ok(raw.length > NPDE_RUNTIME_FILE_LIMIT_BYTES);
	assert.equal(createHash('sha256').update(raw).digest('hex').length, 64);
});
