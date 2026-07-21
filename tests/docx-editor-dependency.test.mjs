import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';
import { inspectReviewSurface } from '../scripts/check-review-surface.mjs';
import {
	createDocxEditorAliases,
	resolveDocxEditorPackagesRoot,
} from '../scripts/lib/docx-editor-aliases.mjs';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('DOCX runtime is vendored behind the local facade and has a reviewer-safe surface', async () => {
	const report = await inspectReviewSurface({ projectRoot });

	assert.deepEqual(report.violations, [], report.violations.join('\n'));
	assert.ok(report.aliasCount > report.reactAliasCount, 'runtime package export aliases should exist');
	assert.equal(report.reactAliasCount, 5, 'React and ReactDOM must stay single-copy aliases');
	assert.ok(report.provenance?.sourceCommit, 'runtime provenance must name the source commit');
	assert.equal(report.facade.bridgeExists, true, 'the runtime bridge must exist');
	assert.equal(report.facade.stylesExists, true, 'the runtime CSS boundary must exist');
});

test('DOCX runtime bridge bundles through esbuild with its required exports', async () => {
	const temporaryDirectory = await mkdtemp(path.join(tmpdir(), 'npde-docx-runtime-bridge-'));
	const outfile = path.join(temporaryDirectory, 'bridge.mjs');

	try {
		await build({
			alias: await createDocxEditorAliases(
				resolveDocxEditorPackagesRoot(projectRoot),
				projectRoot,
			),
			bundle: true,
			entryPoints: [path.join(projectRoot, 'src/docx/runtime/bridge.mjs')],
			format: 'esm',
			loader: { '.css': 'text' },
			logLevel: 'silent',
			outfile,
			platform: 'node',
			target: 'node22',
		});

		const bridge = await import(pathToFileURL(outfile).href);
		for (const [name, expectedType] of Object.entries({
			DocxEditor: 'object',
			clearParagraphMeasureCache: 'function',
			insertTable: 'function',
			setFontSize: 'function',
			setLineSpacing: 'function',
			loadFontFromBuffer: 'function',
			createT: 'function',
			deepMerge: 'function',
			en: 'object',
			loadDocxEditorLocale: 'function',
		})) {
			assert.equal(typeof bridge[name], expectedType, `bridge must export ${name}`);
		}
	} finally {
		await rm(temporaryDirectory, { force: true, recursive: true });
	}
});
