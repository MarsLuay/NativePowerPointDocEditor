import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import {
	CATALOG_MIRROR_RSYNC_EXCLUDES,
	assertCatalogMirrorSurface,
	pruneCatalogMirrorDocxTree,
} from '../scripts/lib/obsidian-catalog-mirror.mjs';

function makeCatalogFixture() {
	const root = mkdtempSync(path.join(tmpdir(), 'npde-catalog-surface-'));
	for (const pkg of ['core', 'react', 'i18n']) {
		const pkgRoot = path.join(root, 'docx-editor', 'packages', pkg);
		mkdirSync(path.join(pkgRoot, 'dist'), { recursive: true });
		writeFileSync(path.join(pkgRoot, 'dist', 'index.js'), 'export {};\n');
		writeFileSync(path.join(pkgRoot, 'package.json'), JSON.stringify({ main: './dist/index.js' }));
	}
	writeFileSync(path.join(root, 'docx-editor', 'package.json'), '{}');
	return root;
}

function writeFixture(filePath, contents) {
	mkdirSync(path.dirname(filePath), { recursive: true });
	writeFileSync(filePath, contents);
}

test('catalog pruning removes every non-runtime DOCX-editor path', (t) => {
	const root = makeCatalogFixture();
	t.after(() => rmSync(root, { recursive: true, force: true }));
	writeFixture(path.join(root, '.code-analysis', 'cache.json'), '{}');
	writeFixture(path.join(root, 'docx-editor', 'scripts', 'generate.ts'), 'export {};');
	writeFixture(path.join(root, 'docx-editor', 'packages', 'core', 'testdata', 'sample.docx'), 'fixture');
	writeFixture(path.join(root, 'docx-editor', 'packages', 'react', 'src', 'index.ts'), 'export {};');

	rmSync(path.join(root, '.code-analysis'), { recursive: true, force: true });
	pruneCatalogMirrorDocxTree(root);

	assert.doesNotThrow(() => assertCatalogMirrorSurface(root));
	assert.equal(existsSync(path.join(root, 'docx-editor', 'scripts')), false);
	assert.equal(existsSync(path.join(root, 'docx-editor', 'packages', 'core', 'testdata')), false);
	assert.equal(existsSync(path.join(root, 'docx-editor', 'packages', 'react', 'src')), false);
});

test('catalog surface guard rejects shipped TypeScript, type metadata, and analyzer caches', (t) => {
	const root = makeCatalogFixture();
	t.after(() => rmSync(root, { recursive: true, force: true }));
	writeFixture(path.join(root, '.code-analysis', 'cache.json'), '{}');
	writeFileSync(path.join(root, 'docx-editor', 'packages', 'core', 'dist', 'leak.ts'), 'export {};');
	writeFileSync(
		path.join(root, 'docx-editor', 'packages', 'react', 'package.json'),
		JSON.stringify({ types: './dist/index.d.ts' }),
	);

	assert.throws(
		() => assertCatalogMirrorSurface(root),
		(error) => {
			assert.match(error.message, /TypeScript must not ship/);
			assert.match(error.message, /type metadata must not ship/);
			assert.match(error.message, /analyzer cache must not ship/);
			return true;
		},
	);
});

test('catalog synchronization excludes every known non-runtime DOCX surface', () => {
	assert.ok(CATALOG_MIRROR_RSYNC_EXCLUDES.includes('docx-editor/scripts/'));
	assert.ok(CATALOG_MIRROR_RSYNC_EXCLUDES.includes('docx-editor/packages/*/testdata/'));
	assert.ok(CATALOG_MIRROR_RSYNC_EXCLUDES.includes('.code-analysis/'));
});
