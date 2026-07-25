import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import {
	createDocxEditorAliases,
	docxEditorPackages,
	resolveDocxEditorPackagesRoot,
} from '../scripts/lib/docx-editor-aliases.mjs';

const root = new URL('../', import.meta.url);
const projectRoot = fileURLToPath(root);

async function readJson(relativePath) {
	return JSON.parse(await readFile(new URL(relativePath, root), 'utf8'));
}

test('DOCX editor packages resolve from the vendor runtime, not source or npm packages', async () => {
	const runtimeRoot = resolveDocxEditorPackagesRoot(projectRoot);
	const [manifest, lockfile, esbuildConfig, testBundler, tsconfig, aliases, provenance] = await Promise.all([
		readJson('package.json'),
		readJson('package-lock.json'),
		readFile(new URL('esbuild.config.mjs', root), 'utf8'),
		readFile(new URL('tests/helpers/load-plugin-modules.mjs', root), 'utf8'),
		readJson('tsconfig.json'),
		createDocxEditorAliases(runtimeRoot),
		readJson('vendor/docx-editor-runtime/provenance.json'),
	]);

	const declaredDependencies = {
		...manifest.dependencies,
		...manifest.devDependencies,
	};
	assert.deepEqual(
		Object.keys(declaredDependencies).filter((name) => name.startsWith('@npde/')),
		[],
	);

	const installedScopedDocxEditorPackages = Object.keys(lockfile.packages).filter((path) =>
		path.startsWith('node_modules/@npde/') || path.startsWith('node_modules/@eigenpal/'),
	);
	assert.deepEqual(installedScopedDocxEditorPackages, []);

	assert.match(esbuildConfig, /createDocxEditorAliases/);
	assert.match(esbuildConfig, /resolveDocxEditorPackagesRoot/);
	assert.match(testBundler, /createDocxEditorAliases/);
	assert.equal(docxEditorPackages['@npde/docx-editor-agents'], undefined);
	assert.equal(provenance.sourceBranch, 'docx-editor-source');
	assert.match(provenance.sourceCommit, /^[0-9a-f]{40}$/);
	await assert.rejects(
		() => readFile(new URL('src/vendor/eigenpal/README.md', root)),
		/ENOENT/,
	);

	for (const [packageName, dirName] of Object.entries(docxEditorPackages)) {
		const runtimeManifest = await readJson(`vendor/docx-editor-runtime/${dirName}/package.json`);
		assert.equal(runtimeManifest.version, '1.9.0');
		assert.equal(runtimeManifest.name, packageName);
		assert.equal(tsconfig.compilerOptions.paths[packageName], undefined);

		for (const [exportPath, target] of Object.entries(runtimeManifest.exports)) {
			const importTarget = typeof target === 'string'
				? target
				: target?.import ?? target?.require ?? target?.default;
			if (typeof importTarget !== 'string') continue;

			const exportTargets = typeof target === 'string'
				? [target]
				: Object.values(target).filter((value) => typeof value === 'string');
			for (const exportTarget of exportTargets) {
				const relativePath = `vendor/docx-editor-runtime/${dirName}/${exportTarget.replace(/^\.\//, '')}`;
				await readFile(new URL(relativePath, root));
			}

			const aliasKey = exportPath === '.'
				? packageName
				: `${packageName}/${exportPath.replace(/^\.\//, '')}`;
			assert.ok(aliases[aliasKey], `${aliasKey} needs a local runtime alias`);
		}
	}

	assert.equal(aliases['@npde/docx-editor-agents/react'], undefined);
	assert.ok(aliases.react?.includes(`${path.sep}node_modules${path.sep}react`), 'react alias must pin to plugin root');
	assert.ok(aliases['react-dom']?.includes(`${path.sep}node_modules${path.sep}react-dom`), 'react-dom alias must pin to plugin root');
});
