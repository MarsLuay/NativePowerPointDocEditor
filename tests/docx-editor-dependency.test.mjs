import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import {
	createDocxEditorAliases,
	docxEditorPackages,
	resolveDocxEditorAgentsStub,
	resolveDocxEditorPackagesRoot,
} from '../scripts/lib/docx-editor-aliases.mjs';

const root = new URL('../', import.meta.url);
const projectRoot = fileURLToPath(root);

async function readJson(relativePath) {
	return JSON.parse(await readFile(new URL(relativePath, root), 'utf8'));
}

test('DOCX editor packages resolve from in-repo docx-editor monorepo, not npm @eigenpal dependencies', async () => {
	const packagesRoot = resolveDocxEditorPackagesRoot(projectRoot);
	const [manifest, lockfile, esbuildConfig, testBundler, tsconfig, aliases] = await Promise.all([
		readJson('package.json'),
		readJson('package-lock.json'),
		readFile(new URL('esbuild.config.mjs', root), 'utf8'),
		readFile(new URL('tests/helpers/load-plugin-modules.mjs', root), 'utf8'),
		readJson('tsconfig.json'),
		createDocxEditorAliases(packagesRoot, {
			agentsStubPath: resolveDocxEditorAgentsStub(projectRoot),
		}),
	]);

	const declaredDependencies = {
		...manifest.dependencies,
		...manifest.devDependencies,
	};
	assert.deepEqual(
		Object.keys(declaredDependencies).filter((name) => name.startsWith('@eigenpal/') || name.startsWith('@npde/')),
		[],
	);

	const installedEigenpalPackages = Object.keys(lockfile.packages).filter((path) =>
		path.startsWith('node_modules/@eigenpal/'),
	);
	assert.deepEqual(installedEigenpalPackages, []);

	assert.match(esbuildConfig, /createDocxEditorAliases/);
	assert.match(esbuildConfig, /resolveDocxEditorPackagesRoot/);
	assert.match(testBundler, /createDocxEditorAliases/);
	assert.equal(docxEditorPackages['@npde/docx-editor-agents'], undefined);

	await assert.rejects(
		() => readFile(new URL('docx-editor/packages/agents-not-a-package/package.json', root)),
		/ENOENT/,
	);
	await assert.rejects(
		() => readFile(new URL('src/vendor/eigenpal/README.md', root)),
		/ENOENT/,
	);

	for (const [packageName, dirName] of Object.entries(docxEditorPackages)) {
		const localManifest = await readJson(`docx-editor/packages/${dirName}/package.json`);
		assert.equal(localManifest.version, '1.9.0');
		assert.ok(tsconfig.compilerOptions.paths[packageName], `${packageName} needs a local TypeScript path`);

		for (const [exportPath, target] of Object.entries(localManifest.exports)) {
			const importTarget = typeof target === 'string'
				? target
				: target?.import ?? target?.require ?? target?.default;
			if (typeof importTarget !== 'string') continue;

			const exportTargets = typeof target === 'string'
				? [target]
				: Object.values(target).filter((value) => typeof value === 'string');
			for (const exportTarget of exportTargets) {
				const relativePath = `docx-editor/packages/${dirName}/${exportTarget.replace(/^\.\//, '')}`;
				await readFile(new URL(relativePath, root));
			}

			const aliasKey = exportPath === '.'
				? packageName
				: `${packageName}/${exportPath.replace(/^\.\//, '')}`;
			assert.ok(aliases[aliasKey], `${aliasKey} needs a local runtime alias`);

			const compatKey = aliasKey.replace('@npde/', '@eigenpal/');
			assert.ok(aliases[compatKey], `${compatKey} needs a compatibility alias for in-dist imports`);
		}
	}

	await readFile(new URL('src/docx/editor/agentsStub/react.mjs', root));
	assert.ok(aliases['@eigenpal/docx-editor-agents/react']);
});
