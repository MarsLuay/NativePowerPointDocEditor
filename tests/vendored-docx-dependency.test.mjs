import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import test from 'node:test';
import {
	createVendoredDocxAliases,
	vendoredDocxPackages,
} from '../scripts/lib/vendored-docx-aliases.mjs';

const root = new URL('../', import.meta.url);
const execFileAsync = promisify(execFile);

async function readJson(relativePath) {
	return JSON.parse(await readFile(new URL(relativePath, root), 'utf8'));
}

test('DOCX editor packages resolve from committed vendor artifacts, not npm dependencies', async () => {
	const [manifest, lockfile, esbuildConfig, testBundler, tsconfig, aliases, trackedVendorResult] = await Promise.all([
		readJson('package.json'),
		readJson('package-lock.json'),
		readFile(new URL('esbuild.config.mjs', root), 'utf8'),
		readFile(new URL('tests/helpers/load-plugin-modules.mjs', root), 'utf8'),
		readJson('tsconfig.json'),
		createVendoredDocxAliases(fileURLToPath(new URL('src/vendor/eigenpal/', root))),
		execFileAsync('git', ['ls-files', '-z', '--', 'src/vendor/eigenpal'], {
			cwd: fileURLToPath(root),
			encoding: 'utf8',
		}),
	]);
	const trackedVendorFiles = new Set(trackedVendorResult.stdout.split('\0').filter(Boolean));

	const declaredDependencies = {
		...manifest.dependencies,
		...manifest.devDependencies,
	};
	assert.deepEqual(
		Object.keys(declaredDependencies).filter((name) => name.startsWith('@eigenpal/')),
		[],
	);

	const installedEigenpalPackages = Object.keys(lockfile.packages).filter((path) =>
		path.startsWith('node_modules/@eigenpal/'),
	);
	assert.deepEqual(installedEigenpalPackages, []);

	assert.match(esbuildConfig, /createVendoredDocxAliases/);
	assert.match(testBundler, /createVendoredDocxAliases/);

	for (const [packageName, vendorDir] of Object.entries(vendoredDocxPackages)) {
		const vendoredManifest = await readJson(`src/vendor/eigenpal/${vendorDir}/package.json`);
		assert.equal(vendoredManifest.name, packageName);
		assert.ok(tsconfig.compilerOptions.paths[packageName], `${packageName} needs a local TypeScript path`);

		for (const [exportPath, target] of Object.entries(vendoredManifest.exports)) {
			const importTarget = typeof target === 'string'
				? target
				: target?.import ?? target?.require ?? target?.default;
			if (typeof importTarget !== 'string') continue;

			const exportTargets = typeof target === 'string'
				? [target]
				: Object.values(target).filter((value) => typeof value === 'string');
			for (const exportTarget of exportTargets) {
				const relativePath = `src/vendor/eigenpal/${vendorDir}/${exportTarget.replace(/^\.\//, '')}`;
				await readFile(new URL(relativePath, root));
				assert.ok(trackedVendorFiles.has(relativePath), `${relativePath} must be committed`);
			}

			const aliasKey = exportPath === '.'
				? packageName
				: `${packageName}/${exportPath.replace(/^\.\//, '')}`;
			assert.ok(aliases[aliasKey], `${aliasKey} needs a local runtime alias`);
		}
	}
});
