import { createHash } from 'node:crypto';
import { readdir, readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(scriptDirectory, '..');
const defaultRuntimeRoot = path.join(projectRoot, 'vendor', 'docx-editor-runtime');
const allowedRuntimeExtensions = new Set(['.js', '.mjs', '.cjs', '.css', '.json']);
const prohibitedManifestKeys = new Set([
	'types',
	'typings',
	'typesversions',
	'dependencies',
	'devdependencies',
	'peerdependencies',
	'optionaldependencies',
	'bundleddependencies',
	'bundledependencies',
	'scripts',
	'bin',
]);

const packages = [
	{ directory: 'core', name: '@npde/docx-editor-core' },
	{ directory: 'react', name: '@npde/docx-editor-react' },
	{ directory: 'i18n', name: '@npde/docx-editor-i18n' },
];

function fail(message) {
	throw new Error(message);
}

function toPosixPath(filePath) {
	return filePath.split(path.sep).join('/');
}

function isAllowedRuntimeFile(filePath) {
	return allowedRuntimeExtensions.has(path.extname(filePath).toLowerCase());
}

function isNoticeFile(name) {
	return /^(?:NOTICE|THIRD_PARTY_NOTICES)(?:[._-].+)?$/i.test(name);
}

async function listFiles(directory, prefix = '') {
	const files = [];
	for (const entry of (await readdir(directory, { withFileTypes: true }))
		.sort((left, right) => left.name.localeCompare(right.name))) {
		const relativePath = prefix ? path.join(prefix, entry.name) : entry.name;
		const absolutePath = path.join(directory, entry.name);
		if (entry.isDirectory()) {
			files.push(...await listFiles(absolutePath, relativePath));
			continue;
		}
		if (entry.isSymbolicLink()) {
			fail(`Runtime snapshot must not contain symbolic links: ${absolutePath}`);
		}
		if (!entry.isFile()) {
			fail(`Runtime snapshot contains an unsupported filesystem entry: ${absolutePath}`);
		}
		files.push(toPosixPath(relativePath));
	}
	return files;
}

async function readJson(filePath, label) {
	try {
		return JSON.parse(await readFile(filePath, 'utf8'));
	} catch (error) {
		fail(`Could not read ${label}: ${error.message}`);
	}
}

async function pathExists(filePath) {
	try {
		await stat(filePath);
		return true;
	} catch (error) {
		if (error && error.code === 'ENOENT') {
			return false;
		}
		throw error;
	}
}

function isSafeRuntimeTarget(target) {
	if (typeof target !== 'string' || !target.startsWith('./') || target.includes('\\')) {
		return false;
	}
	const normalized = path.posix.normalize(target);
	return !normalized.startsWith('../')
		&& normalized !== '..'
		&& (normalized === 'package.json'
			|| (normalized.startsWith('dist/') && isAllowedRuntimeFile(normalized)));
}

function selectAliasTarget(target) {
	if (typeof target === 'string') {
		return target;
	}
	if (!target || Array.isArray(target) || typeof target !== 'object') {
		return undefined;
	}
	return target.import ?? target.require ?? target.default;
}

async function validateExportTargets(packageRoot, target, label) {
	if (typeof target === 'string') {
		if (!isSafeRuntimeTarget(target)) {
			fail(`${label} points outside the runtime allowlist.`);
		}
		const resolvedTarget = path.resolve(packageRoot, target);
		if (!resolvedTarget.startsWith(`${packageRoot}${path.sep}`) || !(await pathExists(resolvedTarget))) {
			fail(`${label} resolves to a missing file.`);
		}
		return;
	}
	if (!target || Array.isArray(target) || typeof target !== 'object') {
		fail(`${label} must be a runtime export target or condition object.`);
	}
	for (const [condition, nestedTarget] of Object.entries(target)) {
		await validateExportTargets(packageRoot, nestedTarget, `${label}.${condition}`);
	}
}

function assertNoProhibitedManifestKeys(value, label) {
	if (!value || typeof value !== 'object') {
		return;
	}
	for (const [key, nestedValue] of Object.entries(value)) {
		if (prohibitedManifestKeys.has(key.toLowerCase())) {
			fail(`${label} contains prohibited manifest field: ${key}`);
		}
		assertNoProhibitedManifestKeys(nestedValue, `${label}.${key}`);
	}
}

async function validateManifest(runtimeRoot, packageSpec, provenance) {
	const packageRoot = path.join(runtimeRoot, packageSpec.directory);
	const manifest = await readJson(path.join(packageRoot, 'package.json'), `${packageSpec.directory}/package.json`);
	if (manifest.name !== packageSpec.name) {
		fail(`${packageSpec.directory}/package.json has unexpected name: ${String(manifest.name)}`);
	}
	if (typeof manifest.version !== 'string' || manifest.version.length === 0) {
		fail(`${packageSpec.directory}/package.json is missing a version.`);
	}
	if (provenance.packages?.[packageSpec.directory] !== manifest.version) {
		fail(`${packageSpec.directory}/package.json version does not match provenance.`);
	}
	assertNoProhibitedManifestKeys(manifest, `${packageSpec.directory}/package.json`);

	if (!manifest.exports || Array.isArray(manifest.exports) || typeof manifest.exports !== 'object') {
		fail(`${packageSpec.directory}/package.json must contain an object-valued exports map.`);
	}
	if (!manifest.exports['.']) {
		fail(`${packageSpec.directory}/package.json is missing its root runtime export.`);
	}

	for (const [exportPath, target] of Object.entries(manifest.exports)) {
		if (exportPath !== '.' && !exportPath.startsWith('./')) {
			fail(`${packageSpec.directory}/package.json has invalid export path: ${exportPath}`);
		}
		await validateExportTargets(packageRoot, target, `${packageSpec.directory}/package.json export ${exportPath}`);
		const aliasTarget = selectAliasTarget(target);
		if (!isSafeRuntimeTarget(aliasTarget)) {
			fail(`${packageSpec.directory}/package.json has an invalid alias target for ${exportPath}.`);
		}
		const resolvedTarget = path.resolve(packageRoot, aliasTarget);
		if (!resolvedTarget.startsWith(`${packageRoot}${path.sep}`) || !(await pathExists(resolvedTarget))) {
			fail(`${packageSpec.directory}/package.json alias ${exportPath} resolves to a missing file.`);
		}
	}

	for (const field of ['main', 'module']) {
		if (manifest[field] === undefined) {
			continue;
		}
		if (!isSafeRuntimeTarget(manifest[field])) {
			fail(`${packageSpec.directory}/package.json has invalid ${field}.`);
		}
		const resolvedTarget = path.resolve(packageRoot, manifest[field]);
		if (!(await pathExists(resolvedTarget))) {
			fail(`${packageSpec.directory}/package.json ${field} points to a missing file.`);
		}
	}
}

function isGeneratedPackageManifest(relativePath) {
	return packages.some((packageSpec) => relativePath === `${packageSpec.directory}/package.json`);
}

function validateFileTree(files) {
	for (const relativePath of files) {
		const segments = relativePath.split('/');
		if (segments.length === 1) {
			if (relativePath === 'LICENSE' || relativePath === 'provenance.json' || isNoticeFile(relativePath)) {
				continue;
			}
			fail(`Runtime snapshot has an unexpected root file: ${relativePath}`);
		}

		const packageSpec = packages.find((candidate) => candidate.directory === segments[0]);
		if (!packageSpec) {
			fail(`Runtime snapshot has an unexpected package path: ${relativePath}`);
		}
		if (segments.length === 2 && segments[1] === 'package.json') {
			continue;
		}
		if (segments[1] !== 'dist' || segments.length < 3) {
			fail(`Runtime snapshot has an unexpected package file: ${relativePath}`);
		}
		if (!isAllowedRuntimeFile(relativePath)) {
			fail(`Runtime snapshot has a non-runtime file: ${relativePath}`);
		}
		if (/\.(?:d\.)?(?:ts|tsx|mts|cts)$/i.test(relativePath)) {
			fail(`Runtime snapshot contains TypeScript or declaration output: ${relativePath}`);
		}
	}
}

async function sha256(filePath) {
	return createHash('sha256').update(await readFile(filePath)).digest('hex');
}

async function validateProvenance(runtimeRoot, files) {
	const provenance = await readJson(path.join(runtimeRoot, 'provenance.json'), 'provenance.json');
	if (provenance.schemaVersion !== 1) {
		fail('provenance.json has an unsupported schema version.');
	}
	if (typeof provenance.sourceRepository !== 'string' || provenance.sourceRepository.length === 0) {
		fail('provenance.json is missing sourceRepository.');
	}
	if (provenance.sourceBranch !== 'docx-editor-source') {
		fail('provenance.json must identify docx-editor-source as the source branch.');
	}
	if (typeof provenance.sourceCommit !== 'string' || !/^[0-9a-f]{40}$/i.test(provenance.sourceCommit)) {
		fail('provenance.json must include a full source commit SHA.');
	}
	if (provenance.license !== 'Apache-2.0') {
		fail('provenance.json must identify the Apache-2.0 license.');
	}
	if (!provenance.packages || typeof provenance.packages !== 'object') {
		fail('provenance.json is missing package versions.');
	}
	if (!provenance.fileHashes || Array.isArray(provenance.fileHashes) || typeof provenance.fileHashes !== 'object') {
		fail('provenance.json is missing file hashes.');
	}

	const hashedFiles = files.filter((relativePath) => !isGeneratedPackageManifest(relativePath) && relativePath !== 'provenance.json');
	if (provenance.fileCount !== hashedFiles.length) {
		fail('provenance.json fileCount does not match the copied runtime and license files.');
	}
	const provenancePaths = Object.keys(provenance.fileHashes).sort();
	if (JSON.stringify(provenancePaths) !== JSON.stringify([...hashedFiles].sort())) {
		fail('provenance.json file hashes must cover every copied runtime and license file exactly once.');
	}

	for (const relativePath of provenancePaths) {
		if (relativePath.includes('..') || path.isAbsolute(relativePath)) {
			fail(`provenance.json contains an unsafe hash path: ${relativePath}`);
		}
		const expectedHash = provenance.fileHashes[relativePath];
		if (typeof expectedHash !== 'string' || !/^[0-9a-f]{64}$/i.test(expectedHash)) {
			fail(`provenance.json has an invalid SHA-256 hash for ${relativePath}.`);
		}
		const actualHash = await sha256(path.join(runtimeRoot, relativePath));
		if (actualHash !== expectedHash) {
			fail(`Runtime file hash mismatch: ${relativePath}`);
		}
	}

	return provenance;
}

function resolveRuntimeRoot() {
	const args = process.argv.slice(2);
	if (args.length === 0) {
		return defaultRuntimeRoot;
	}
	if (args.length === 2 && args[0] === '--runtime-dir') {
		return path.resolve(process.cwd(), args[1]);
	}
	fail('Usage: node scripts/check-docx-runtime.mjs [--runtime-dir <directory>]');
}

async function main() {
	const runtimeRoot = resolveRuntimeRoot();
	if (!(await pathExists(runtimeRoot))) {
		fail(`Missing DOCX runtime snapshot: ${runtimeRoot}`);
	}
	const licenseText = await readFile(path.join(runtimeRoot, 'LICENSE'), 'utf8').catch(() => {
		fail('Runtime snapshot is missing LICENSE.');
	});
	if (!/Apache License[\s\S]*Version 2\.0/i.test(licenseText)) {
		fail('Runtime snapshot LICENSE is not Apache-2.0.');
	}

	const files = await listFiles(runtimeRoot);
	validateFileTree(files);
	if (files.includes('core/dist/mcp-cli.mjs')) {
		fail('Runtime snapshot must not include core/dist/mcp-cli.mjs.');
	}

	const provenance = await validateProvenance(runtimeRoot, files);
	if (typeof provenance.excludedFiles?.['core/dist/mcp-cli.mjs'] !== 'string') {
		fail('provenance.json must explain why core/dist/mcp-cli.mjs is excluded.');
	}
	for (const packageSpec of packages) {
		await validateManifest(runtimeRoot, packageSpec, provenance);
	}

	console.log(`DOCX runtime snapshot is valid (${files.length} files).`);
}

main().catch((error) => {
	console.error(`DOCX runtime validation failed: ${error.message}`);
	process.exitCode = 1;
});
