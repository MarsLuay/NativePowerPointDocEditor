import { createHash } from 'node:crypto';
import {
	copyFile,
	mkdir,
	readdir,
	readFile,
	rename,
	rm,
	stat,
	writeFile,
} from 'node:fs/promises';
import { execFile } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(scriptDirectory, '..');
const runtimeRoot = path.join(projectRoot, 'vendor', 'docx-editor-runtime');
const sourceBranch = 'docx-editor-source';
const allowedRuntimeExtensions = new Set(['.js', '.mjs', '.cjs', '.css', '.json']);

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

function isSafeRuntimeTarget(target) {
	if (typeof target !== 'string' || !target.startsWith('./') || target.includes('\\')) {
		return false;
	}

	const normalized = path.posix.normalize(target);
	if (normalized.startsWith('../') || normalized === '..') {
		return false;
	}

	return normalized === 'package.json'
		|| (normalized.startsWith('dist/') && isAllowedRuntimeFile(normalized));
}

function assertRuntimeTarget(target, label) {
	if (!isSafeRuntimeTarget(target)) {
		fail(`${label} points outside the runtime allowlist: ${String(target)}`);
	}
	return target;
}

function sanitizeExportConditions(value, label) {
	if (typeof value === 'string') {
		return assertRuntimeTarget(value, label);
	}

	if (!value || Array.isArray(value) || typeof value !== 'object') {
		fail(`${label} must be a runtime export target or condition object.`);
	}

	const sanitized = {};
	for (const [condition, target] of Object.entries(value)) {
		if (condition === 'types' || condition === 'typings' || condition === 'typesVersions') {
			continue;
		}

		const runtimeTarget = sanitizeExportConditions(target, `${label}.${condition}`);
		if (runtimeTarget !== undefined) {
			sanitized[condition] = runtimeTarget;
		}
	}

	if (Object.keys(sanitized).length === 0) {
		return undefined;
	}

	return sanitized;
}

function sanitizeExports(exportsField, packageName) {
	if (!exportsField || Array.isArray(exportsField) || typeof exportsField !== 'object') {
		fail(`${packageName} must provide an object-valued exports map.`);
	}

	const sanitized = {};
	for (const [exportPath, target] of Object.entries(exportsField)) {
		if (exportPath === 'types' || exportPath === 'typings' || exportPath === 'typesVersions') {
			continue;
		}
		if (exportPath !== '.' && !exportPath.startsWith('./')) {
			fail(`${packageName} has an invalid export path: ${exportPath}`);
		}

		const runtimeTarget = sanitizeExportConditions(target, `${packageName}.exports[${exportPath}]`);
		if (runtimeTarget !== undefined) {
			sanitized[exportPath] = runtimeTarget;
		}
	}

	if (!sanitized['.']) {
		fail(`${packageName} has no runtime root export after type export removal.`);
	}

	return sanitized;
}

function createRuntimeManifest(sourceManifest, expectedName) {
	if (sourceManifest.name !== expectedName) {
		fail(`Expected ${expectedName}, found ${String(sourceManifest.name)}.`);
	}
	if (typeof sourceManifest.version !== 'string' || sourceManifest.version.length === 0) {
		fail(`${expectedName} is missing a version.`);
	}

	const manifest = {
		name: sourceManifest.name,
		version: sourceManifest.version,
		license: sourceManifest.license ?? 'Apache-2.0',
		sideEffects: sourceManifest.sideEffects === true,
	};

	if (typeof sourceManifest.description === 'string') {
		manifest.description = sourceManifest.description;
	}
	for (const field of ['main', 'module']) {
		if (typeof sourceManifest[field] === 'string') {
			manifest[field] = assertRuntimeTarget(sourceManifest[field], `${expectedName}.${field}`);
		}
	}

	manifest.exports = sanitizeExports(sourceManifest.exports, expectedName);
	return manifest;
}

async function listFiles(directory, prefix = '') {
	const entries = await readdir(directory, { withFileTypes: true });
	const files = [];

	for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
		const relativePath = prefix ? path.join(prefix, entry.name) : entry.name;
		const absolutePath = path.join(directory, entry.name);
		if (entry.isDirectory()) {
			files.push(...await listFiles(absolutePath, relativePath));
			continue;
		}
		if (entry.isSymbolicLink()) {
			fail(`Refusing to vendor symbolic link: ${absolutePath}`);
		}
		if (!entry.isFile()) {
			fail(`Refusing to vendor unsupported filesystem entry: ${absolutePath}`);
		}
		files.push(relativePath);
	}

	return files;
}

async function sha256(filePath) {
	return createHash('sha256').update(await readFile(filePath)).digest('hex');
}

async function copyRuntimeDist(sourceDist, destinationDist, packageDirectory, hashes) {
	const copied = [];
	const omitted = [];

	for (const relativePath of await listFiles(sourceDist)) {
		const normalizedPath = toPosixPath(relativePath);
		if (packageDirectory === 'core' && normalizedPath === 'mcp-cli.mjs') {
			omitted.push(`core/dist/${normalizedPath}`);
			continue;
		}
		if (!isAllowedRuntimeFile(normalizedPath)) {
			omitted.push(`${packageDirectory}/dist/${normalizedPath}`);
			continue;
		}

		const sourceFile = path.join(sourceDist, relativePath);
		const destinationFile = path.join(destinationDist, relativePath);
		await mkdir(path.dirname(destinationFile), { recursive: true });
		await copyFile(sourceFile, destinationFile);

		const hashPath = `${packageDirectory}/dist/${normalizedPath}`;
		hashes[hashPath] = await sha256(destinationFile);
		copied.push(hashPath);
	}

	if (copied.length === 0) {
		fail(`${sourceDist} did not contain any allowlisted runtime files.`);
	}

	return { copied, omitted };
}

function normalizeRepository(remoteUrl) {
	const githubMatch = remoteUrl.match(/github\.com[/:]([^/]+\/[^/#]+?)(?:\.git)?$/i);
	return githubMatch ? githubMatch[1] : remoteUrl;
}

async function runGit(sourceDirectory, args) {
	const { stdout } = await execFileAsync('git', ['-C', sourceDirectory, ...args], {
		encoding: 'utf8',
	});
	return stdout.trim();
}

async function verifySourceWorktree(sourceDirectory) {
	let isWorktree;
	try {
		isWorktree = await runGit(sourceDirectory, ['rev-parse', '--is-inside-work-tree']);
	} catch {
		fail(`DOCX source is not a Git worktree: ${sourceDirectory}`);
	}
	if (isWorktree !== 'true') {
		fail(`DOCX source is not a Git worktree: ${sourceDirectory}`);
	}

	const branch = await runGit(sourceDirectory, ['branch', '--show-current']);
	if (branch !== sourceBranch) {
		fail(`DOCX source must be on ${sourceBranch}; found ${branch || 'a detached HEAD'}.`);
	}

	const commit = await runGit(sourceDirectory, ['rev-parse', 'HEAD']);
	const remoteUrl = await runGit(sourceDirectory, ['remote', 'get-url', 'origin']).catch(() => '');
	if (!remoteUrl) {
		fail(`DOCX source worktree has no origin remote: ${sourceDirectory}`);
	}

	return {
		commit,
		repository: normalizeRepository(remoteUrl),
	};
}

async function readJson(filePath, label) {
	try {
		return JSON.parse(await readFile(filePath, 'utf8'));
	} catch (error) {
		fail(`Could not read ${label} at ${filePath}: ${error.message}`);
	}
}

function sortedObject(object) {
	return Object.fromEntries(Object.entries(object).sort(([left], [right]) => left.localeCompare(right)));
}

function isNoticeFile(name) {
	return /^(?:NOTICE|THIRD_PARTY_NOTICES)(?:[._-].+)?$/i.test(name);
}

async function copyLicenseAndNotices(sourceDocxRoot, stagingRoot, hashes) {
	const sourceLicense = path.join(sourceDocxRoot, 'LICENSE');
	const licenseText = await readFile(sourceLicense, 'utf8').catch(() => {
		fail(`Missing Apache-2.0 LICENSE at ${sourceLicense}.`);
	});
	if (!/Apache License[\s\S]*Version 2\.0/i.test(licenseText)) {
		fail(`Expected Apache-2.0 text in ${sourceLicense}.`);
	}

	const destinationLicense = path.join(stagingRoot, 'LICENSE');
	await copyFile(sourceLicense, destinationLicense);
	hashes.LICENSE = await sha256(destinationLicense);

	const notices = [];
	for (const entry of await readdir(sourceDocxRoot, { withFileTypes: true })) {
		if (!entry.isFile() || !isNoticeFile(entry.name)) {
			continue;
		}
		const sourceNotice = path.join(sourceDocxRoot, entry.name);
		const destinationNotice = path.join(stagingRoot, entry.name);
		await copyFile(sourceNotice, destinationNotice);
		hashes[entry.name] = await sha256(destinationNotice);
		notices.push(entry.name);
	}

	return notices.sort((left, right) => left.localeCompare(right));
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

async function replaceDirectory(stagingRoot, destinationRoot) {
	const parentDirectory = path.dirname(destinationRoot);
	const backupRoot = path.join(
		parentDirectory,
		`.docx-editor-runtime.backup-${process.pid}-${Date.now()}`,
	);
	let movedExistingDirectory = false;
	let installed = false;

	try {
		if (await pathExists(destinationRoot)) {
			await rename(destinationRoot, backupRoot);
			movedExistingDirectory = true;
		}
		await rename(stagingRoot, destinationRoot);
		installed = true;
	} catch (error) {
		if (
			movedExistingDirectory
			&& !(await pathExists(destinationRoot))
			&& await pathExists(backupRoot)
		) {
			await rename(backupRoot, destinationRoot);
		}
		throw error;
	} finally {
		if (installed && movedExistingDirectory) {
			await rm(backupRoot, { recursive: true, force: true });
		}
	}
}

async function validateStagingDirectory(stagingRoot) {
	const checker = path.join(scriptDirectory, 'check-docx-runtime.mjs');
	try {
		await execFileAsync(process.execPath, [checker, '--runtime-dir', stagingRoot], {
			cwd: projectRoot,
			encoding: 'utf8',
		});
	} catch (error) {
		const output = [error.stdout, error.stderr].filter(Boolean).join('\n').trim();
		fail(`Generated runtime snapshot failed validation.${output ? `\n${output}` : ''}`);
	}
}

function resolveSourceDirectory() {
	const args = process.argv.slice(2);
	if (args.length > 1) {
		fail('Usage: node scripts/vendor-docx-editor-runtime.mjs [source-worktree]');
	}
	const sourceArgument = args[0]
		?? process.env.DOCX_EDITOR_SOURCE_DIR
		?? path.join(projectRoot, '..', 'NPDE-docx-editor-source');
	return path.resolve(process.cwd(), sourceArgument);
}

async function main() {
	const sourceDirectory = resolveSourceDirectory();
	const source = await verifySourceWorktree(sourceDirectory);
	const sourceDocxRoot = path.join(sourceDirectory, 'docx-editor');
	const stagingRoot = path.join(
		path.dirname(runtimeRoot),
		`.docx-editor-runtime.staging-${process.pid}-${Date.now()}`,
	);
	const hashes = {};
	const packageVersions = {};
	const copiedFiles = [];
	const omittedFiles = [];
	let noticeFiles = [];

	await mkdir(path.dirname(runtimeRoot), { recursive: true });
	await rm(stagingRoot, { recursive: true, force: true });

	try {
		await mkdir(stagingRoot, { recursive: true });

		for (const packageSpec of packages) {
			const sourcePackageRoot = path.join(sourceDocxRoot, 'packages', packageSpec.directory);
			const sourceDist = path.join(sourcePackageRoot, 'dist');
			if (!(await pathExists(sourceDist)) || !(await stat(sourceDist)).isDirectory()) {
				fail(`Missing built dist directory: ${sourceDist}`);
			}

			const sourceManifest = await readJson(
				path.join(sourcePackageRoot, 'package.json'),
				`${packageSpec.name} package.json`,
			);
			const destinationPackageRoot = path.join(stagingRoot, packageSpec.directory);
			await mkdir(destinationPackageRoot, { recursive: true });
			const runtimeManifest = createRuntimeManifest(sourceManifest, packageSpec.name);
			await writeFile(
				path.join(destinationPackageRoot, 'package.json'),
				`${JSON.stringify(runtimeManifest, null, '\t')}\n`,
			);

			const copied = await copyRuntimeDist(
				sourceDist,
				path.join(destinationPackageRoot, 'dist'),
				packageSpec.directory,
				hashes,
			);
			copiedFiles.push(...copied.copied);
			omittedFiles.push(...copied.omitted);
			packageVersions[packageSpec.directory] = runtimeManifest.version;
		}

		noticeFiles = await copyLicenseAndNotices(sourceDocxRoot, stagingRoot, hashes);
		const provenance = {
			schemaVersion: 1,
			sourceRepository: source.repository,
			sourceBranch,
			sourceCommit: source.commit,
			license: 'Apache-2.0',
			packages: packageVersions,
			fileCount: copiedFiles.length + 1 + noticeFiles.length,
			fileHashes: sortedObject(hashes),
			excludedFiles: {
				'core/dist/mcp-cli.mjs': 'CLI entry point excluded because package bin metadata is removed and the Obsidian runtime does not invoke it.',
			},
		};
		await writeFile(
			path.join(stagingRoot, 'provenance.json'),
			`${JSON.stringify(provenance, null, '\t')}\n`,
		);

		await validateStagingDirectory(stagingRoot);
		await replaceDirectory(stagingRoot, runtimeRoot);
	} catch (error) {
		await rm(stagingRoot, { recursive: true, force: true });
		throw error;
	}

	const skippedMcpCli = omittedFiles.includes('core/dist/mcp-cli.mjs');
	console.log(`Vendored ${copiedFiles.length + 1 + noticeFiles.length} runtime files from ${source.commit} into vendor/docx-editor-runtime.`);
	if (skippedMcpCli) {
		console.log('Excluded core/dist/mcp-cli.mjs: it is a package CLI bin entry, not an Obsidian runtime dependency.');
	}
}

const invokedPath = process.argv[1] ? path.resolve(process.argv[1]) : null;
if (invokedPath && pathToFileURL(invokedPath).href === import.meta.url) {
	main().catch((error) => {
		console.error(`DOCX runtime vendoring failed: ${error.message}`);
		process.exitCode = 1;
	});
}
