import { access, copyFile, cp, lstat, mkdir, mkdtemp, readFile, readdir, rename, rm } from 'node:fs/promises';
import path from 'node:path';

function isMissing(error) {
	return error && typeof error === 'object' && error.code === 'ENOENT';
}

async function pathExists(target) {
	try {
		await lstat(target);
		return true;
	} catch (error) {
		if (isMissing(error)) return false;
		throw error;
	}
}

async function assertMatchingFile(source, target) {
	const [sourceBytes, targetBytes] = await Promise.all([readFile(source), readFile(target)]);
	if (!sourceBytes.equals(targetBytes)) {
		throw new Error(`Deploy parity mismatch: ${source} != ${target}`);
	}
}

async function listFiles(root, relativePath = '') {
	const directory = path.join(root, relativePath);
	const entries = await readdir(directory, { withFileTypes: true });
	const files = [];
	for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
		const childPath = path.join(relativePath, entry.name);
		if (entry.isDirectory()) {
			files.push(...await listFiles(root, childPath));
		} else if (entry.isFile()) {
			files.push(childPath);
		} else {
			throw new Error(`Unsupported deploy artifact entry: ${path.join(root, childPath)}`);
		}
	}
	return files;
}

async function listFilesIfPresent(root) {
	return await pathExists(root) ? listFiles(root) : [];
}

async function assertMatchingTree(source, target) {
	const [sourceFiles, targetFiles] = await Promise.all([listFiles(source), listFiles(target)]);
	if (sourceFiles.length !== targetFiles.length || sourceFiles.some((file, index) => file !== targetFiles[index])) {
		throw new Error(`Deploy tree mismatch: ${source} != ${target}`);
	}
	await Promise.all(sourceFiles.map((file) => assertMatchingFile(path.join(source, file), path.join(target, file))));
}

async function replaceArtifact({ staged, target, backup }) {
	const hadTarget = await pathExists(target);
	if (hadTarget) {
		await mkdir(path.dirname(backup), { recursive: true });
		await rename(target, backup);
	}

	try {
		await rename(staged, target);
		return { target, backup, hadTarget };
	} catch (error) {
		if (hadTarget) {
			await rename(backup, target);
		}
		throw error;
	}
}

async function removeArtifact({ target, backup }) {
	if (!await pathExists(target)) {
		return null;
	}
	await mkdir(path.dirname(backup), { recursive: true });
	await rename(target, backup);
	return { target, backup, hadTarget: true };
}

async function restoreArtifacts(replacements) {
	for (const replacement of [...replacements].reverse()) {
		await rm(replacement.target, { force: true, recursive: true });
		if (replacement.hadTarget) {
			await rename(replacement.backup, replacement.target);
		}
	}
}

async function deployDirectoryContents({ source, staged, target, backupRoot, replacements }) {
	const [sourceFiles, targetFiles] = await Promise.all([listFiles(source), listFilesIfPresent(target)]);
	const sourceFileSet = new Set(sourceFiles);
	await mkdir(target, { recursive: true });

	for (const relativePath of sourceFiles) {
		const targetPath = path.join(target, relativePath);
		await mkdir(path.dirname(targetPath), { recursive: true });
		replacements.push(await replaceArtifact({
			staged: path.join(staged, relativePath),
			target: targetPath,
			backup: path.join(backupRoot, relativePath),
		}));
	}

	for (const relativePath of targetFiles) {
		if (sourceFileSet.has(relativePath)) continue;
		const targetPath = path.join(target, relativePath);
		const replacement = await removeArtifact({
			target: targetPath,
			backup: path.join(backupRoot, relativePath),
		});
		if (replacement) replacements.push(replacement);
	}
}

/**
 * Replace only build-owned plugin artifacts. File-level swaps keep deployment
 * transactional without renaming directories that Obsidian may hold open.
 * Settings, debug logs, and other runtime files stay untouched.
 */
export async function deployPluginArtifacts({ sourceDir, targetDir, files, directories }) {
	await access(targetDir);
	const stageRoot = await mkdtemp(path.join(path.dirname(targetDir), '.npde-deploy-'));
	const backupRoot = path.join(stageRoot, 'backups');
	const replacements = [];

	try {
		for (const file of files) {
			const source = path.join(sourceDir, file);
			const staged = path.join(stageRoot, file);
			await copyFile(source, staged);
			await assertMatchingFile(source, staged);
		}
		for (const directory of directories) {
			const source = path.join(sourceDir, directory);
			const staged = path.join(stageRoot, directory);
			await cp(source, staged, { recursive: true });
			await assertMatchingTree(source, staged);
		}

		for (const file of files) {
			replacements.push(await replaceArtifact({
				staged: path.join(stageRoot, file),
				target: path.join(targetDir, file),
				backup: path.join(backupRoot, file),
			}));
		}
		for (const directory of directories) {
			await deployDirectoryContents({
				source: path.join(sourceDir, directory),
				staged: path.join(stageRoot, directory),
				target: path.join(targetDir, directory),
				backupRoot: path.join(backupRoot, directory),
				replacements,
			});
		}

		for (const file of files) {
			await assertMatchingFile(path.join(sourceDir, file), path.join(targetDir, file));
		}
		for (const directory of directories) {
			await assertMatchingTree(path.join(sourceDir, directory), path.join(targetDir, directory));
		}
	} catch (error) {
		await restoreArtifacts(replacements);
		throw error;
	} finally {
		await Promise.all(replacements.map((replacement) => rm(replacement.backup, { force: true, recursive: true })));
		await rm(stageRoot, { force: true, recursive: true });
	}
}
