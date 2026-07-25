import { lstat, mkdir, rm, symlink } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const runtimeRoot = path.join(projectRoot, 'vendor', 'docx-editor-runtime');
const scopedPackagesRoot = path.join(projectRoot, 'node_modules', '@npde');
const packages = [
	['docx-editor-core', 'core'],
	['docx-editor-i18n', 'i18n'],
	['docx-editor-react', 'react'],
];

await mkdir(scopedPackagesRoot, { recursive: true });

for (const [packageName, directory] of packages) {
	const target = path.join(runtimeRoot, directory);
	const link = path.join(scopedPackagesRoot, packageName);
	try {
		await lstat(link);
		await rm(link, { recursive: true, force: true });
	} catch (error) {
		if (error && typeof error === 'object' && error.code !== 'ENOENT') throw error;
	}
	await symlink(target, link, 'dir');
}
