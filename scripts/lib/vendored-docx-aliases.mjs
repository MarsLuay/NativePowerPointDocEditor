import { readFile } from 'node:fs/promises';
import path from 'node:path';

export const vendoredDocxPackages = {
	'@eigenpal/docx-editor-agents': 'docx-editor-agents',
	'@eigenpal/docx-editor-core': 'docx-editor-core',
	'@eigenpal/docx-editor-i18n': 'docx-editor-i18n',
	'@eigenpal/docx-editor-react': 'docx-editor-react',
};

export async function createVendoredDocxAliases(vendorRoot) {
	const aliases = {};

	for (const [packageName, vendorDirName] of Object.entries(vendoredDocxPackages)) {
		const packageDir = path.resolve(vendorRoot, vendorDirName);
		const packageJson = JSON.parse(await readFile(path.join(packageDir, 'package.json'), 'utf8'));

		for (const [exportPath, target] of Object.entries(packageJson.exports ?? {})) {
			const importTarget = typeof target === 'string'
				? target
				: target?.import ?? target?.require ?? target?.default;

			if (typeof importTarget !== 'string') {
				continue;
			}

			const aliasKey = exportPath === '.'
				? packageName
				: `${packageName}/${exportPath.replace(/^\.\//, '')}`;
			aliases[aliasKey] = path.join(packageDir, importTarget);
		}
	}

	return aliases;
}
