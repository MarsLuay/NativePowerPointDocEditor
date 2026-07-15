import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

/** Package folder names under docx-editor/packages/. */
export const docxEditorPackages = {
	'@npde/docx-editor-core': 'core',
	'@npde/docx-editor-i18n': 'i18n',
	'@npde/docx-editor-react': 'react',
};

const COMPAT_PREFIX = {
	'@npde/docx-editor-core': '@eigenpal/docx-editor-core',
	'@npde/docx-editor-i18n': '@eigenpal/docx-editor-i18n',
	'@npde/docx-editor-react': '@eigenpal/docx-editor-react',
};

export function resolvePluginProjectRoot(fromImportMetaUrl) {
	let dir = path.dirname(fileURLToPath(fromImportMetaUrl));
	while (dir !== path.dirname(dir)) {
		if (
			existsSync(path.join(dir, 'package.json'))
			&& existsSync(path.join(dir, 'manifest.json'))
			&& existsSync(path.join(dir, 'docx-editor'))
		) {
			return dir;
		}
		if (
			existsSync(path.join(dir, 'package.json'))
			&& existsSync(path.join(dir, 'manifest.json'))
		) {
			return dir;
		}
		dir = path.dirname(dir);
	}
	throw new Error('Could not find plugin project root.');
}

export function resolveDocxEditorPackagesRoot(projectRoot) {
	return path.join(projectRoot, 'docx-editor', 'packages');
}

/**
 * @param {string} packagesRoot absolute path to docx-editor/packages
 */
export async function createDocxEditorAliases(packagesRoot) {
	const aliases = {};

	for (const [packageName, dirName] of Object.entries(docxEditorPackages)) {
		const packageDir = path.resolve(packagesRoot, dirName);
		const packageJson = JSON.parse(await readFile(path.join(packageDir, 'package.json'), 'utf8'));

		for (const [exportPath, target] of Object.entries(packageJson.exports ?? {})) {
			const importTarget = typeof target === 'string'
				? target
				: target?.import ?? target?.require ?? target?.default;

			if (typeof importTarget !== 'string') {
				continue;
			}

			const resolved = path.join(packageDir, importTarget);
			const aliasKey = exportPath === '.'
				? packageName
				: `${packageName}/${exportPath.replace(/^\.\//, '')}`;
			aliases[aliasKey] = resolved;

			const compatBase = COMPAT_PREFIX[packageName];
			if (compatBase) {
				const compatKey = exportPath === '.'
					? compatBase
					: `${compatBase}/${exportPath.replace(/^\.\//, '')}`;
				aliases[compatKey] = resolved;
			}
		}
	}

	// Force a single React / ReactDOM copy. docx-editor's bun install can
	// nest a different React under node_modules/.bun; dual copies crash hooks
	// (`Cannot read properties of null (reading 'useState')`).
	const projectRoot = path.resolve(packagesRoot, '..', '..');
	const reactRoot = path.join(projectRoot, 'node_modules', 'react');
	const reactDomRoot = path.join(projectRoot, 'node_modules', 'react-dom');
	if (existsSync(reactRoot) && existsSync(reactDomRoot)) {
		aliases.react = reactRoot;
		aliases['react/jsx-runtime'] = path.join(reactRoot, 'jsx-runtime.js');
		aliases['react/jsx-dev-runtime'] = path.join(reactRoot, 'jsx-dev-runtime.js');
		aliases['react-dom'] = reactDomRoot;
		aliases['react-dom/client'] = path.join(reactDomRoot, 'client.js');
	}

	return aliases;
}
