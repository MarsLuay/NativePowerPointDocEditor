import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

/** Runtime package folder names under vendor/docx-editor-runtime/. */
export const docxEditorPackages = {
	'@npde/docx-editor-core': 'core',
	'@npde/docx-editor-i18n': 'i18n',
	'@npde/docx-editor-react': 'react',
};

export function resolvePluginProjectRoot(fromImportMetaUrl) {
	let dir = path.dirname(fileURLToPath(fromImportMetaUrl));
	while (dir !== path.dirname(dir)) {
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
	return path.join(projectRoot, 'vendor', 'docx-editor-runtime');
}

/**
 * @param {string} runtimeRoot absolute path to vendor/docx-editor-runtime
 * @param {string} [projectRoot] plugin root; pass explicitly for robust React aliases
 */
export async function createDocxEditorAliases(
	runtimeRoot,
	projectRoot = path.resolve(runtimeRoot, '..', '..'),
) {
	const aliases = {};

	for (const [packageName, dirName] of Object.entries(docxEditorPackages)) {
		const packageDir = path.resolve(runtimeRoot, dirName);
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
		}
	}

	// Force a single React / ReactDOM copy. Vendored runtime modules must use
	// the plugin's React pair; dual copies crash hooks
	// (`Cannot read properties of null (reading 'useState')`).
	const resolvedProjectRoot = path.resolve(projectRoot);
	const reactRoot = path.join(resolvedProjectRoot, 'node_modules', 'react');
	const reactDomRoot = path.join(resolvedProjectRoot, 'node_modules', 'react-dom');
	aliases.react = reactRoot;
	aliases['react/jsx-runtime'] = path.join(reactRoot, 'jsx-runtime.js');
	aliases['react/jsx-dev-runtime'] = path.join(reactRoot, 'jsx-dev-runtime.js');
	aliases['react-dom'] = reactDomRoot;
	aliases['react-dom/client'] = path.join(reactDomRoot, 'client.js');

	return aliases;
}
