import {
	createDocxEditorAliases,
	resolveDocxEditorPackagesRoot,
} from '../../scripts/lib/docx-editor-aliases.mjs';

const aliasPromises = new Map();

/**
 * Runtime aliases used by test-local esbuild bundles. Tests compile source files
 * directly, so they do not inherit the aliases from esbuild.config.mjs.
 */
export function getDocxRuntimeAliases(projectRoot) {
	if (!aliasPromises.has(projectRoot)) {
		aliasPromises.set(
			projectRoot,
			createDocxEditorAliases(resolveDocxEditorPackagesRoot(projectRoot), projectRoot),
		);
	}
	return aliasPromises.get(projectRoot);
}
