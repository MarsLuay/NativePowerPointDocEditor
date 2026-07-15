/**
 * Obsidian Community catalog ESLint scans the public GitHub repo for .ts/.tsx.
 * Shipping docx-editor package *source* (layout-painter, agents, vue, etc.) fails
 * catalog checks even when local eslint.config.mts ignores docx-editor/**.
 *
 * Vault / ObsidianNotes keeps full monorepo source and package typings for rebuilds.
 * The standalone NativePowerPointDocEditor catalog mirror ships JS runtime only for
 * packages consumed at runtime: core, react, i18n (`.js` / `.mjs` / `.cjs` / `.css`).
 * No package `.d.ts` / `.d.mts` / `.d.cts` and no `types` / `typesVersions` /
 * `exports.*.types` pointers on the public surface — types live in the vault only.
 */
import { readdirSync, rmSync, writeFileSync } from 'node:fs';
import path from 'node:path';

export const CATALOG_DOCX_PACKAGES = ['core', 'react', 'i18n'];

/** Paths relative to plugin root (rsync --exclude-from style, no leading ./). */
export const CATALOG_MIRROR_RSYNC_EXCLUDES = [
	'.git/',
	'.DS_Store',
	'node_modules/',
	'**/node_modules/',
	'.code-analysis/',
	'*.map',
	'main.js',
	'ai/capabilities.json',
	'data.json',
	'results/',
	'tmp-meta-main.js',
	'test-ime-probe.docx',
	'ctxbar-review.mjs',
	'probe*.mjs',
	'tests/_probe-*.mjs',
	'scripts/visual-output/',
	'test-results/',
	'test_files/',
	'docs/AGENT-API.md',
	'docx-editor/node_modules/',
	'docx-editor/**/node_modules/',
	'docx-editor/**/.turbo/',
	'docx-editor/**/tsconfig.tsbuildinfo',
	'docx-editor/reference/',
	'docx-editor/examples/',
	'docx-editor/docs/',
	'docx-editor/.gstack/',
	'docx-editor/screenshots/',
	'docx-editor/.ralph/',
	'docx-editor/packages/agents/',
	'docx-editor/packages/vue/',
	'docx-editor/packages/nuxt/',
	'docx-editor/packages/*/src/',
	'docx-editor/packages/*/tests/',
	'docx-editor/packages/*/tsconfig.json',
	'docx-editor/packages/*/tsup.config.ts',
	'docx-editor/packages/*/vitest.config.ts',
	'docx-editor/packages/*/vitest.config.mts',
	'docx-editor/packages/*/tailwind*.{cjs,js,ts,mjs}',
	'docx-editor/packages/*/postcss.config.*',
	'docx-editor/packages/*/scripts/',
	// Best-effort: exclude package declaration emit (post-sync delete is the guarantee).
	'docx-editor/packages/*/**/*.d.ts',
	'docx-editor/packages/*/**/*.d.mts',
	'docx-editor/packages/*/**/*.d.cts',
	'docx-editor/packages/*/*.d.ts',
	'docx-editor/packages/*/*.d.mts',
	'docx-editor/packages/*/*.d.cts',
	'docx-editor/bun.lock',
	'docx-editor/bun.lockb',
	'docx-editor/turbo.json',
	'docx-editor/tsconfig*.json',
	'docx-editor/.eslintrc*',
	'docx-editor/eslint.config.*',
	'docx-editor/.prettierrc*',
	'docx-editor/.prettierignore',
	'docx-editor/.husky/',
	'docx-editor/.changeset/',
	'docx-editor/.claude/',
	'docx-editor/.github/',
	'docx-editor/openspec/',
];

export function writeCatalogMirrorExcludeFile(filePath) {
	writeFileSync(filePath, `${CATALOG_MIRROR_RSYNC_EXCLUDES.join('\n')}\n`, 'utf8');
}

/**
 * Strip typings surface from a package.json object (top-level + nested exports).
 * Mutates a deep clone; returns the clone.
 * @param {Record<string, unknown>} pkgJson
 * @returns {Record<string, unknown>}
 */
export function stripPackageTypingsFromPackageJson(pkgJson) {
	const out = structuredClone(pkgJson);
	delete out.types;
	delete out.typings;
	delete out.typesVersions;

	function stripNode(node) {
		if (!node || typeof node !== 'object') return;
		if (Array.isArray(node)) {
			for (const item of node) stripNode(item);
			return;
		}
		for (const key of Object.keys(node)) {
			if (key === 'types' || key === 'typings') {
				delete node[key];
				continue;
			}
			stripNode(node[key]);
		}
	}

	if (out.exports) stripNode(out.exports);
	return out;
}

/** Recursively delete `*.d.ts` / `*.d.mts` / `*.d.cts` under a package root. */
export function removePackageDeclarationFiles(pkgRoot) {
	const stack = [pkgRoot];
	while (stack.length) {
		const dir = stack.pop();
		for (const entry of readdirSync(dir, { withFileTypes: true })) {
			const full = path.join(dir, entry.name);
			if (entry.isDirectory()) {
				stack.push(full);
				continue;
			}
			if (/\.d\.[cm]?ts$/i.test(entry.name)) {
				rmSync(full, { force: true });
			}
		}
	}
}
