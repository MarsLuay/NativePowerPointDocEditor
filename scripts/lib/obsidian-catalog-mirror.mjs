/**
 * Obsidian Community catalog ESLint scans the public GitHub repo for .ts/.tsx.
 * Shipping docx-editor package *source* (layout-painter, agents, vue, etc.) fails
 * catalog checks even when local eslint.config.mts ignores docx-editor/**.
 *
 * Vault / ObsidianNotes keeps full monorepo source for rebuilds.
 * The standalone NativePowerPointDocEditor mirror must be dist-only for packages
 * consumed at runtime: core, react, i18n.
 */
import { writeFileSync } from 'node:fs';

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
