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
import { existsSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import path from 'node:path';

export const CATALOG_DOCX_PACKAGES = ['core', 'react', 'i18n'];
const CATALOG_DOCX_ROOT_FILES = new Set(['LICENSE', 'README.md', 'SOURCE_MIRROR.md', 'package.json']);
const CATALOG_DOCX_PACKAGE_FILES = new Set(['LICENSE', 'README.md', 'package.json']);

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
	'docx-editor/packages/*/testdata/',
	'docx-editor/packages/*/tsconfig.json',
	'docx-editor/packages/*/tsup.config.ts',
	'docx-editor/packages/*/vitest.config.ts',
	'docx-editor/packages/*/vitest.config.mts',
	'docx-editor/packages/*/tailwind*.{cjs,js,ts,mjs}',
	'docx-editor/packages/*/postcss.config.*',
	'docx-editor/packages/*/scripts/',
	'docx-editor/scripts/',
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
	// Vault-local leftover / release-branch vendor snapshot. Catalog mirror uses
	// `docx-editor/packages/{core,react,i18n}/dist` via docx-editor-aliases.
	'vendor/docx-editor-runtime/',
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

/**
 * Reduce the vendored editor to the smallest catalog-safe runtime contract.
 * This is deliberately an allowlist: new upstream source/test/tool directories
 * cannot silently become public review input when the mirror is synchronized.
 */
export function pruneCatalogMirrorDocxTree(mirrorRoot) {
	const docxRoot = path.join(mirrorRoot, 'docx-editor');
	if (!existsSync(docxRoot)) return;

	for (const entry of readdirSync(docxRoot, { withFileTypes: true })) {
		if (entry.name === 'packages' || CATALOG_DOCX_ROOT_FILES.has(entry.name)) continue;
		rmSync(path.join(docxRoot, entry.name), { recursive: true, force: true });
	}

	const packagesRoot = path.join(docxRoot, 'packages');
	if (!existsSync(packagesRoot)) return;
	for (const entry of readdirSync(packagesRoot, { withFileTypes: true })) {
		const pkgRoot = path.join(packagesRoot, entry.name);
		if (!entry.isDirectory() || !CATALOG_DOCX_PACKAGES.includes(entry.name)) {
			rmSync(pkgRoot, { recursive: true, force: true });
			continue;
		}
		for (const child of readdirSync(pkgRoot, { withFileTypes: true })) {
			if (child.name === 'dist' || CATALOG_DOCX_PACKAGE_FILES.has(child.name)) continue;
			rmSync(path.join(pkgRoot, child.name), { recursive: true, force: true });
		}
	}
}

function collectCatalogSurfaceViolations(root, relative = '') {
	const directory = path.join(root, relative);
	const violations = [];
	for (const entry of readdirSync(directory, { withFileTypes: true })) {
		const childRelative = path.posix.join(relative, entry.name);
		if (entry.isDirectory()) {
			violations.push(...collectCatalogSurfaceViolations(root, childRelative));
			continue;
		}
		if (/\.(?:ts|tsx|mts|cts)$/i.test(entry.name)) {
			violations.push(`${childRelative}: TypeScript must not ship in the catalog DOCX runtime`);
		}
	}
	return violations;
}

function containsTypeMetadata(value) {
	if (!value || typeof value !== 'object') return false;
	if (Array.isArray(value)) return value.some(containsTypeMetadata);
	return Object.entries(value).some(([key, child]) => (
		key === 'types' || key === 'typings' || key === 'typesVersions' || containsTypeMetadata(child)
	));
}

/** Throw when a public mirror contains non-runtime DOCX-editor material. */
export function assertCatalogMirrorSurface(mirrorRoot) {
	const violations = [];
	const analyzerCache = path.join(mirrorRoot, '.code-analysis');
	if (existsSync(analyzerCache)) violations.push('.code-analysis: analyzer cache must not ship');

	const docxRoot = path.join(mirrorRoot, 'docx-editor');
	if (!existsSync(docxRoot)) violations.push('docx-editor: missing runtime tree');
	else {
		const packagesRoot = path.join(docxRoot, 'packages');
		if (!existsSync(packagesRoot)) violations.push('docx-editor/packages: missing runtime packages');
		else {
			const packages = readdirSync(packagesRoot, { withFileTypes: true })
				.filter((entry) => entry.isDirectory())
				.map((entry) => entry.name)
				.sort();
			if (packages.join(',') !== CATALOG_DOCX_PACKAGES.slice().sort().join(',')) {
				violations.push(`docx-editor/packages: expected only ${CATALOG_DOCX_PACKAGES.join(', ')}`);
			}
			for (const pkg of CATALOG_DOCX_PACKAGES) {
				const pkgRoot = path.join(packagesRoot, pkg);
				if (!existsSync(pkgRoot)) continue;
				const distRoot = path.join(pkgRoot, 'dist');
				if (!existsSync(distRoot) || readdirSync(distRoot).length === 0) {
					violations.push(`docx-editor/packages/${pkg}/dist: missing runtime output`);
				}
				const packageJson = path.join(pkgRoot, 'package.json');
				if (existsSync(packageJson) && containsTypeMetadata(JSON.parse(readFileSync(packageJson, 'utf8')))) {
					violations.push(`docx-editor/packages/${pkg}/package.json: type metadata must not ship`);
				}
			}
		}
		violations.push(...collectCatalogSurfaceViolations(docxRoot).map((message) => `docx-editor/${message}`));
	}

	if (violations.length) {
		throw new Error(`[catalog-mirror] Unsafe public surface:\n${violations.join('\n')}`);
	}
}
