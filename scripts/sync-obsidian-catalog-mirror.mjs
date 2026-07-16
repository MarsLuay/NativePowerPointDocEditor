#!/usr/bin/env node
/**
 * Sync plugin tree → clean-clone / standalone NativePowerPointDocEditor mirror
 * with catalog-safe exclusions (JS-only dist for docx-editor packages).
 *
 * Usage:
 *   node scripts/sync-obsidian-catalog-mirror.mjs <dest-dir>
 *
 * Prefers `rsync -a --delete`. Falls back to a filtered recursive copy when
 * rsync is unavailable. Always writes committed package dist for core/react/i18n,
 * then strips package typings (declaration files + package.json types fields).
 */
import { spawnSync } from 'node:child_process';
import {
	cpSync,
	existsSync,
	mkdtempSync,
	readFileSync,
	readdirSync,
	rmSync,
	writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
	CATALOG_DOCX_PACKAGES,
	CATALOG_MIRROR_RSYNC_EXCLUDES,
	assertCatalogMirrorSurface,
	pruneCatalogMirrorDocxTree,
	removePackageDeclarationFiles,
	stripPackageTypingsFromPackageJson,
	writeCatalogMirrorExcludeFile,
} from './lib/obsidian-catalog-mirror.mjs';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const destArg = process.argv[2];

if (!destArg) {
	console.error('Usage: node scripts/sync-obsidian-catalog-mirror.mjs <dest-dir>');
	process.exit(1);
}

const destRoot = path.resolve(destArg);
if (!existsSync(path.join(destRoot, '.git'))) {
	console.error(`[sync-catalog-mirror] Dest must be a git checkout: ${destRoot}`);
	process.exit(1);
}

for (const pkg of CATALOG_DOCX_PACKAGES) {
	const distDir = path.join(projectRoot, 'docx-editor', 'packages', pkg, 'dist');
	if (!existsSync(distDir) || readdirSync(distDir).length === 0) {
		console.error(`[sync-catalog-mirror] Missing package dist: ${distDir}`);
		console.error('Run: npm run build:docx-editor');
		process.exit(1);
	}
}

const excludeFile = path.join(mkdtempSync(path.join(tmpdir(), 'npde-catalog-rsync-')), 'excludes.txt');
writeCatalogMirrorExcludeFile(excludeFile);

function which(bin) {
	const result = spawnSync(process.platform === 'win32' ? 'where' : 'which', [bin], {
		encoding: 'utf8',
	});
	return result.status === 0 ? result.stdout.trim().split('\n')[0] : null;
}

function pathMatchesAnyGlob(relPosix, patterns) {
	for (const pattern of patterns) {
		let body = pattern.replace(/\/+$/, '');
		const dirOnly = pattern.endsWith('/');
		body = body
			.replace(/[.+^${}()|[\]\\]/g, '\\$&')
			.replace(/\*\*/g, '<<<GLOBSTAR>>>')
			.replace(/\*/g, '[^/]*')
			.replace(/<<<GLOBSTAR>>>/g, '.*');
		const re = new RegExp(`^${body}${dirOnly ? '(?:/.*)?$' : '$'}`);
		if (re.test(relPosix)) return true;
	}
	return false;
}

function filteredCopy(srcRoot, outRoot) {
	for (const name of readdirSync(outRoot)) {
		if (name === '.git') continue;
		rmSync(path.join(outRoot, name), { recursive: true, force: true });
	}
	cpSync(srcRoot, outRoot, {
		recursive: true,
		filter: (src) => {
			if (path.basename(src) === '.git') return false;
			const rel = path.relative(srcRoot, src).split(path.sep).join('/');
			if (!rel || rel === '.') return true;
			return !pathMatchesAnyGlob(rel, CATALOG_MIRROR_RSYNC_EXCLUDES);
		},
	});
}

function stripCatalogPackageTypings(pkgRoot) {
	removePackageDeclarationFiles(pkgRoot);
	const pkgJsonPath = path.join(pkgRoot, 'package.json');
	if (!existsSync(pkgJsonPath)) return;
	const raw = JSON.parse(readFileSync(pkgJsonPath, 'utf8'));
	const stripped = stripPackageTypingsFromPackageJson(raw);
	writeFileSync(pkgJsonPath, `${JSON.stringify(stripped, null, 2)}\n`, 'utf8');
}

const rsync = which('rsync');
if (rsync) {
	const result = spawnSync(
		rsync,
		[
			'-a',
			'--delete',
			`--exclude-from=${excludeFile}`,
			'--exclude=.git/',
			`${projectRoot}/`,
			`${destRoot}/`,
		],
		{ stdio: 'inherit' },
	);
	if (result.status !== 0) process.exit(result.status ?? 1);
} else {
	console.warn('[sync-catalog-mirror] rsync missing; using filtered recursive copy');
	filteredCopy(projectRoot, destRoot);
}

// rsync must retain the destination .git metadata, so excluded paths are pruned
// explicitly rather than using --delete-excluded. This also removes stale files
// left by earlier catalog syncs before the surface assertion runs.
rmSync(path.join(destRoot, '.code-analysis'), { recursive: true, force: true });
pruneCatalogMirrorDocxTree(destRoot);

// Strip declarations and type pointers from the retained runtime packages.
for (const pkg of CATALOG_DOCX_PACKAGES) {
	const pkgRoot = path.join(destRoot, 'docx-editor', 'packages', pkg);
	if (!existsSync(pkgRoot)) continue;
	stripCatalogPackageTypings(pkgRoot);
}

const marker = path.join(destRoot, 'docx-editor', 'CATALOG_SURFACE.md');
writeFileSync(
	marker,
	[
		'# Catalog surface (public mirror)',
		'',
		'This checkout ships **JS-only** `docx-editor/packages/{core,react,i18n}`',
		'(`.js` / `.mjs` / `.cjs` / `.css`). No package `.d.ts` and no `types`',
		'fields in those `package.json` files.',
		'',
		'Full TypeScript sources and package typings live in the ObsidianNotes vault',
		'authoritative tree only. Runtime JS is unchanged by sync.',
		'',
		'Rebuild from vault: `npm run build:docx-editor` then re-sync.',
		'',
	].join('\n'),
	'utf8',
);

assertCatalogMirrorSurface(destRoot);

console.log(`[sync-catalog-mirror] Synced catalog-safe tree → ${destRoot}`);
