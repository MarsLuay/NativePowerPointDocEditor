#!/usr/bin/env node
/**
 * Rebuild docx-editor package dist (i18n → core → react) from source in ./docx-editor.
 * Obsidian users still only download main.js; this only regenerates local dist inputs.
 *
 * Requires: bun (docx-editor package scripts use bun/tsup). Install with:
 *   curl -fsSL https://bun.sh/install | bash
 * Or: npm install -g bun
 *
 * Catalog-safe public mirrors omit package `src/`. In that layout this script
 * verifies committed `packages/{core,react,i18n}/dist` and exits 0 (no bun).
 */
import { spawnSync } from 'node:child_process';
import { existsSync, readdirSync, renameSync, rmSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { CATALOG_DOCX_PACKAGES } from './lib/obsidian-catalog-mirror.mjs';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const monorepoRoot = path.join(projectRoot, 'docx-editor');

function run(cmd, args, cwd) {
	const result = spawnSync(cmd, args, { cwd, stdio: 'inherit', shell: false });
	if (result.status !== 0) {
		process.exit(result.status ?? 1);
	}
}

function which(bin) {
	const result = spawnSync(process.platform === 'win32' ? 'where' : 'which', [bin], {
		encoding: 'utf8',
	});
	if (result.status === 0) {
		return result.stdout.trim().split('\n')[0];
	}
	if (bin === 'bun') {
		const homeBun = path.join(process.env.HOME || '', '.bun', 'bin', 'bun');
		if (existsSync(homeBun)) {
			return homeBun;
		}
	}
	return null;
}

function assertCommittedDist() {
	for (const pkg of CATALOG_DOCX_PACKAGES) {
		const distDir = path.join(monorepoRoot, 'packages', pkg, 'dist');
		if (!existsSync(distDir) || readdirSync(distDir).length === 0) {
			console.error(`[build:docx-editor] Missing committed dist: ${distDir}`);
			process.exit(1);
		}
	}
}

if (!existsSync(path.join(monorepoRoot, 'package.json'))) {
	console.error(`[build:docx-editor] Missing ${monorepoRoot}`);
	process.exit(1);
}

const coreSrc = path.join(monorepoRoot, 'packages', 'core', 'src');
const hasPackageSource = existsSync(coreSrc);
if (!hasPackageSource) {
	console.log('[build:docx-editor] No package source (catalog surface). Using committed dist.');
	assertCommittedDist();
	console.log('[build:docx-editor] Dist OK.');
	process.exit(0);
}

const bun = which('bun');
if (!bun) {
	console.error('[build:docx-editor] bun is required to rebuild docx-editor packages from source.');
	console.error('Install: https://bun.sh  (or `npm install -g bun`), then re-run.');
	console.error('Until then, committed dist under docx-editor/packages/*/dist is used (same as before).');
	process.exit(1);
}

const monorepoNodeModules = path.join(monorepoRoot, 'node_modules');
const hadMonorepoNodeModules = existsSync(monorepoNodeModules);
if (!hadMonorepoNodeModules) {
	console.log('[build:docx-editor] Installing monorepo dependencies…');
	run(bun, ['install'], monorepoRoot);
}

// Isolate from plugin root node_modules (duplicate @types collide under tsup dts).
const env = {
	...process.env,
	NODE_PATH: '',
	npm_config_prefix: monorepoRoot,
};

function runIsolated(cmd, args, cwd) {
	const result = spawnSync(cmd, args, { cwd, stdio: 'inherit', shell: false, env });
	if (result.status !== 0) {
		process.exit(result.status ?? 1);
	}
}

const parentTrustedTypes = path.join(projectRoot, 'node_modules', '@types', 'trusted-types');
const parentTrustedTypesBak = `${parentTrustedTypes}.publish-bak`;
let hidParentTrustedTypes = false;
if (existsSync(parentTrustedTypes) && !existsSync(parentTrustedTypesBak)) {
	renameSync(parentTrustedTypes, parentTrustedTypesBak);
	hidParentTrustedTypes = true;
}

try {
	console.log('[build:docx-editor] Building i18n → core → react…');
	runIsolated(bun, ['run', '--filter', '@npde/docx-editor-i18n', 'build'], monorepoRoot);
	runIsolated(bun, ['run', '--filter', '@npde/docx-editor-core', 'build'], monorepoRoot);
	runIsolated(bun, ['run', '--filter', '@npde/docx-editor-react', 'build'], monorepoRoot);
} finally {
	if (hidParentTrustedTypes && existsSync(parentTrustedTypesBak)) {
		renameSync(parentTrustedTypesBak, parentTrustedTypes);
	}
}

// A clean dist build temporarily installs its dependencies. Preserve a caller's
// existing workspace, though: deleting it makes the monorepo's test/typecheck
// commands fail immediately after a successful package rebuild.
if (!hadMonorepoNodeModules && existsSync(monorepoNodeModules)) {
	console.log('[build:docx-editor] Removing build-only monorepo node_modules…');
	rmSync(monorepoNodeModules, { recursive: true, force: true });
}

// Keep sideEffects true for Obsidian esbuild (CSS / chunk side imports).
for (const pkg of ['i18n', 'core', 'react']) {
	const pkgPath = path.join(monorepoRoot, 'packages', pkg, 'package.json');
	run(process.execPath, ['-e', `
		const fs = require('fs');
		const p = ${JSON.stringify(pkgPath)};
		const j = JSON.parse(fs.readFileSync(p, 'utf8'));
		j.sideEffects = true;
		fs.writeFileSync(p, JSON.stringify(j, null, 2) + '\\n');
	`], projectRoot);
}

console.log('[build:docx-editor] Done. Run npm run build / npm run dev to refresh main.js.');
