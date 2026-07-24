#!/usr/bin/env node
/**
 * Vault monorepo: full `tsc -noEmit -skipLibCheck`.
 * Catalog surface (JS-only package dist, no `.d.ts`): skip tsc — esbuild resolves runtime JS.
 */
import { existsSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const catalogSurface = path.join(projectRoot, 'docx-editor', 'CATALOG_SURFACE.md');

if (existsSync(catalogSurface)) {
	console.log(
		'[typecheck] Catalog surface (JS-only packages) — skipping tsc; esbuild uses package JS.',
	);
	process.exit(0);
}

const result = spawnSync(
	process.platform === 'win32' ? 'npx.cmd' : 'npx',
	['tsc', '-noEmit', '-skipLibCheck'],
	{ cwd: projectRoot, stdio: 'inherit', shell: process.platform === 'win32' },
);
process.exit(result.status ?? 1);
