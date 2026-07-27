#!/usr/bin/env node
/**
 * Sync the vendor-only plugin release tree to a standalone catalog mirror.
 *
 * Usage:
 *   node scripts/sync-obsidian-catalog-mirror.mjs <dest-dir>
 *
 * Editable DOCX sources belong on docx-editor-source. This export excludes a
 * local source checkout and requires the mirror to pass the review-surface
 * guard after synchronization.
 */
import { spawnSync } from 'node:child_process';
import {
	cpSync,
	existsSync,
	readdirSync,
	rmSync,
} from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const destArg = process.argv[2];
const excludedDirectories = new Set([
	'.git',
	'.code-analysis',
	'.serena',
	'node_modules',
	'results',
	'test-results',
	'docx-editor',
]);

if (!destArg) {
	console.error('Usage: node scripts/sync-obsidian-catalog-mirror.mjs <dest-dir>');
	process.exit(1);
}

const destRoot = path.resolve(destArg);
if (!existsSync(path.join(destRoot, '.git'))) {
	console.error(`[sync-catalog-mirror] Dest must be a git checkout: ${destRoot}`);
	process.exit(1);
}

function which(bin) {
	const result = spawnSync(process.platform === 'win32' ? 'where' : 'which', [bin], {
		encoding: 'utf8',
	});
	return result.status === 0 ? result.stdout.trim().split('\n')[0] : null;
}

function filteredCopy(srcRoot, outRoot) {
	for (const name of readdirSync(outRoot)) {
		if (name === '.git') continue;
		rmSync(path.join(outRoot, name), { recursive: true, force: true });
	}
	cpSync(srcRoot, outRoot, {
		recursive: true,
		filter: (src) => {
			if (excludedDirectories.has(path.basename(src))) return false;
			const rel = path.relative(srcRoot, src).split(path.sep).join('/');
			if (!rel || rel === '.') return true;
			return !rel.endsWith('.map');
		},
	});
}

const rsync = which('rsync');
if (rsync) {
	const result = spawnSync(
		rsync,
		[
			'-a',
			'--delete',
			'--exclude=.git',
			'--exclude=.code-analysis/',
			'--exclude=.serena/',
			'--exclude=node_modules/',
			'--exclude=results/',
			'--exclude=test-results/',
			'--exclude=docx-editor/',
			'--exclude=*.map',
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

// rsync retains excluded paths, so delete stale source trees explicitly.
rmSync(path.join(destRoot, '.code-analysis'), { recursive: true, force: true });
rmSync(path.join(destRoot, '.serena'), { recursive: true, force: true });
rmSync(path.join(destRoot, 'docx-editor'), { recursive: true, force: true });

const reviewCheck = spawnSync(
	process.execPath,
	['scripts/check-review-surface.mjs'],
	{ cwd: destRoot, stdio: 'inherit' },
);
if (reviewCheck.status !== 0) process.exit(reviewCheck.status ?? 1);

console.log(`[sync-catalog-mirror] Synced vendor-only release tree → ${destRoot}`);
