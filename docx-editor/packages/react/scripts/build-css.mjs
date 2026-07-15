#!/usr/bin/env node
/**
 * Build React package CSS without storing `@tailwind` in shared source.
 * Writes a temp entry that prepends `@tailwind utilities` then imports core.
 */
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const reactRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const coreEditorCss = path.resolve(reactRoot, '../core/src/styles/editor.css');
const outCss = path.join(reactRoot, 'dist/styles.css');
const config = path.join(reactRoot, 'tailwind.config.js');

const tmpDir = mkdtempSync(path.join(tmpdir(), 'docx-editor-react-css-'));
const entry = path.join(tmpDir, 'entry.css');
const importPath = coreEditorCss.split(path.sep).join('/');

try {
	writeFileSync(
		entry,
		[
			'/* Generated Tailwind entry — not checked into source. */',
			'@tailwind utilities;',
			`@import "${importPath}";`,
			'',
		].join('\n'),
		'utf8',
	);

	const result = spawnSync(
		'bunx',
		['tailwindcss', '-c', config, '-i', entry, '-o', outCss, '--minify'],
		{ cwd: reactRoot, stdio: 'inherit' },
	);
	process.exit(result.status === null ? 1 : result.status);
} finally {
	rmSync(tmpDir, { recursive: true, force: true });
}
