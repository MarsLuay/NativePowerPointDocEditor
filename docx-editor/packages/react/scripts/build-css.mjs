#!/usr/bin/env node
/**
 * Build React package CSS without storing `@tailwind` in shared source.
 * Writes a temp entry with `@tailwind utilities`, then prepends core
 * `editor.css` (theme tokens under `.docx-editor-root`) onto the output.
 *
 * Tailwind v3 does not reliably inline a trailing `@import` of the core sheet
 * into `--minify` output — without the prepend, portal menus use
 * `bg-popover` / `var(--doc-surface)` with undefined tokens → transparent
 * dropdowns over the page canvas.
 */
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
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

try {
	writeFileSync(
		entry,
		[
			'/* Generated Tailwind entry — not checked into source. */',
			'@tailwind utilities;',
			'',
		].join('\n'),
		'utf8',
	);

	const result = spawnSync(
		'bunx',
		['tailwindcss', '-c', config, '-i', entry, '-o', outCss, '--minify'],
		{ cwd: reactRoot, stdio: 'inherit' },
	);
	if (result.status !== 0) {
		process.exit(result.status === null ? 1 : result.status);
	}

	const core = readFileSync(coreEditorCss, 'utf8');
	const utilities = readFileSync(outCss, 'utf8');
	if (!core.includes('--doc-surface:') || !core.includes('.docx-editor-root')) {
		console.error('✘ core editor.css missing .docx-editor-root theme tokens');
		process.exit(1);
	}
	writeFileSync(outCss, `${core}\n${utilities}`, 'utf8');

	const merged = readFileSync(outCss, 'utf8');
	if (!merged.includes('--doc-surface:') || !merged.includes('--popover:')) {
		console.error('✘ dist/styles.css missing theme tokens after merge');
		process.exit(1);
	}
	console.log(`✓ wrote ${path.relative(reactRoot, outCss)} (${merged.length} bytes, tokens + utilities)`);
} finally {
	rmSync(tmpDir, { recursive: true, force: true });
}
