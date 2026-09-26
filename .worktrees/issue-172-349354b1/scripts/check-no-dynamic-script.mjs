// Obsidian's automated plugin review rejects main.js when it contains dynamic
// <script> element creation (createElement("script")). Fail the build if the
// pattern leaks through after a dependency upgrade.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';

const bundlePath = path.resolve('main.js');
const source = await readFile(bundlePath, 'utf8');
const scriptCreateRe = /createElement\(\s*["']script["']\s*\)/g;
const matches = [...source.matchAll(scriptCreateRe)];

assert.equal(
	matches.length,
	0,
	`${bundlePath} still contains ${matches.length} dynamic <script> element creation(s). ` +
		'Obsidian plugin review will reject this build. Check react-dom / polyfill shims.',
);

console.log(`No dynamic script-element creation in ${bundlePath}.`);
