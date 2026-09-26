import { spawn } from 'node:child_process';
import { existsSync, readdirSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const testsDirectory = path.join(projectRoot, 'tests');
const sourceOnlyTests = new Set([
	'docx-list-delete.test.mjs',
	'docx-empty-paragraph-font.test.mjs',
	'docx-plain-text-insert.test.mjs',
	'docx-persistence-gaps.test.mjs',
	'docx-enter-forward.test.mjs',
]);
const hasDocxSource = existsSync(
	path.join(projectRoot, 'docx-editor', 'packages', 'core', 'src'),
);
const testFiles = readdirSync(testsDirectory)
	.filter((name) => name.endsWith('.test.mjs'))
	.filter((name) => hasDocxSource || !sourceOnlyTests.has(name))
	.sort()
	.map((name) => path.join('tests', name));

const child = spawn(process.execPath, ['--test', ...testFiles], {
	cwd: projectRoot,
	stdio: 'inherit',
});
child.once('exit', (code, signal) => {
	if (signal) {
		process.kill(process.pid, signal);
		return;
	}
	process.exitCode = code ?? 1;
});
