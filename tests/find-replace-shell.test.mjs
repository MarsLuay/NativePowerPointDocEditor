import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

async function loadFindReplaceShellModule() {
  const outdir = await mkdtemp(path.join(tmpdir(), 'npde-find-replace-shell-test-'));
  const outfile = path.join(outdir, 'find-replace-shell.cjs');
  try {
    await build({
      absWorkingDir: projectRoot,
      entryPoints: ['src/find/findReplaceShell.ts'],
      bundle: true,
      format: 'cjs',
      logLevel: 'silent',
      outfile,
      platform: 'node',
      target: 'node22',
    });
    const mod = await import(pathToFileURL(outfile).href);
    return mod;
  } finally {
    await rm(outdir, { recursive: true, force: true }).catch(() => {});
  }
}

const { formatFindResultStatus, wrapMatchIndex } = await loadFindReplaceShellModule();

const labels = {
  noSearch: 'No search term',
  noMatches: 'No matches found',
  resultCount: (current, total) => `${current} of ${total}`,
};

test('formatFindResultStatus returns noSearch when query is empty or whitespace', () => {
  assert.equal(formatFindResultStatus('', 0, 5, labels), 'No search term');
  assert.equal(formatFindResultStatus('   ', 0, 5, labels), 'No search term');
  assert.equal(formatFindResultStatus('\t\n', 2, 10, labels), 'No search term');
});

test('formatFindResultStatus returns noMatches when matchCount is zero or negative', () => {
  assert.equal(formatFindResultStatus('test', 0, 0, labels), 'No matches found');
  assert.equal(formatFindResultStatus('test', 0, -1, labels), 'No matches found');
});

test('formatFindResultStatus formats result count correctly with 1-based indexing', () => {
  assert.equal(formatFindResultStatus('test', 0, 5, labels), '1 of 5');
  assert.equal(formatFindResultStatus('test', 2, 5, labels), '3 of 5');
  assert.equal(formatFindResultStatus('test', 4, 5, labels), '5 of 5');
});

test('wrapMatchIndex handles zero or negative match count', () => {
  assert.equal(wrapMatchIndex(0, 1, 0), 0);
  assert.equal(wrapMatchIndex(2, -1, -5), 0);
});

test('wrapMatchIndex navigates forward and backward with wrap-around', () => {
  // Match count = 3 (valid indices 0, 1, 2)
  assert.equal(wrapMatchIndex(0, 1, 3), 1);
  assert.equal(wrapMatchIndex(1, 1, 3), 2);
  assert.equal(wrapMatchIndex(2, 1, 3), 0); // Wraps to start

  assert.equal(wrapMatchIndex(0, -1, 3), 2); // Wraps to end
  assert.equal(wrapMatchIndex(2, -1, 3), 1);
  assert.equal(wrapMatchIndex(1, -1, 3), 0);
});
