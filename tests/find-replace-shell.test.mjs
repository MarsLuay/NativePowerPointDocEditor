import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

let cachedModule;

async function loadFindReplaceShellModule() {
  if (cachedModule) return cachedModule;
  const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-find-replace-shell-'));
  const outfile = path.join(outputDirectory, 'find-replace-shell.cjs');
  await build({
    entryPoints: [path.join(projectRoot, 'src/find/findReplaceShell.ts')],
    bundle: true,
    format: 'cjs',
    logLevel: 'silent',
    outfile,
    platform: 'node',
    target: 'node22',
  });
  cachedModule = require(outfile);
  return cachedModule;
}

test('wrapMatchIndex handles empty or zero match counts', async () => {
  const { wrapMatchIndex } = await loadFindReplaceShellModule();

  assert.equal(wrapMatchIndex(0, 1, 0), 0);
  assert.equal(wrapMatchIndex(0, -1, 0), 0);
  assert.equal(wrapMatchIndex(2, 1, -1), 0);
  assert.equal(wrapMatchIndex(5, -1, -5), 0);
});

test('wrapMatchIndex wraps correctly in forward and backward directions', async () => {
  const { wrapMatchIndex } = await loadFindReplaceShellModule();

  // Forward navigation (+1)
  assert.equal(wrapMatchIndex(0, 1, 3), 1);
  assert.equal(wrapMatchIndex(1, 1, 3), 2);
  assert.equal(wrapMatchIndex(2, 1, 3), 0); // Wrap to start

  // Backward navigation (-1)
  assert.equal(wrapMatchIndex(0, -1, 3), 2); // Wrap to end
  assert.equal(wrapMatchIndex(2, -1, 3), 1);
  assert.equal(wrapMatchIndex(1, -1, 3), 0);

  // Zero direction (stay at current index)
  assert.equal(wrapMatchIndex(1, 0, 3), 1);

  // Single match count
  assert.equal(wrapMatchIndex(0, 1, 1), 0);
  assert.equal(wrapMatchIndex(0, -1, 1), 0);

  // Large match count
  assert.equal(wrapMatchIndex(9, 1, 10), 0);
  assert.equal(wrapMatchIndex(0, -1, 10), 9);
});

test('formatFindResultStatus formats search status labels correctly', async () => {
  const { formatFindResultStatus } = await loadFindReplaceShellModule();
  const labels = {
    noSearch: 'No search term',
    noMatches: 'No matches found',
    resultCount: (current, total) => `${current} of ${total}`,
  };

  // Empty or whitespace queries
  assert.equal(formatFindResultStatus('', 0, 5, labels), 'No search term');
  assert.equal(formatFindResultStatus('   ', 0, 5, labels), 'No search term');

  // No matches
  assert.equal(formatFindResultStatus('test', 0, 0, labels), 'No matches found');
  assert.equal(formatFindResultStatus('test', 0, -1, labels), 'No matches found');

  // Active matches (currentIndex is 0-based, result display is 1-based)
  assert.equal(formatFindResultStatus('test', 0, 5, labels), '1 of 5');
  assert.equal(formatFindResultStatus('test', 2, 5, labels), '3 of 5');
  assert.equal(formatFindResultStatus('test', 4, 5, labels), '5 of 5');
});
