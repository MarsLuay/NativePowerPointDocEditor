import assert from 'node:assert/strict';
import test from 'node:test';
import {
  formatPptxJsEngineVersionMismatch,
  readInstalledPptxSvgVersion,
  readPptxJsEngineVersion,
  resolveProjectRoot,
} from '../scripts/lib/pptx-svg-version.mjs';

const projectRoot = resolveProjectRoot(import.meta.url);

test('local pptxJsEngine version stamp matches installed pptx-svg', () => {
  const installed = readInstalledPptxSvgVersion(projectRoot);
  const local = readPptxJsEngineVersion(projectRoot);
  assert.equal(local, installed);
});

test('formatPptxJsEngineVersionMismatch mentions regenerate path', () => {
  const message = formatPptxJsEngineVersionMismatch({
    installed: '0.5.11',
    local: '0.5.10',
  });
  assert.match(message, /pptxJsEngine\.mjs/);
  assert.match(message, /0\.5\.11/);
  assert.match(message, /0\.5\.10/);
});
