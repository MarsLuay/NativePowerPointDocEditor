import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
  formatPptxJsEngineVersionMismatch,
  readInstalledPptxSvgVersion,
  readVendoredPptxJsEngineVersion,
  resolveProjectRoot,
} from '../scripts/lib/pptx-svg-version.mjs';

const projectRoot = resolveProjectRoot(import.meta.url);

test('vendored pptx-js-engine version stamp matches installed pptx-svg', () => {
  const installed = readInstalledPptxSvgVersion(projectRoot);
  const vendored = readVendoredPptxJsEngineVersion(projectRoot);
  assert.equal(vendored, installed);
});

test('version mismatch helper explains how to regenerate', () => {
  const message = formatPptxJsEngineVersionMismatch({
    installed: '0.6.0',
    vendored: '0.5.10',
  });

  assert.match(message, /0\.6\.0/);
  assert.match(message, /0\.5\.10/);
  assert.match(message, /npm run regen:pptx-js/);
});
