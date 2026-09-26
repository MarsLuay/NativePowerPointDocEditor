import assert from 'node:assert/strict';
import test from 'node:test';
import { loadTextUtilsModule } from './helpers/load-plugin-modules.mjs';

test('previousTextForInlineApply prefers session baseline over live SVG preview', async () => {
  const { previousTextForInlineApply } = await loadTextUtilsModule();

  const resolved = previousTextForInlineApply({
    sessionBaseline: 'old fourteen',
    targetText: 'stale target',
    liveSvgText: 'new twenty chars!!!',
  });

  assert.equal(resolved.previousText, 'old fourteen');
  assert.equal(resolved.source, 'session-baseline');
});

test('previousTextForInlineApply uses target text when no session baseline', async () => {
  const { previousTextForInlineApply } = await loadTextUtilsModule();

  const resolved = previousTextForInlineApply({
    targetText: 'paragraph baseline',
    liveSvgText: 'preview already updated',
  });

  assert.equal(resolved.previousText, 'paragraph baseline');
  assert.equal(resolved.source, 'target');
});

test('previousTextForInlineApply falls back to live SVG only last', async () => {
  const { previousTextForInlineApply } = await loadTextUtilsModule();

  const resolved = previousTextForInlineApply({
    liveSvgText: 'svg only',
  });

  assert.equal(resolved.previousText, 'svg only');
  assert.equal(resolved.source, 'live-svg');
});
