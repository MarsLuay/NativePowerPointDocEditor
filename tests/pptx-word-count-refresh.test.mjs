import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { test } from 'node:test';

import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
let modulePromise;

async function loadWordCountRefreshModule() {
  modulePromise ??= bundleSource(
    'src/powerpoint/wordCountRefresh.ts',
    'word-count-refresh.cjs',
  ).then((outfile) => require(outfile));
  return modulePromise;
}

test('presentation word count skips formatting-only PowerPoint commands', async () => {
  const { commandMayChangePresentationWordCount } = await loadWordCountRefreshModule();

  assert.equal(commandMayChangePresentationWordCount({
    type: 'set-run-style-ranges',
    slideIndex: 0,
    shapeIndex: 18,
    ranges: [{ paragraphIndex: 0, startOffset: 0, endOffset: 4 }],
    change: { fontSizePt: 28 },
  }), false);
  assert.equal(commandMayChangePresentationWordCount({
    type: 'set-paragraph-alignment',
    slideIndex: 0,
    shapeIndex: 18,
    paragraphIndex: 0,
    align: 'ctr',
  }), false);
  assert.equal(commandMayChangePresentationWordCount({
    type: 'apply-list-style',
    slideIndex: 0,
    shapeIndex: 18,
    paragraphIndex: 0,
    style: 'bullet',
  }), false);
});

test('presentation word count refreshes after text-changing PowerPoint commands', async () => {
  const { commandMayChangePresentationWordCount } = await loadWordCountRefreshModule();

  assert.equal(commandMayChangePresentationWordCount({
    type: 'update-paragraph-text',
    slideIndex: 0,
    shapeIndex: 18,
    paragraphIndex: 0,
    text: 'Updated poster copy',
  }), true);
  assert.equal(commandMayChangePresentationWordCount({
    type: 'delete-shape',
    slideIndex: 0,
    shapeIndex: 18,
  }), true);
});
