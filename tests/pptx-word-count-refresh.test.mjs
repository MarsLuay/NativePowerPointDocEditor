import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';

import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');
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
  assert.equal(commandMayChangePresentationWordCount({
    type: 'apply-list-style-ranges',
    slideIndex: 0,
    shapeIndex: 18,
    ranges: [{ paragraphIndex: 0, start: 0, end: 4 }],
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

test('presentation word count is deferred while inline editing so Enter does not render the whole deck', () => {
  const view = readFileSync(join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'), 'utf8');
  const sessionHandler = view.slice(
    view.indexOf('private handlePresentationSessionEvent'),
    view.indexOf('private schedulePresentationWordCountRefresh'),
  );

  assert.match(sessionHandler, /this\.schedulePresentationWordCountRefresh\('session-save'\)/);
  assert.doesNotMatch(sessionHandler, /this\.refreshPresentationWordCount\(\)/);
  assert.match(view, /Deferred PowerPoint word count refresh while inline editing/);
  assert.match(view, /this\.schedulePresentationWordCountRefresh\('inline-editor-closed'\)/);
  assert.match(view, /this\.cancelPresentationWordCountRefresh\(\);/);
});

test('incremental PowerPoint shape replacement keeps non-chart display work local', () => {
  const view = readFileSync(join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'), 'utf8');
  const replacement = view.slice(
    view.indexOf('private replaceShapeFragmentsInPlace'),
    view.indexOf('private async renderEditedShape'),
  );

  assert.match(replacement, /normalizeShapeForDisplay\(replacement\)/);
  assert.match(replacement, /this\.markGeneratedTextEditability\(replacement\)/);
  assert.doesNotMatch(replacement, /normalizeSvgForDisplay\(svg\)/);
  assert.match(replacement, /replacement\.getAttribute\('data-ooxml-shape-type'\) === 'chart'/);
});

test('Enter uses the local paragraph preview before falling back to the expensive PowerPoint renderer', () => {
  const view = readFileSync(join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'), 'utf8');
  const split = view.slice(
    view.indexOf('private async splitInlineParagraph'),
    view.indexOf('/**\n   * Insert a lightweight local paragraph preview'),
  );

  assert.match(split, /this\.previewInlineParagraphSplit\(/);
  assert.match(split, /const rendered = previewed \|\| await this\.renderEditedShape\(target\.shapeIndex\)/);
  assert.match(split, /renderStrategy: previewed \? 'live-dom' : 'engine-render'/);
});

test('Backspace structural edits use local paragraph previews before the renderer fallback', () => {
  const view = readFileSync(join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'), 'utf8');
  const removeEmpty = view.slice(
    view.indexOf('private async removeInlineEmptyPrecedingParagraph'),
    view.indexOf('private async mergeInlinePrecedingParagraph'),
  );
  const merge = view.slice(
    view.indexOf('private async mergeInlinePrecedingParagraph'),
    view.indexOf('/** Shift local SVG paragraph identities'),
  );

  assert.match(removeEmpty, /this\.previewInlineEmptyPrecedingParagraphRemoval\(/);
  assert.match(merge, /this\.previewInlinePrecedingParagraphMerge\(/);
  assert.match(merge, /const rendered = previewed \|\| await this\.renderEditedShape\(target\.shapeIndex\)/);
  assert.doesNotMatch(view, /if \(!canPreview \|\| currentText\.length > 0\) return false;/);
});
