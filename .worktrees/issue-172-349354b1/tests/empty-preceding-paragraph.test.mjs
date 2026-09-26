import assert from 'node:assert/strict';
import { test } from 'node:test';
import { assertExportRoundTrips, loadEngine } from './helpers/pptx-action-harness.mjs';

test('Backspace engine mutation removes an empty predecessor and preserves the current paragraph', async () => {
  const engine = await loadEngine('features.pptx');
  const shapeIndex = await engine.insertTextBox(0);
  await engine.updateShapeText(0, shapeIndex, 'Keep this paragraph');
  const split = await engine.splitParagraph(0, shapeIndex, 0, 0);

  assert.equal(split.paragraphIndex, 1);
  assert.equal(engine.hasEmptyPrecedingParagraph(0, shapeIndex, 1), true);
  const result = await engine.removeEmptyPrecedingParagraph(0, shapeIndex, 1);

  assert.deepEqual(result, {
    removed: true,
    paragraphIndex: 0,
    beforeParagraphCount: 2,
    afterParagraphCount: 1,
    reason: 'removed',
  });
  assert.equal(engine.getParagraphRunText(0, shapeIndex, 0), 'Keep this paragraph');
  assert.equal(engine.hasEmptyPrecedingParagraph(0, shapeIndex, 0), false);
  await assertExportRoundTrips('remove empty preceding paragraph', engine);
});
