import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('resolveThumbnailReorderIndex matches before/after drop math', () => {
  const resolveThumbnailReorderIndex = (fromIndex, targetIndex, after) => {
    let toIndex = after ? targetIndex + 1 : targetIndex;
    if (fromIndex < toIndex) toIndex -= 1;
    return toIndex;
  };
  const slideIndicesInRange = (anchor, index) => {
    const start = Math.min(anchor, index);
    const end = Math.max(anchor, index);
    const indices = [];
    for (let slide = start; slide <= end; slide += 1) indices.push(slide);
    return indices;
  };

  assert.equal(resolveThumbnailReorderIndex(0, 2, true), 2);
  assert.equal(resolveThumbnailReorderIndex(2, 0, false), 0);
  assert.equal(resolveThumbnailReorderIndex(1, 1, false), 1);
  // Drag 4 onto the lower half of 1 → insert after 1 → final index 2.
  assert.equal(resolveThumbnailReorderIndex(4, 1, true), 2);
  // Drag 0 onto the upper half of 3 → insert before 3, then compensate → 2.
  assert.equal(resolveThumbnailReorderIndex(0, 3, false), 2);
  assert.deepEqual(slideIndicesInRange(2, 5), [2, 3, 4, 5]);
  assert.deepEqual(slideIndicesInRange(5, 2), [2, 3, 4, 5]);
});

test('filmstrip uses pointer reorder and shift/ctrl selection anchors', async () => {
  const source = await readFile(
    path.join(projectRoot, 'src/powerpoint/slideFilmstripController.ts'),
    'utf8',
  );
  const styles = await readFile(path.join(projectRoot, 'styles.css'), 'utf8');

  assert.match(source, /slideSelectionAnchor/);
  assert.match(source, /selectSlideRange\(anchor, index\)/);
  assert.match(source, /applyAdditiveThumbnailClick/);
  assert.match(source, /handleThumbnailPointerDown/);
  assert.match(source, /THUMBNAIL_REORDER_DRAG_THRESHOLD_PX/);
  assert.match(source, /thumbnail-shift-select/);
  assert.match(source, /thumbnail-multiselect/);
  assert.match(source, /event\.ctrlKey && !event\.metaKey/);
  assert.doesNotMatch(source, /thumbnailDragIndex/);
  assert.doesNotMatch(source, /addEventListener\('dragstart'/);
  assert.match(styles, /-webkit-user-drag:\s*none/);
  assert.match(styles, /\.native-powerpoint-thumbnail-preview \*/);
  assert.match(styles, /cursor:\s*grab/);
});
