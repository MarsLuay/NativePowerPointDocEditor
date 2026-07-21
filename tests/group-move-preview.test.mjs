import assert from 'node:assert/strict';
import test from 'node:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

test('group move live-previews selected SVG shapes, not only selection outlines', () => {
  const view = readFileSync(join(root, 'src/powerpoint/ui/NativePowerPointView.ts'), 'utf8');
  assert.match(view, /applyGroupMovePreview\(/);
  assert.match(view, /this\.applyGroupMovePreview\(/);
  assert.match(
    view,
    /shape\.setAttribute\('transform', original \? `\$\{transform\} \$\{original\}` : transform\)/,
  );
  assert.match(view, /restoreGroupShapePreviews\(groupDrag\)/);
  assert.match(view, /op: 'group-move-preview'/);

  const updateGroupDrag = view.slice(
    view.indexOf('private updateGroupDrag'),
    view.indexOf('private applyGroupMovePreview'),
  );
  assert.match(updateGroupDrag, /applyGroupMovePreview/);
});
