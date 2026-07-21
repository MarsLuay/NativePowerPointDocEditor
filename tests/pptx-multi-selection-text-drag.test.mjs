import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { loadNativePowerPointViewModule } from './helpers/load-plugin-modules.mjs';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('text within a multi-selection remains a group-drag surface', async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));

  view.selectedShapeIndices = new Set([2, 5]);
  view.selectedShapeIndex = null;

  assert.equal(view.shouldStartGroupDragFromText(2, false), true);
  assert.equal(view.shouldStartGroupDragFromText(5, true), false, 'additive selection must still toggle');
  assert.equal(view.shouldStartGroupDragFromText(7, false), false, 'outside text must select normally');

  view.selectedShapeIndices = new Set([2]);
  assert.equal(view.shouldStartGroupDragFromText(2, false), false, 'single selection still enters text editing');
});

test('SVG text pointer handling routes multi-selection to drag before inline editing', async () => {
  const source = await readFile(
    path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'),
    'utf8',
  );
  const pointerDownStart = source.indexOf("this.svgEl.addEventListener('pointerdown'");
  const groupDragStart = source.indexOf('if (this.shouldStartGroupDragFromText(shapeIndex, additive))', pointerDownStart);
  const inlineEditStart = source.indexOf('this.handleInlineTextPointerDown(event, target)', groupDragStart);

  assert.ok(groupDragStart > pointerDownStart, 'text pointer handler must check the drag condition');
  assert.ok(inlineEditStart > groupDragStart, 'group drag must win before inline text editing');
  assert.match(source.slice(groupDragStart, inlineEditStart), /this\.suppressNextClick = true/);
  assert.match(source.slice(groupDragStart, inlineEditStart), /this\.startGroupDrag\(event\)/);
});
