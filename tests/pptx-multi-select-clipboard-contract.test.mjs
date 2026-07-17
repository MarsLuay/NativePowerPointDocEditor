import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('POTX view routes multi-selection copy and paste through plural clipboard APIs', async () => {
  const source = await readFile(
    path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'),
    'utf8',
  );

  assert.match(source, /if \(this\.getSelectedIndices\(\)\.length > 0\) \{/);
  assert.match(source, /const shapeIndexes = this\.getSelectedIndices\(\);/);
  assert.match(source, /this\.engine\.copyShapes\(this\.currentSlide, shapeIndexes\)/);
  assert.match(source, /this\.engine\.pasteShapes\(this\.objectClipboard, this\.currentSlide\)/);
  assert.match(source, /this\.applyMultiSelection\(shapeIndexes\);/);
  assert.match(source, /count: shapeIndexes\.length,/);
});
