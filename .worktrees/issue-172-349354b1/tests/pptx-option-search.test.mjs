import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

test('PPTX header Search filters editor actions while presentation Find remains available', () => {
  const menuBar = readFileSync(join(root, 'src/powerpoint/menuBarController.ts'), 'utf8');
  const view = readFileSync(join(root, 'src/powerpoint/ui/NativePowerPointView.ts'), 'utf8');

  assert.match(menuBar, /kind: 'search'/);
  assert.match(menuBar, /native-powerpoint-menubar-search-input/);
  assert.match(menuBar, /\[entry\.label, \.\.\.\(entry\.keywords \?\? \[\]\)\]/);
  assert.match(menuBar, /Dispatching PowerPoint option-search command/);
  assert.match(view, /kind: 'search',\s+label: this\.tb\('search'\)/);
  assert.match(view, /addMenuItems\(this\.tb\('file'\), this\.getFileMenuItems\(\)\)/);
  assert.match(view, /id: 'reset-image-aspect-ratio'/);
  assert.match(view, /label: this\.tb\('resetImageAspectRatio'\)/);
  assert.match(view, /keywords: \['image', 'picture', 'aspect', 'ratio', 'reset', 'default'\]/);
  assert.match(view, /resetSelectedImageAspectRatio\(selectedImageIndex\)/);
  assert.match(view, /id: 'find-in-presentation'/);
  assert.match(view, /onClick: \(\) => this\.findController\.open\(\)/);
});
