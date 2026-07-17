import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('POTX view deletes a multi-selection through one batched engine mutation', async () => {
  const source = await readFile(
    path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'),
    'utf8',
  );
  const methodStart = source.indexOf('private async deleteSelectedShapes(): Promise<void>');
  const methodEnd = source.indexOf('\n  private async copySelectedShape()', methodStart);
  const method = source.slice(methodStart, methodEnd);

  assert.ok(methodStart >= 0 && methodEnd > methodStart, 'multi-delete view method is present');
  assert.match(method, /this\.engine\.deleteShapes\(this\.currentSlide, indices\)/);
  assert.doesNotMatch(method, /for \(const index of indices\).*?this\.engine\.deleteShape/s);
  assert.match(method, /this\.scheduleThumbnailRefresh\(this\.currentSlide\)/);
  assert.doesNotMatch(method, /await this\.renderThumbnails\(\)/);
});
