import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const viewSource = await readFile(
  path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'),
  'utf8',
);

function methodSource(start, end) {
  const startAt = viewSource.indexOf(start);
  const endAt = viewSource.indexOf(end, startAt);
  assert.ok(startAt >= 0, `missing ${start}`);
  assert.ok(endAt > startAt, `missing boundary ${end}`);
  return viewSource.slice(startAt, endAt);
}

test('single-shape transform mirrors its current thumbnail instead of scheduling a full render', () => {
  const source = methodSource('  private async commitTransform(', '  private getSvgPoint(');

  assert.match(source, /mirrorCurrentThumbnailShape\(shapeIndex, 'single-shape-live-transform'\)/);
  assert.match(source, /mirrorCurrentThumbnailShape\(shapeIndex, 'single-shape-rendered-transform'\)/);
  assert.doesNotMatch(source, /scheduleThumbnailRefresh\(this\.currentSlide\)/);
});

test('group transforms mirror changed shape groups instead of scheduling a full thumbnail render', () => {
  const source = methodSource('  private async commitGroupTransforms(', '  private getSelectedBox(');

  assert.match(source, /syncCurrentThumbnailShapes\(changes\.map\(\(change\) => change\.shapeIndex\)\)/);
  assert.doesNotMatch(source, /scheduleThumbnailRefresh\(this\.currentSlide\)/);
});

test('active slide thumbnails clone the visible canvas instead of re-rendering the poster', async () => {
  const controllerSource = await readFile(
    path.join(projectRoot, 'src/powerpoint/slideFilmstripController.ts'),
    'utf8',
  );
  const start = controllerSource.indexOf('  async renderThumbnailAt(index: number): Promise<void>');
  const end = controllerSource.indexOf('  /** Clone the current canvas SVG', start);
  assert.ok(start >= 0 && end > start, 'renderThumbnailAt source boundaries must exist');
  const method = controllerSource.slice(start, end);

  assert.match(method, /index === this\.host\.currentSlide/);
  assert.match(method, /this\.cloneActiveSlideSvgForThumbnail\(\)/);
  assert.match(method, /source === 'engine-render'/);
  assert.match(method, /source,/);
});
