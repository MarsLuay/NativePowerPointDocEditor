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
  const methodStart = source.indexOf('private async deleteSelectedShapesUnlocked(): Promise<void>');
  const methodEnd = source.indexOf('\n  private async copySelectedShape()', methodStart);
  const method = source.slice(methodStart, methodEnd);

  assert.ok(methodStart >= 0 && methodEnd > methodStart, 'multi-delete unlocked method is present');
  assert.match(method, /this\.engine\.deleteShapes\(this\.currentSlide, indices\)/);
  assert.doesNotMatch(method, /for \(const index of indices\).*?this\.engine\.deleteShape/s);
  assert.match(method, /this\.scheduleThumbnailRefresh\(this\.currentSlide\)/);
  assert.doesNotMatch(method, /await this\.renderThumbnails\(\)/);
  assert.match(method, /prepareForShapeDeletion\(indices\)/);
  assert.match(method, /clearSelection\(\{ skipTextCommit: true \}\)/);
});

test('overlapping shape deletes queue a fresh-selection rerun', async () => {
  const source = await readFile(
    path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'),
    'utf8',
  );
  const methodStart = source.indexOf('private async runExclusiveShapeDeletion(run: () => Promise<void>): Promise<void>');
  const methodEnd = source.indexOf('\n  private async deleteSelectedShape()', methodStart);
  const method = source.slice(methodStart, methodEnd);
  const singleStart = source.indexOf('private async deleteSelectedShape(): Promise<void>');
  const singleEnd = source.indexOf('\n  private async deleteSelectedShapes(): Promise<void>', singleStart);
  const single = source.slice(singleStart, singleEnd);

  assert.ok(methodStart >= 0 && methodEnd > methodStart, 'exclusive shape deletion helper is present');
  assert.match(method, /shapeDeletionRerunRequested = true/);
  assert.match(method, /Queued overlapping PowerPoint shape delete/);
  assert.match(method, /while \(this\.shapeDeletionRerunRequested\)/);
  assert.doesNotMatch(method, /Ignored overlapping PowerPoint shape delete/);
  // Indices must be read inside the exclusive callback so a queued rerun
  // cannot reuse a pre-renumber snapshot from the first Delete keypress.
  assert.match(single, /const shapeIndex = this\.selectedShapeIndex;/);
  assert.ok(
    single.indexOf('runExclusiveShapeDeletion') < single.indexOf('const shapeIndex = this.selectedShapeIndex'),
    'single-delete resolves shapeIndex inside the exclusive callback',
  );
});

test('cut and duplicate use multi-select indices like copy', async () => {
  const source = await readFile(
    path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'),
    'utf8',
  );
  const cutStart = source.indexOf('private async cutSelectedShape(): Promise<void>');
  const cutEnd = source.indexOf('\n  private async pasteWithoutFormatting()', cutStart);
  const cut = source.slice(cutStart, cutEnd);
  const duplicateStart = source.indexOf('private async duplicateSelectedShape(): Promise<void>');
  const duplicateEnd = source.indexOf('\n  private async cutSelectedShape()', duplicateStart);
  const duplicate = source.slice(duplicateStart, duplicateEnd);

  assert.match(cut, /getSelectedIndices\(\)/);
  assert.match(cut, /copyShapes\(this\.currentSlide, shapeIndexes\)/);
  assert.match(duplicate, /getSelectedIndices\(\)/);
  assert.match(duplicate, /pasteShapes\(clipboard, this\.currentSlide\)/);
  assert.match(
    source,
    /getSelectedShapeElement\(\)[\s\S]*parentElement\?\.closest\('g\[data-ooxml-shape-idx\]'\)/,
  );
});

test('whole-shape and empty-paragraph previews hide list marker containers', async () => {
  const source = await readFile(
    path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'),
    'utf8',
  );
  const wholeStart = source.indexOf('private beginInlineWholeShapeReplacement(target: ShapeTextEditTarget): boolean');
  const wholeEnd = source.indexOf('\n  private restoreInlineWholeShapeReplacementPreview', wholeStart);
  const whole = source.slice(wholeStart, wholeEnd);
  const syncStart = source.indexOf('private syncShapeParagraphPreview(');
  const syncEnd = source.indexOf('\n  private replaceLiveShapeTextFrame(', syncStart);
  const sync = source.slice(syncStart, syncEnd);

  assert.ok(wholeStart >= 0 && wholeEnd > wholeStart, 'whole-shape replacement method is present');
  assert.match(whole, /setLiveListMarkersHidden\(target\.shapeIndex, true\)/);
  assert.match(source, /private setLiveListMarkersHidden\(/);
  assert.match(source, /data-ooxml-preview-marker-hidden/);
  assert.match(sync, /setLiveListMarkersHidden\(target\.shapeIndex, logicalEmpty, target\.paragraphIndex\)/);
  assert.match(sync, /inlineWholeShapeReplacement\?\.shapeIndex === target\.shapeIndex/);
});

test('Delete key routes filmstrip focus to slide deletion', async () => {
  const viewSource = await readFile(
    path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'),
    'utf8',
  );
  const filmstripSource = await readFile(
    path.join(projectRoot, 'src/powerpoint/slideFilmstripController.ts'),
    'utf8',
  );
  const handlerStart = viewSource.indexOf('private registerKeyboardHandlers(): void');
  const handlerEnd = viewSource.indexOf('\n  private ', handlerStart + 1);
  const handler = viewSource.slice(handlerStart, handlerEnd);

  assert.ok(handlerStart >= 0 && handlerEnd > handlerStart, 'keyboard handler is present');
  assert.match(handler, /lastInteractionRegion === 'thumbnails'/);
  assert.doesNotMatch(
    handler,
    /lastInteractionRegion === 'thumbnails'[\s\S]*&& !hasShapeSelection/,
  );
  assert.match(handler, /void this\.deleteSelectedShape\(\)/);
  assert.doesNotMatch(
    viewSource,
    /registerDomEvent\(sidebar, 'pointerdown'/,
    'sidebar pointerdown must not steal Delete from the canvas',
  );
  assert.match(
    filmstripSource,
    /index === fromSlide[\s\S]*clearSelection\(\)/,
    're-clicking the active thumbnail clears stale shape selection',
  );
});
