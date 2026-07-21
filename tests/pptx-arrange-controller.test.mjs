import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadArrangeControllerModule } from './helpers/load-plugin-modules.mjs';

test('ArrangeController ignores a second reorder while the first reload is in flight', async () => {
  const { ArrangeController } = await loadArrangeControllerModule();
  let resolveCommand;
  let commandCount = 0;
  const commands = [];
  const selection = [];
  const host = {
    engine: {},
    currentSlide: 0,
    ensureEditable: () => true,
    canEdit: () => true,
    getSelectedIndices: () => [17],
    captureHistoryEntry: async () => ({ label: 'Reorder objects' }),
    recordHistoryEntry: () => {},
    markDirty: () => {},
    renderCurrentSlide: async () => true,
    renderThumbnails: async () => {},
    applyMultiSelection: (indexes) => selection.push(indexes),
    selectShape: () => {},
    commitGroupTransforms: async () => {},
    createIconButton: () => ({}),
    updateToolbarButton: () => {},
    session: {
      applyCommand: (command) => {
        commandCount += 1;
        commands.push(command);
        return new Promise((resolve) => { resolveCommand = resolve; });
      },
    },
  };
  const controller = new ArrangeController(host);

  const first = controller.reorderSelection('backward');
  const duplicate = controller.reorderSelection('backward');
  await Promise.resolve();
  assert.equal(commandCount, 1, 'duplicate click must not start a stale-index mutation');
  assert.equal(commands[0].intersectingOnly, true, 'toolbar reorder must use overlap-aware ordering');

  resolveCommand([16]);
  await Promise.all([first, duplicate]);
  assert.deepEqual(selection, [[16]]);
});

test('ArrangeController does not record history or re-render when no object overlaps', async () => {
  const { ArrangeController } = await loadArrangeControllerModule();
  let captureCount = 0;
  let historyCount = 0;
  let renderCount = 0;
  const host = {
    engine: {},
    currentSlide: 0,
    ensureEditable: () => true,
    canEdit: () => true,
    getSelectedIndices: () => [17],
    captureHistoryEntry: async () => { captureCount += 1; return { label: 'Reorder objects' }; },
    recordHistoryEntry: () => { historyCount += 1; },
    markDirty: () => {},
    renderCurrentSlide: async () => { renderCount += 1; return true; },
    renderThumbnails: async () => {},
    applyMultiSelection: () => {},
    selectShape: () => {},
    commitGroupTransforms: async () => {},
    createIconButton: () => ({}),
    updateToolbarButton: () => {},
    session: { applyCommand: async () => null },
  };
  const controller = new ArrangeController(host);

  await controller.reorderSelection('forward');

  assert.equal(captureCount, 1, 'history snapshot remains necessary before a possible mutation');
  assert.equal(historyCount, 0, 'an overlap-free reorder must not create an undo step');
  assert.equal(renderCount, 0, 'an overlap-free reorder must not trigger a slide reload');
});
