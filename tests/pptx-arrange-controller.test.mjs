import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadArrangeControllerModule } from './helpers/load-plugin-modules.mjs';

test('ArrangeController ignores a second reorder while the first reload is in flight', async () => {
  const { ArrangeController } = await loadArrangeControllerModule();
  let resolveCommand;
  let commandCount = 0;
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
      applyCommand: () => {
        commandCount += 1;
        return new Promise((resolve) => { resolveCommand = resolve; });
      },
    },
  };
  const controller = new ArrangeController(host);

  const first = controller.reorderSelection('backward');
  const duplicate = controller.reorderSelection('backward');
  await Promise.resolve();
  assert.equal(commandCount, 1, 'duplicate click must not start a stale-index mutation');

  resolveCommand([16]);
  await Promise.all([first, duplicate]);
  assert.deepEqual(selection, [[16]]);
});
