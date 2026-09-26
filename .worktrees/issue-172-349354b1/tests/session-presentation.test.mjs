import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';
import { build } from 'esbuild';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);

let modulePromise;
async function loadPresentationSessionModule() {
  modulePromise ??= (async () => {
    const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-presentation-session-'));
    const outfile = path.join(outputDirectory, 'PresentationSession.cjs');
    await build({
      entryPoints: [path.join(projectRoot, 'src/powerpoint/session/PresentationSession.ts')],
      bundle: true,
      format: 'cjs',
      logLevel: 'silent',
      outfile,
      platform: 'node',
      target: 'node22',
      external: ['obsidian', 'pptx-svg', 'pptx-svg/wasm'],
    });

    const Module = require('node:module');
    const originalLoad = Module._load;
    Module._load = function load(request, parent, isMain) {
      if (request === 'obsidian') return { Notice: class Notice {} };
      if (request === 'pptx-svg') {
        return {
          PptxRenderer: class PptxRenderer {},
          buildZip: async () => new ArrayBuffer(),
          extractZip: async () => ({}),
        };
      }
      if (request === 'pptx-svg/wasm') return new ArrayBuffer();
      return originalLoad.call(this, request, parent, isMain);
    };
    try {
      return require(outfile);
    } finally {
      Module._load = originalLoad;
    }
  })();
  return modulePromise;
}

function createSaveHost() {
  return {
    app: {},
    t: (key) => key,
    statusEl: null,
    getSettings: () => ({ autosaveEnabled: false }),
    getFile: () => null,
    getLoadedFile: () => null,
    getEngine: () => null,
    getSourcePackage: () => null,
    getSourceBuffer: () => null,
    setSource: () => {},
    isCurrentPresentation: () => false,
    ensureEditable: () => true,
    getViewOnlyReason: () => '',
  };
}

test('PresentationSession publishes selection snapshots', async () => {
  const { PresentationSession } = await loadPresentationSessionModule();
  const session = new PresentationSession(createSaveHost());
  const events = [];
  const unsubscribe = session.subscribe((event) => events.push(event));

  session.selectShapes([3, 1, 3]);
  assert.deepEqual(session.selection.shapeIndexes, [3, 1]);
  assert.deepEqual(session.snapshot().selection.shapeIndexes, [3, 1]);
  assert.equal(events.at(-1).type, 'selection');

  session.clearSelection();
  assert.deepEqual(session.selection.shapeIndexes, []);
  assert.equal(events.at(-1).type, 'selection');

  unsubscribe();
  session.selectShapes([2]);
  assert.equal(events.length, 2);
});

test('PresentationSession emits command intent without dirtying a noop', async () => {
  const { PresentationSession } = await loadPresentationSessionModule();
  const session = new PresentationSession(createSaveHost());
  const events = [];
  session.subscribe((event) => events.push(event));

  session.applyCommand({ type: 'noop', reason: 'test' });

  assert.equal(session.dirty, false);
  assert.equal(events.at(-1).type, 'command');
  assert.equal(events.at(-1).command.type, 'noop');
});

test('PresentationSession commits mutations through its executor before marking dirty', async () => {
  const { PresentationSession } = await loadPresentationSessionModule();
  const commands = [];
  const previousWindow = globalThis.window;
  globalThis.window = { setTimeout: () => 1, clearTimeout: () => {} };
  const session = new PresentationSession(createSaveHost(), {
    mutationExecutor: {
      async execute(command) {
        commands.push(command);
        assert.equal(session.dirty, false);
        return 17;
      },
    },
  });

  try {
    const result = await session.applyCommand({ type: 'insert-text-box', slideIndex: 0 });

    assert.equal(result, 17);
    assert.deepEqual(commands, [{ type: 'insert-text-box', slideIndex: 0 }]);
    assert.equal(session.dirty, true);
    assert.equal(session.editVersion, 1);
  } finally {
    globalThis.window = previousWindow;
  }
});

test('PresentationSession does not dirty an overlap-aware reorder with no structural target', async () => {
  const { PresentationSession } = await loadPresentationSessionModule();
  const session = new PresentationSession(createSaveHost(), {
    mutationExecutor: { execute: async () => null },
  });

  const result = await session.applyCommand({
    type: 'reorder-shapes',
    slideIndex: 0,
    shapeIndexes: [4],
    mode: 'forward',
    intersectingOnly: true,
  });

  assert.equal(result, null);
  assert.equal(session.dirty, false);
  assert.equal(session.editVersion, 0);
});

test('PresentationSession leaves state unchanged when a mutation fails', async () => {
  const { PresentationSession } = await loadPresentationSessionModule();
  const failure = new Error('mutation failed');
  const session = new PresentationSession(createSaveHost(), {
    mutationExecutor: { execute: async () => { throw failure; } },
  });

  await assert.rejects(
    () => session.applyCommand({ type: 'insert-text-box', slideIndex: 0 }),
    failure,
  );

  assert.equal(session.dirty, false);
  assert.equal(session.editVersion, 0);
});

test('PresentationSession exposes history outcomes and snapshots', async () => {
  const { PresentationSession } = await loadPresentationSessionModule();
  const session = new PresentationSession(createSaveHost(), {
    history: { undo: () => true, redo: () => false },
  });
  const events = [];
  session.subscribe((event) => events.push(event));

  assert.equal(await session.undo(), true);
  assert.equal(await session.redo(), false);

  assert.deepEqual(
    events.map(({ type, action }) => [type, action]),
    [['history', 'undo']],
  );
  assert.ok(events[0].snapshot, 'the successful history action publishes a snapshot');
});

test('PresentationSession waits for async history completion before reporting an undo', async () => {
  const { PresentationSession } = await loadPresentationSessionModule();
  let completeUndo;
  const session = new PresentationSession(createSaveHost(), {
    history: {
      undo: () => new Promise((resolve) => { completeUndo = resolve; }),
      redo: () => false,
    },
  });
  const events = [];
  session.subscribe((event) => events.push(event));

  const undo = session.undo();
  assert.deepEqual(events, [], 'history is not announced before the restore settles');

  completeUndo(true);
  assert.equal(await undo, true);
  assert.deepEqual(events.map(({ type, action }) => [type, action]), [['history', 'undo']]);
});
