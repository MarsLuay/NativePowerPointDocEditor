import assert from "node:assert/strict";
import { test } from "node:test";
import {
  loadDocxSessionModule,
  loadFakeDocxEditorAdapterModule,
} from "./helpers/load-plugin-modules.mjs";

function createSessionOptions(persist) {
  return {
    adapter: {
      serialize: async () => new ArrayBuffer(8),
      prepareForWrite: async (buffer) => buffer,
      validate: async (buffer) => buffer,
      persist,
    },
    getContext: () => ({ file: "example.docx" }),
    autosave: {
      enabled: () => false,
      delayMs: () => 0,
      source: "autosave",
    },
  };
}

test("DocxSession publishes dirty and successful save state", async () => {
  const { DocxSession } = await loadDocxSessionModule();
  const persisted = [];
  const session = new DocxSession(createSessionOptions(async (buffer) => persisted.push(buffer)));
  const snapshots = [];
  const unsubscribe = session.subscribe((snapshot) => snapshots.push(snapshot));

  session.markDirty();
  assert.equal(session.dirty, true);
  assert.equal(session.editVersion, 1);
  assert.equal(await session.save(), true);
  unsubscribe();

  assert.equal(persisted.length, 1);
  assert.deepEqual(snapshots.map(({ saveState }) => saveState), ["dirty", "saving", "clean"]);
  assert.deepEqual(session.snapshot(), {
    dirty: false,
    editVersion: 1,
    saveState: "clean",
    saveError: undefined,
  });
});

test("DocxSession exposes failed save state to subscribers", async () => {
  const { DocxSession } = await loadDocxSessionModule();
  const failure = new Error("write failed");
  const session = new DocxSession(createSessionOptions(async () => { throw failure; }));
  const snapshots = [];
  session.subscribe((snapshot) => snapshots.push(snapshot));

  session.markDirty();
  assert.equal(await session.save(), false);

  assert.equal(session.saveState, "failed");
  assert.equal(session.dirty, true);
  assert.equal(snapshots.at(-1).saveError, failure);
});

test("FakeDocxEditorAdapter records modes and publishes chrome updates", async () => {
  const { FakeDocxEditorAdapter } = await loadFakeDocxEditorAdapterModule();
  const adapter = new FakeDocxEditorAdapter();
  let notifications = 0;
  const subscription = adapter.observeChrome(() => { notifications += 1; });

  adapter.setMode("viewing");
  adapter.emitChrome();
  subscription.dispose();
  adapter.emitChrome();

  assert.deepEqual(adapter.modes, ["viewing"]);
  assert.equal(notifications, 1);
  assert.equal(await adapter.serialize(), null);
});

test("DocxSession preserves a newer edit made while saving", async () => {
  const { DocxSession } = await loadDocxSessionModule();
  let releasePersist;
  const persistGate = new Promise((resolve) => {
    releasePersist = resolve;
  });
  let enteredPersist = false;
  const session = new DocxSession(createSessionOptions(async () => {
    enteredPersist = true;
    await persistGate;
  }));

  session.markDirty();
  const save = session.save();
  for (let i = 0; i < 20 && !enteredPersist; i++) {
    await Promise.resolve();
  }
  assert.equal(enteredPersist, true);
  session.markDirty();
  releasePersist();

  assert.equal(await save, true);
  assert.equal(session.dirty, true);
  assert.equal(session.editVersion, 2);
  assert.equal(session.saveState, "dirty");
});

test("DocxSession persists the live dirty version when edits land during serialize", async () => {
  const { DocxSession } = await loadDocxSessionModule();
  const persisted = [];
  let releaseFirstSerialize;
  const firstSerialize = new Promise((resolve) => {
    releaseFirstSerialize = resolve;
  });
  let serializations = 0;
  const session = new DocxSession({
    adapter: {
      serialize: async () => {
        serializations += 1;
        if (serializations === 1) await firstSerialize;
        return new ArrayBuffer(serializations);
      },
      prepareForWrite: async (buffer) => buffer,
      validate: async (buffer) => buffer,
      persist: async (buffer) => {
        persisted.push(buffer.byteLength);
      },
    },
    getContext: () => ({ file: "example.docx" }),
    autosave: {
      enabled: () => false,
      delayMs: () => 0,
      source: "autosave",
    },
  });

  session.markDirty();
  const save = session.save();
  await Promise.resolve();
  session.markDirty();
  releaseFirstSerialize();

  assert.equal(await save, true);
  assert.deepEqual(persisted, [2]);
  assert.equal(serializations, 2);
  assert.equal(session.dirty, false);
  assert.equal(session.editVersion, 2);
  assert.equal(session.saveState, "clean");
});

test("DocxSession waits for an in-flight save before an external reload", async () => {
  const { DocxSession } = await loadDocxSessionModule();
  let releasePersist;
  const persistStarted = new Promise((resolve) => {
    releasePersist = resolve;
  });
  const session = new DocxSession(createSessionOptions(async () => persistStarted));

  session.markDirty();
  const save = session.save();
  const idle = session.waitForIdle();
  let idleResolved = false;
  void idle.then(() => { idleResolved = true; }).catch(() => {});
  await Promise.resolve();
  assert.equal(idleResolved, false);

  releasePersist();
  assert.equal(await save, true);
  await idle;
  assert.equal(idleResolved, true);
});

test("DocxSession flush-before-close leaves clean after in-flight autosave", async () => {
  const { DocxSession } = await loadDocxSessionModule();
  let releasePersist;
  const persistGate = new Promise((resolve) => {
    releasePersist = resolve;
  });
  let enteredPersist = false;
  const session = new DocxSession(createSessionOptions(async () => {
    enteredPersist = true;
    await persistGate;
  }));

  session.markDirty();
  const inFlight = session.save("autosave");
  for (let i = 0; i < 20 && !enteredPersist; i++) {
    await Promise.resolve();
  }
  assert.equal(enteredPersist, true);
  assert.equal(session.dirty, true);

  // Close-path flush: await current/queued save, then idle — no prompt needed.
  const flush = (async () => {
    let saved = true;
    if (session.dirty) {
      saved = await session.save("autosave");
    }
    await session.waitForIdle();
    return saved && !session.dirty;
  })();

  releasePersist();
  assert.equal(await inFlight, true);
  assert.equal(await flush, true);
  assert.equal(session.dirty, false);
  assert.equal(session.saveState, "clean");
});

test("FakeDocxEditorAdapter delegates editor operations through its options", async () => {
  const { FakeDocxEditorAdapter } = await loadFakeDocxEditorAdapterModule();
  const match = { from: 1, to: 3 };
  const calls = [];
  const adapter = new FakeDocxEditorAdapter({
    getSelectedText: () => "selected",
    find: (query, options) => {
      calls.push(["find", query, options]);
      return [match];
    },
    select: (value) => value === match,
    replace: (value, replacement) => value === match && replacement === "next",
    replaceAll: (values, replacement) => values[0] === match && replacement === "all",
  });

  const options = { matchCase: true, wholeWord: false };
  assert.equal(adapter.getSelectedText(), "selected");
  assert.deepEqual(adapter.find("find", options), [match]);
  assert.equal(adapter.select(match), true);
  assert.equal(adapter.replace(match, "next"), true);
  assert.equal(adapter.replaceAll([match], "all"), true);
  assert.deepEqual(calls, [["find", "find", options]]);
});
