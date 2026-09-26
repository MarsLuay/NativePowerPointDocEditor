import assert from 'node:assert/strict';
import test from 'node:test';
import { loadTextUtilsModule } from './helpers/load-plugin-modules.mjs';

test('Mac delete key with code Backspace is backward even if key says Delete', async () => {
  const { isBackwardDeleteKey } = await loadTextUtilsModule();
  assert.equal(isBackwardDeleteKey({ key: 'Delete', code: 'Backspace' }), true);
});

test('forward Delete keeps forward semantics', async () => {
  const { isBackwardDeleteKey } = await loadTextUtilsModule();
  assert.equal(isBackwardDeleteKey({ key: 'Delete', code: 'Delete' }), false);
});

test('normal Backspace is backward', async () => {
  const { isBackwardDeleteKey } = await loadTextUtilsModule();
  assert.equal(isBackwardDeleteKey({ key: 'Backspace', code: 'Backspace' }), true);
});

test('i|n + Mac delete (misreported key) removes i, not n', async () => {
  const { isBackwardDeleteKey } = await loadTextUtilsModule();
  const text = 'in';
  const selectionStart = 1;
  const event = { key: 'Delete', code: 'Backspace' };
  const backward = isBackwardDeleteKey(event);
  const deleteStart = backward ? selectionStart - 1 : selectionStart;
  const deleteEnd = backward ? selectionStart : selectionStart + 1;
  assert.equal(text.slice(deleteStart, deleteEnd), 'i');
});
