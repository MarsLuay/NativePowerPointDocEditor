import { describe, expect, test } from 'bun:test';
import type { EditorView } from 'prosemirror-view';
import { isCurrentLiveEditorView } from './fileIOSaveGuards';

function view(isDestroyed = false): EditorView {
  return { isDestroyed } as EditorView;
}

describe('isCurrentLiveEditorView', () => {
  test('accepts the unchanged live view', () => {
    const captured = view();

    expect(isCurrentLiveEditorView(captured, captured)).toBe(true);
  });

  test('rejects a view destroyed while save awaited serialization', () => {
    const captured = view(true);

    expect(isCurrentLiveEditorView(captured, captured)).toBe(false);
  });

  test('rejects a replaced view after editor reload', () => {
    expect(isCurrentLiveEditorView(view(), view())).toBe(false);
  });
});
