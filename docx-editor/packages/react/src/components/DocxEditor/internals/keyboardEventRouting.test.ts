import { describe, expect, test } from 'bun:test';

import {
  isCaretNavigationKey,
  shouldKeepPagedEditorFocused,
  shouldReplayDeletionFromContainer,
  shouldReplayKeyFromContainer,
} from './keyboardEventRouting';

function eventTargetContainer(editorTarget: EventTarget) {
  return {
    contains: (node: Node) => node === editorTarget,
  };
}

describe('shouldReplayKeyFromContainer', () => {
  test('does not replay a Backspace that bubbled from the hidden editor', () => {
    const editorTarget = {} as EventTarget;
    let replayCount = 0;

    if (
      shouldReplayDeletionFromContainer(
        { defaultPrevented: false, target: editorTarget },
        eventTargetContainer(editorTarget),
      )
    ) {
      replayCount += 1;
    }

    // With two consecutive empty paragraphs, replaying the PM keymap would
    // join both paragraphs for one user Backspace. The bubbled event must run
    // through the keymap only once.
    expect(replayCount).toBe(0);
  });

  test('does not replay an event already handled by ProseMirror', () => {
    const editorTarget = {} as EventTarget;

    expect(
      shouldReplayKeyFromContainer(
        { defaultPrevented: true, target: editorTarget },
        eventTargetContainer(editorTarget),
      ),
    ).toBe(false);
  });

  test('replays an unhandled key targeted at the paged surface', () => {
    const editorTarget = {} as EventTarget;
    const pagedSurfaceTarget = {} as EventTarget;

    expect(
      shouldReplayKeyFromContainer(
        { defaultPrevented: false, target: pagedSurfaceTarget },
        eventTargetContainer(editorTarget),
      ),
    ).toBe(true);
  });
});

describe('isCaretNavigationKey', () => {
  test('matches unchorded caret keys', () => {
    expect(isCaretNavigationKey('ArrowLeft')).toBe(true);
    expect(isCaretNavigationKey('ArrowDown')).toBe(true);
    expect(isCaretNavigationKey('Home')).toBe(true);
    expect(isCaretNavigationKey('a')).toBe(false);
    expect(isCaretNavigationKey('PageDown')).toBe(false);
  });

  test('ignores modified or composing caret keys', () => {
    expect(isCaretNavigationKey('ArrowDown', { ctrl: true })).toBe(false);
    expect(isCaretNavigationKey('Home', { meta: true })).toBe(false);
    expect(isCaretNavigationKey('ArrowLeft', { composing: true })).toBe(false);
  });
});

describe('shouldKeepPagedEditorFocused', () => {
  test('keeps focus when relatedTarget is inside the paged container', () => {
    const inside = { nodeType: 1 } as Node;
    expect(shouldKeepPagedEditorFocused(inside, eventTargetContainer(inside as unknown as EventTarget))).toBe(
      true,
    );
  });

  test('keeps focus when relatedTarget is the portaled hidden PM', () => {
    const hiddenPm = {
      closest: (selector: string) => (selector === '.paged-editor__hidden-pm' ? {} : null),
    };
    expect(shouldKeepPagedEditorFocused(hiddenPm, { contains: () => false })).toBe(true);
  });

  test('keeps focus when relatedTarget is a hidden header/footer PM', () => {
    const hfPm = {
      closest: (selector: string) => (selector === '[data-hf-r-id]' ? {} : null),
    };
    expect(shouldKeepPagedEditorFocused(hfPm, { contains: () => false })).toBe(true);
  });

  test('drops focus when relatedTarget is outside the editor', () => {
    const outside = {
      closest: () => null,
    };
    expect(shouldKeepPagedEditorFocused(outside, { contains: () => false })).toBe(false);
    expect(shouldKeepPagedEditorFocused(null, { contains: () => false })).toBe(false);
  });
});
