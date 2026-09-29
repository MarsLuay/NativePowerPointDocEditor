import { describe, expect, test } from 'bun:test';

import { shouldReplayDeletionFromContainer } from './keyboardEventRouting';

function eventTargetContainer(editorTarget: EventTarget) {
  return {
    contains: (node: Node) => node === editorTarget,
  };
}

describe('shouldReplayDeletionFromContainer', () => {
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
      shouldReplayDeletionFromContainer(
        { defaultPrevented: true, target: editorTarget },
        eventTargetContainer(editorTarget),
      ),
    ).toBe(false);
  });

  test('replays an unhandled deletion targeted at the paged surface', () => {
    const editorTarget = {} as EventTarget;
    const pagedSurfaceTarget = {} as EventTarget;

    expect(
      shouldReplayDeletionFromContainer(
        { defaultPrevented: false, target: pagedSurfaceTarget },
        eventTargetContainer(editorTarget),
      ),
    ).toBe(true);
  });
});
