export interface EventTargetContainer {
  contains(node: Node): boolean;
}

/**
 * Returns whether the paged-editor container should replay a deletion key.
 *
 * A keydown from the hidden ProseMirror editor bubbles through the paged
 * surface. Replaying that event runs the keymap twice, which is observable as
 * two paragraphs being joined by one Backspace. Only replay unhandled keys
 * whose original target was the paged surface itself (the focus-recovery
 * path).
 */
export function shouldReplayDeletionFromContainer(
  event: Pick<KeyboardEvent, 'defaultPrevented' | 'target'>,
  editorDom: EventTargetContainer,
): boolean {
  if (event.defaultPrevented) {
    return false;
  }

  const target = event.target;
  return target === null || !editorDom.contains(target as Node);
}
