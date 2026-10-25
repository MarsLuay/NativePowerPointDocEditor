export interface EventTargetContainer {
  contains(node: Node): boolean;
}

export interface ClosestTarget {
  closest?: (selector: string) => unknown;
}

/** Hidden body PM host. Portaled to document.body so PM scroll restore cannot clobber the page scroller. */
export const HIDDEN_PROSEMIRROR_SELECTOR = '.paged-editor__hidden-pm';

/** Persistent hidden header/footer PM hosts. */
export const HIDDEN_HEADER_FOOTER_PM_SELECTOR = '[data-hf-r-id]';

const CARET_NAVIGATION_KEYS = new Set([
  'ArrowLeft',
  'ArrowRight',
  'ArrowUp',
  'ArrowDown',
  'Home',
  'End',
]);

/**
 * Returns whether the paged-editor container should replay a key into hidden PM.
 *
 * A keydown from the hidden ProseMirror editor bubbles through the paged
 * surface (React portal). Replaying that event runs the keymap twice, which
 * is observable as two paragraphs being joined by one Backspace. Only replay
 * unhandled keys whose original target was the paged surface itself (the
 * focus-recovery path).
 */
export function shouldReplayKeyFromContainer(
  event: Pick<KeyboardEvent, 'defaultPrevented' | 'target'>,
  editorDom: EventTargetContainer,
): boolean {
  if (event.defaultPrevented) {
    return false;
  }

  const target = event.target;
  return target === null || !editorDom.contains(target as Node);
}

/** Backspace/Delete share the caret-key replay rule. */
export const shouldReplayDeletionFromContainer = shouldReplayKeyFromContainer;

export function isCaretNavigationKey(
  key: string,
  modifiers: { ctrl?: boolean; meta?: boolean; alt?: boolean; composing?: boolean } = {},
): boolean {
  if (modifiers.ctrl || modifiers.meta || modifiers.alt || modifiers.composing) {
    return false;
  }
  return CARET_NAVIGATION_KEYS.has(key);
}

function isNodeLike(value: unknown): value is Node {
  return typeof value === 'object' && value !== null && 'nodeType' in (value as Node);
}

/**
 * Hidden PM lives in a body portal, so container.contains(pm) is false.
 * Focus moving there is still inside the editor; treating it as a blur
 * hides the caret overlay while arrow keys scroll the page scroller.
 */
export function shouldKeepPagedEditorFocused(
  relatedTarget: ClosestTarget | Node | null | undefined,
  container: EventTargetContainer | null | undefined,
): boolean {
  if (!relatedTarget) {
    return false;
  }
  if (container && isNodeLike(relatedTarget) && container.contains(relatedTarget)) {
    return true;
  }
  const closest = (relatedTarget as ClosestTarget).closest;
  if (typeof closest === 'function') {
    return Boolean(
      closest.call(relatedTarget, HIDDEN_PROSEMIRROR_SELECTOR)
      || closest.call(relatedTarget, HIDDEN_HEADER_FOOTER_PM_SELECTOR),
    );
  }
  return false;
}
