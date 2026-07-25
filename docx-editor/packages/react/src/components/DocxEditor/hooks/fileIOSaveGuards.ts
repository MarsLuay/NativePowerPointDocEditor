import type { EditorView } from 'prosemirror-view';

/**
 * A save can yield while React replaces the paged editor. Only mutate the
 * captured view when it is still the live, usable instance.
 */
export function isCurrentLiveEditorView(
  capturedView: EditorView | null | undefined,
  currentView: EditorView | null | undefined
): capturedView is EditorView {
  return capturedView != null && capturedView === currentView && !capturedView.isDestroyed;
}
