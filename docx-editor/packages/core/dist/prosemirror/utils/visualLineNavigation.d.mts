/**
 * Visual line navigation helpers — implements Word/Google-Docs-style
 * ArrowUp / ArrowDown with sticky X across visual lines (not just
 * paragraphs). Lifted from packages/react/src/paged-editor/
 * useVisualLineNavigation.ts so both adapters share the algorithm.
 *
 * Frontend-agnostic: takes a `getContainer: () => HTMLElement | null`
 * callback and a mutable sticky-state object, returns the same
 * function quartet React's hook returns.
 *
 * @remarks
 * Tagged `@internal` post-1.0 cut. Both adapters re-export this through
 * their own composables (`useVisualLineNavigation`); consumers should
 * prefer those. The subpath stays in `package.json` `exports` for
 * back-compat; expect it to move behind a public surface in a future
 * major.
 *
 * @packageDocumentation
 * @internal
 */
import { EditorView } from 'prosemirror-view';

/** @internal */
interface VisualLineState {
    stickyX: number | null;
    lastVisualLineIndex: number;
}
/** @internal */
declare function createVisualLineState(): VisualLineState;
/** @internal */
declare function getCaretClientX(container: HTMLElement, pmPos: number): number | null;
/** @internal */
declare function findLineElementAtPosition(container: HTMLElement, pmPos: number): HTMLElement | null;
/** @internal */
declare function findPositionOnLineAtClientX(lineEl: HTMLElement, clientX: number): number | null;
/**
 * Handle PM ArrowUp / ArrowDown with visual-line awareness + sticky
 * X. Returns true if the event was handled and PM should not run
 * its default behaviour. Mutates `state` so consecutive presses
 * keep the same sticky X.
 */
/** @internal */
declare function handleVisualLineKeyDown(state: VisualLineState, view: EditorView, event: KeyboardEvent, container: HTMLElement | null): boolean;

export { type VisualLineState, createVisualLineState, findLineElementAtPosition, findPositionOnLineAtClientX, getCaretClientX, handleVisualLineKeyDown };
