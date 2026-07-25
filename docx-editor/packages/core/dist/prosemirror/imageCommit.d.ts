import { Node } from 'prosemirror-model';
import { EditorView } from 'prosemirror-view';

/**
 * Image resize / drag-move PM commits shared by the React and Vue adapters.
 *
 * The float-vs-inline fork is identical across adapters; only the DOM lookups
 * differ (which `.layout-page-content` the drop landed in; how the inline drop
 * position is hit-tested). Those stay in each adapter, which passes the
 * resolved EMU offsets (float) or drop position (inline) into the pure commits
 * here. Each commit dispatches and returns the PM position the caller should
 * re-select as a NodeSelection (or null on no-op / failure).
 */

/**
 * An image resize handle: the 4 corners ('nw'/'ne'/'se'/'sw') resize both axes
 * and keep the aspect ratio; the 4 edge midpoints ('n'/'s'/'e'/'w') resize a
 * single axis so the user can deliberately stretch the image (break aspect).
 */
type ImageResizeHandle = 'nw' | 'ne' | 'se' | 'sw' | 'n' | 's' | 'e' | 'w';
/**
 * New image dimensions for a resize drag, shared by the React and Vue overlays
 * (issue #266). Corner handles drive both axes (aspect-locked unless
 * `lockAspect` is false, e.g. Shift held); edge handles drive one axis and
 * never lock. The non-driven axis is returned unchanged. Driven axes are
 * clamped to a sane pixel range.
 */
declare function calculateResizedImageDimensions(handle: ImageResizeHandle, deltaX: number, deltaY: number, startWidth: number, startHeight: number, lockAspect: boolean): {
    width: number;
    height: number;
};
/** True when the image is floating (anchored) rather than inline. */
declare function isFloatingImage(node: Node): boolean;
/**
 * Resize commit: set the image node's `width`/`height`. Returns `pmPos` to
 * re-select, or null if the position no longer holds an image.
 */
declare function commitImageResize(view: EditorView, pmPos: number, newWidth: number, newHeight: number): number | null;
/**
 * Floating drag commit: rewrite the anchor's margin-relative `position`
 * offsets (in EMU) so the image lands at the drop point while staying
 * floating. Returns `pmPos` to re-select, or null on no-op / failure.
 */
declare function commitImageFloatMove(view: EditorView, pmPos: number, hOffsetEmu: number, vOffsetEmu: number): number | null;
/**
 * Inline drag commit: move the image node to `dropPos` via a delete + insert
 * pair. Returns the PM position to re-select, or null when the drop is a
 * no-op (same slot) or fails.
 */
declare function commitImageInlineMove(view: EditorView, pmPos: number, dropPos: number): number | null;

export { type ImageResizeHandle, calculateResizedImageDimensions, commitImageFloatMove, commitImageInlineMove, commitImageResize, isFloatingImage };
