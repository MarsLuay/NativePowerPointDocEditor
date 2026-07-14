/**
 * RenderedDomContext Implementation
 *
 * Provides DOM-based position mapping for the LayoutPainter output.
 * Uses the same data-pm-start/data-pm-end attribute pattern as the
 * selection overlay in PagedEditor.
 * @packageDocumentation
 * @public
 */
import { RenderedDomContext, PositionCoordinates } from './types.mjs';
import 'prosemirror-state';
import 'prosemirror-view';
import 'prosemirror-model';

/**
 * RenderedDomContext Implementation
 *
 * Provides DOM-based position mapping for the LayoutPainter output.
 * Uses the same data-pm-start/data-pm-end attribute pattern as the
 * selection overlay in PagedEditor.
 * @packageDocumentation
 * @public
 */

/**
 * Implementation of RenderedDomContext.
 *
 * This class provides position mapping between ProseMirror document
 * positions and pixel coordinates in the rendered DOM. It uses the
 * data-pm-start and data-pm-end attributes that LayoutPainter adds
 * to span elements.
 */
declare class RenderedDomContextImpl implements RenderedDomContext {
    pagesContainer: HTMLElement;
    zoom: number;
    constructor(pagesContainer: HTMLElement, zoom?: number);
    /**
     * Get pixel coordinates for a ProseMirror position.
     * Uses the browser's text rendering via Range API for precise positioning.
     */
    getCoordinatesForPosition(pmPos: number): PositionCoordinates | null;
    /**
     * Find DOM elements that overlap with a ProseMirror position range.
     */
    findElementsForRange(from: number, to: number): Element[];
    /**
     * Get bounding rectangles for a range of text.
     * Handles line wraps by returning multiple rects.
     */
    getRectsForRange(from: number, to: number): Array<{
        x: number;
        y: number;
        width: number;
        height: number;
    }>;
    /**
     * Get the offset of the pages container from its parent viewport.
     * This is needed for positioning overlays that are rendered in the
     * viewport container rather than directly in the pages container.
     */
    getContainerOffset(): {
        x: number;
        y: number;
    };
}
/**
 * Create a RenderedDomContext for a pages container element.
 *
 * @param pagesContainer - The container element holding rendered pages
 * @param zoom - Current zoom level (default 1)
 */
declare function createRenderedDomContext(pagesContainer: HTMLElement, zoom?: number): RenderedDomContext;

export { RenderedDomContextImpl, createRenderedDomContext };
