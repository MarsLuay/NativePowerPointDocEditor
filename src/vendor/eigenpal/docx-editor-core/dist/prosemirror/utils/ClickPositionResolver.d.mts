/**
 * Click Position Resolver
 *
 * Provides fast, indexed lookups for click-to-position mapping.
 * Caches page, fragment, and run positions for O(log n) lookups
 * instead of O(n) DOM traversal.
 * @packageDocumentation
 * @public
 */
/**
 * Cached page information for fast Y-coordinate lookup.
 */
interface PageInfo {
    element: HTMLElement;
    index: number;
    top: number;
    bottom: number;
}
/**
 * Cached fragment information.
 */
interface FragmentInfo {
    element: HTMLElement;
    blockId: string;
    top: number;
    bottom: number;
    left: number;
    right: number;
}
/**
 * Cached run (text span) information.
 */
interface RunInfo {
    element: HTMLElement;
    pmStart: number;
    pmEnd: number;
    left: number;
    right: number;
    isTab: boolean;
}
/**
 * Result of a position lookup.
 */
interface PositionLookupResult {
    pageIndex: number;
    pmPosition: number;
    element: HTMLElement;
}
/**
 * ClickPositionResolver provides fast click-to-position mapping
 * by caching DOM element positions and using binary search.
 */
declare class ClickPositionResolver {
    #private;
    /**
     * Rebuild the entire index from the container.
     * Call this after layout changes.
     */
    rebuild(container: HTMLElement): void;
    /**
     * Mark the index as dirty (needs rebuild).
     */
    invalidate(): void;
    /**
     * Check if index needs rebuilding.
     */
    isDirty(): boolean;
    /**
     * Get the page at a Y coordinate using binary search.
     */
    getPageAtY(clientY: number): PageInfo | null;
    /**
     * Get the fragment at a point within a page.
     */
    getFragmentAtPoint(pageIndex: number, _clientX: number, clientY: number): FragmentInfo | null;
    /**
     * Get the run at an X coordinate within a fragment.
     */
    getRunAtX(blockId: string, clientX: number): RunInfo | null;
    /**
     * Get the exact PM position within a run using binary search.
     */
    getPositionInRun(run: RunInfo, clientX: number): number;
    /**
     * Get PM position from client coordinates.
     * Main entry point for click-to-position mapping.
     */
    getPositionAtPoint(clientX: number, clientY: number): PositionLookupResult | null;
    /**
     * Get the element containing a PM position.
     * Useful for caret positioning.
     */
    getElementAtPosition(pmPos: number): HTMLElement | null;
    /**
     * Get all pages info (for debugging).
     */
    getPages(): ReadonlyArray<PageInfo>;
    /**
     * Get debug info.
     */
    getDebugInfo(): {
        pageCount: number;
        fragmentCount: number;
        runCount: number;
        dirty: boolean;
    };
}

export { ClickPositionResolver, type PositionLookupResult };
