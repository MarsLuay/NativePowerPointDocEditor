/**
 * Pure layout algorithm: given a list of sidebar items + their
 * anchor sources, return a list of resolved Y positions with
 * collision avoidance applied. Lifted from React's
 * components/sidebar/resolveItemPositions.ts so the Vue adapter
 * can call it directly.
 *
 * The algorithm walks the items in three priority stages:
 *   1. explicit fixedY (already in scroll-container coords)
 *   2. anchorKey lookup in anchorPositions (layout-engine coords)
 *   3. DOM-based anchor lookup via renderedDomContext.getRectsForRange
 *   4. last-known cache so cards don't pop out during transient layout
 *
 * Then it sorts by target Y and pushes overlapping cards down by
 * their height + MIN_CARD_GAP.
 * @packageDocumentation
 * @public
 */
import { SidebarItem, RenderedDomContext } from './types.js';
import 'prosemirror-state';
import 'prosemirror-view';
import 'prosemirror-model';

/**
 * Pure layout algorithm: given a list of sidebar items + their
 * anchor sources, return a list of resolved Y positions with
 * collision avoidance applied. Lifted from React's
 * components/sidebar/resolveItemPositions.ts so the Vue adapter
 * can call it directly.
 *
 * The algorithm walks the items in three priority stages:
 *   1. explicit fixedY (already in scroll-container coords)
 *   2. anchorKey lookup in anchorPositions (layout-engine coords)
 *   3. DOM-based anchor lookup via renderedDomContext.getRectsForRange
 *   4. last-known cache so cards don't pop out during transient layout
 *
 * Then it sorts by target Y and pushes overlapping cards down by
 * their height + MIN_CARD_GAP.
 * @packageDocumentation
 * @public
 */

/** Generic sidebar item with optional fixedY / estimatedHeight. */
interface ResolvableSidebarItem extends SidebarItem {
    fixedY?: number;
    estimatedHeight?: number;
}
interface ResolvedPosition<T extends ResolvableSidebarItem = ResolvableSidebarItem> {
    item: T;
    y: number;
}
declare function resolveItemPositions<T extends ResolvableSidebarItem>(items: T[], anchorPositions: Map<string, number>, renderedDomContext: RenderedDomContext | null, zoom: number, cardHeights: Map<string, number>, lastKnown: Map<string, number>): ResolvedPosition<T>[];

export { type ResolvableSidebarItem, type ResolvedPosition, resolveItemPositions };
