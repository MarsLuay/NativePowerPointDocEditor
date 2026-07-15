import { Page } from './layout-engine/types.js';

/**
 * Page geometry needed to translate OOXML `relativeFrom` anchors into painter
 * coordinates. All values are in CSS pixels.
 */
interface PageGeometry {
    pageWidth: number;
    pageHeight: number;
    marginLeft: number;
    marginTop: number;
    contentWidth: number;
    contentHeight: number;
}
interface AnchoredObjectPositionInput {
    width: number;
    height: number;
    position?: {
        horizontal?: {
            relativeTo?: string;
            posOffset?: number;
            align?: string;
        };
        vertical?: {
            relativeTo?: string;
            posOffset?: number;
            align?: string;
        };
    };
    cssFloat?: 'left' | 'right' | 'none';
}
interface AnchoredObjectPosition {
    x: number;
    y: number;
    side: 'left' | 'right';
}
declare function pageGeometryFromPage(page: Pick<Page, 'size' | 'margins'>): PageGeometry;
declare function resolveAnchoredObjectPosition(object: AnchoredObjectPositionInput, fragmentY: number, contentWidth: number, geometry?: PageGeometry): AnchoredObjectPosition;
/**
 * Content-area top Y (px) of an anchored object, resolving its OOXML vertical
 * anchor (`page` / `margin` / `topMargin` / `bottomMargin` / `paragraph`, with
 * `align` or `posOffset`). Exposed so the measure pipeline can reserve a
 * `topAndBottom` band at the exact Y the painter will place the object —
 * the two paths must not diverge (dual-renderer rule).
 */
declare function resolveAnchoredObjectVerticalTop(object: AnchoredObjectPositionInput, fragmentY: number, geometry?: PageGeometry): number;

export { type PageGeometry as P, resolveAnchoredObjectVerticalTop as a, pageGeometryFromPage as p, resolveAnchoredObjectPosition as r };
