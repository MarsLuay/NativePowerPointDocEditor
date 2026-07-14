import { WrapTextDirection, FlowBlock, Measure } from './layout-engine/types.js';
import { P as PageGeometry } from './anchoredObjectPosition-CS-8BfhO.js';

interface FloatingExclusionRect {
    /** Which side the object is on for simple one-sided wrapping. */
    side: 'left' | 'right';
    /** X position relative to the content area. */
    x: number;
    /** Y position relative to the content area. */
    y: number;
    width: number;
    height: number;
    distTop: number;
    distBottom: number;
    distLeft: number;
    distRight: number;
    wrapText?: WrapTextDirection;
    wrapType?: string;
}
interface FloatingImageZone {
    leftMargin: number;
    rightMargin: number;
    topY: number;
    bottomY: number;
    segments?: FloatingLineSegmentZone[];
    /**
     * Full-width vertical band (OOXML `topAndBottom` wrap): no text fits beside
     * it, so any line overlapping `[topY, bottomY]` is pushed below the band.
     */
    fullWidthBlock?: boolean;
}
interface FloatingLineSegmentZone {
    leftOffset: number;
    availableWidth: number;
}
interface FloatingLineMargins {
    leftMargin: number;
    rightMargin: number;
    segments?: FloatingLineSegmentZone[];
}
declare function rectsToFloatingZones(rects: FloatingExclusionRect[], contentWidth: number): FloatingImageZone[];
declare function getFloatingMargins(lineY: number, lineHeight: number, zones: FloatingImageZone[] | undefined, paragraphYOffset: number): FloatingLineMargins;

/**
 * Floating-aware block measurement pipeline.
 *
 * Pre-scans a block list to extract exclusion zones from anchored images,
 * floating tables, and floating text boxes; groups co-located floats so
 * their combined exclusion applies starting from the earliest anchor; then
 * walks the blocks calling the caller-supplied `measureBlock` with the
 * active zones and cumulative Y at each step.
 *
 * Adapters (React, Vue) provide their own `measureBlock` so they can
 * decide e.g. whether to cache paragraph measures. The orchestration,
 * extraction, and grouping live here so both adapters stay in lockstep.
 *
 * @packageDocumentation
 * @public
 */

/**
 * Block-measurement callback shape passed to {@link measureBlocksWithFloats}.
 * Adapters (React, Vue) supply this so they can decide platform-specific
 * concerns (e.g. paragraph-measure caching, per-section width) while
 * sharing the floating-zone orchestration. This is adapter-author API,
 * not end-consumer API.
 *
 * @public
 */
type MeasureBlockFn = (block: FlowBlock, contentWidth: number, floatingZones?: FloatingImageZone[], cumulativeY?: number) => Measure;
/**
 * Page geometry (CSS px) used to resolve page/margin-relative anchored objects
 * into content-area coordinates — currently the vertical anchor of a top-pinned
 * `topAndBottom` band. Same shape the painter uses (see `pageGeometryFromPage`),
 * so both paths resolve to identical positions.
 *
 * @public
 */
type FloatPageGeometry = PageGeometry;
/**
 * Walk `blocks` and produce one `Measure` per block. Before measuring, this
 * extracts floating exclusion zones (images / floating tables / floating
 * textboxes), groups overlapping co-located floats, and threads the active
 * zones plus cumulative Y into each `measureBlock` call.
 *
 * Pass `pageGeometry` whenever the document may contain page/margin-anchored
 * `topAndBottom` text boxes (e.g. a title banner pinned to the page top):
 * without it their reserved band falls back to flow-relative Y and the band
 * won't line up with where the painter places the box. Build it with the
 * shared `pageGeometryFromPage` helper.
 *
 * @public
 */
declare function measureBlocksWithFloats(blocks: FlowBlock[], contentWidth: number | number[], measureBlock: MeasureBlockFn, pageGeometry?: FloatPageGeometry): Measure[];

export { type FloatPageGeometry as F, type MeasureBlockFn as M, type FloatingExclusionRect as a, type FloatingImageZone as b, type FloatingLineSegmentZone as c, getFloatingMargins as g, measureBlocksWithFloats as m, rectsToFloatingZones as r };
