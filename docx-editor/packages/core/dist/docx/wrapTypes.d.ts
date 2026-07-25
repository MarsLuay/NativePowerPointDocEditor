/**
 * OOXML image wrap-type taxonomy.
 *
 * `wp:inline` flows in the line. `wp:anchor` covers all positioned variants:
 *   - `square` / `tight` / `through` — text wraps around the image
 *   - `topAndBottom` — text breaks above and below
 *   - `behind` / `inFront` (`wp:wrapNone`) — image paints out of flow
 * @packageDocumentation
 * @public
 */
type WrapType = 'inline' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behind' | 'inFront';
/** True for wrap types that anchor at a position (i.e. not `inline`, not `topAndBottom`). */
declare function isFloatingWrapType(wrapType: string | undefined): boolean;
/** True for `wp:wrapNone` variants (`behind` / `inFront`) — positioned but ignore text-flow. */
declare function isWrapNone(wrapType: string | undefined): boolean;
/** True for wrap types where text flows around the image (`square` / `tight` / `through`). */
declare function wrapsAroundText(wrapType: string | undefined): boolean;

export { type WrapType, isFloatingWrapType, isWrapNone, wrapsAroundText };
