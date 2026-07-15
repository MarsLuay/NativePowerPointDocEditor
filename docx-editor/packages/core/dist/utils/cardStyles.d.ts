/**
 * Sidebar card chrome — shared between React and Vue. Numeric
 * pixel values rather than `'8px'` strings so both adapters'
 * CSSProperties shapes accept them. Lifted from
 * `packages/react/src/components/sidebar/cardStyles.ts` and the
 * Vue mirror so there's one canonical table.
 * @packageDocumentation
 * @public
 */
/**
 * Tiny shared CSSProperties shape so framework-agnostic style
 * objects in core don't need to import from React or Vue. Both
 * adapters' `CSSProperties` types are structurally compatible
 * with this looser shape (string | number values, camelCased keys).
 */
type CSSProperties = Record<string, string | number | undefined>;

/**
 * Sidebar card chrome — shared between React and Vue. Numeric
 * pixel values rather than `'8px'` strings so both adapters'
 * CSSProperties shapes accept them. Lifted from
 * `packages/react/src/components/sidebar/cardStyles.ts` and the
 * Vue mirror so there's one canonical table.
 * @packageDocumentation
 * @public
 */

declare const CARD_STYLE_COLLAPSED: CSSProperties;
declare const CARD_STYLE_EXPANDED: CSSProperties;

export { CARD_STYLE_COLLAPSED, CARD_STYLE_EXPANDED };
