import { Document } from './types/document.mjs';
import { s as SdtType, t as SdtProperties, u as SdtDataBinding, D as DocumentBody, v as BlockSdt, w as InlineSdt, B as BlockContent } from './content-BZ9rYecc.mjs';

/**
 * Content-control (SDT) addressing for the document model.
 *
 * Content controls (`w:sdt`) are the natural anchor for template logic and
 * agent edits: they survive the round trip (see the parser + serializer) and
 * carry a stable `tag`/`alias`/`id`. This module discovers and edits them
 * without a DOM or an editor instance, so server-side pipelines and AI agents
 * can find an anchor by tag and act on it.
 *
 * Both **block-level** (`w:sdt` wrapping paragraphs/tables) and **inline**
 * (`w:sdt` inside a paragraph) controls are addressed, including inline
 * controls inside table cells (and nested tables). With
 * `{ includeHeadersFooters: true }` the walk also covers header/footer parts.
 *
 * Not surfaced (model limitations): a block SDT placed directly inside a table
 * cell (`TableCell.content` cannot hold one), an inline SDT inside a hyperlink
 * (`Hyperlink.children` excludes it), and controls inside tracked-change
 * wrappers.
 */

/** Filter for {@link findContentControls}. All provided fields must match (AND). */
interface ContentControlFilter {
    /** Developer identifier (`w:tag`), exact match. */
    tag?: string;
    /** Friendly name (`w:alias`), exact match. */
    alias?: string;
    /** Numeric id (`w:id`), exact match. */
    id?: number;
    /** Control type projection (`richText`, `dropDownList`, …). */
    type?: SdtType;
}
/**
 * Where a control lives. `body` = the main document story; `header`/`footer`
 * = a page-furniture part addressed by its relationship id (the key into
 * `package.headers`/`package.footers`).
 */
type ContentControlLocation = {
    part: 'body';
} | {
    part: 'header' | 'footer';
    rId: string;
};
/** A discovered content control plus enough context to address and edit it. */
interface ContentControlInfo {
    /** Developer identifier (`w:tag`). */
    tag?: string;
    /** Friendly name (`w:alias`). */
    alias?: string;
    /** Numeric id (`w:id`). */
    id?: number;
    /** Control type projection. */
    sdtType: SdtType;
    /** Lock setting, if any. A locked control should refuse content edits. */
    lock?: SdtProperties['lock'];
    /** Dropdown/combobox list items, if modeled. */
    listItems?: {
        displayText: string;
        value: string;
    }[];
    /** Placeholder docPart reference, if any. */
    placeholder?: string;
    /** Whether the control is currently showing placeholder text (`w:showingPlcHdr`). */
    showingPlaceholder?: boolean;
    /** Checkbox state, for checkbox controls. */
    checked?: boolean;
    /** Date format string, for date controls. */
    dateFormat?: string;
    /** XML data binding (`w:dataBinding`), if the control is bound. */
    dataBinding?: SdtDataBinding;
    /** Plain text of the control's content (paragraphs/tables/nested controls flattened). */
    text: string;
    /**
     * Block-index path to this control: top-level `[i]`, a control nested in the
     * i-th block's content `[i, j]`, and so on. For inline / cell controls it is
     * the block indices of the nearest enclosing blocks.
     */
    path: number[];
    /** Nesting depth = number of enclosing content controls (0 = not inside another control). */
    depth: number;
    /** Block-level (`w:sdt` at block level) or inline (`w:sdt` inside a paragraph). */
    kind: 'block' | 'inline';
    /** Where the control lives (body vs a header/footer part). */
    location: ContentControlLocation;
}
/** Plain text of a control's content, descending into tables and nested SDTs. */
declare function getContentControlText(control: BlockSdt | InlineSdt): string;
/** Options for {@link findContentControls}. */
interface FindContentControlsOptions {
    /**
     * When `true`, also search header/footer parts — but only when a full
     * {@link Document} is passed (a bare `DocumentBody` carries no parts, so this
     * then searches the body only and never throws). Defaults to `false` (the main
     * document story only).
     */
    includeHeadersFooters?: boolean;
}
/**
 * Find every content control in the document — block-level AND inline —
 * optionally filtered by tag/alias/id/type. Results are in strict document
 * order; nested controls follow their parent.
 *
 * The walk descends body blocks, block SDTs, tables (row-major, including
 * nested tables) into cell content, and paragraph inline content (inline
 * SDTs, recursing into nested inline SDTs). With `{ includeHeadersFooters: true }`
 * and a full {@link Document}, header then footer parts are searched after the
 * body, each sorted by relationship id for deterministic order.
 *
 * Not surfaced (model limitations, documented): a block SDT placed directly
 * inside a table cell (`TableCell.content` is `(Paragraph | Table)[]`), an
 * inline SDT inside a hyperlink (`Hyperlink.children` excludes it), and
 * controls buried inside tracked-change wrappers.
 */
declare function findContentControls(input: Document | DocumentBody, filter?: ContentControlFilter, options?: FindContentControlsOptions): ContentControlInfo[];
/** Convenience: the first control matching `filter`, or `undefined`. */
declare function findContentControl(input: Document | DocumentBody, filter: ContentControlFilter, options?: FindContentControlsOptions): ContentControlInfo | undefined;
/** No control matched the filter. */
declare class ContentControlNotFoundError extends Error {
    constructor(filter: ContentControlFilter);
}
/** The matched control's lock forbids the attempted edit (pass `force` to override). */
declare class ContentControlLockedError extends Error {
    constructor(lock: SdtProperties['lock'], op: 'edit' | 'remove');
}
/**
 * The control's type doesn't support free text/block replacement (e.g. a
 * dropdown, date, checkbox, or picture control), so writing arbitrary content
 * would desync the type marker from its value. Use a type-specific setter, or
 * pass `{ force: true }` to override.
 */
declare class ContentControlTypeError extends Error {
    constructor(sdtType: SdtType);
}
/**
 * The control is bound to a Custom XML data store (`w:dataBinding`). Writing its
 * content won't stick — Word re-renders the control from the bound XML node — so
 * the write is refused. Update the data store instead, or pass `{ force: true }`.
 */
declare class ContentControlBoundError extends Error {
    constructor();
}
/**
 * A `BlockContent[]` replacement was targeted at an inline control, which can
 * only hold inline content (runs, hyperlinks, fields, nested inline SDTs, math).
 * Splicing a paragraph into a paragraph would be invalid OOXML — pass a string,
 * or a single paragraph whose content is entirely inline.
 */
declare class ContentControlKindError extends Error {
    constructor(detail: string);
}
/**
 * Replace the content of the first control matching `filter` — block-level OR
 * inline (including inside table cells and, with `includeHeadersFooters: true`,
 * headers and footers). `replacement` may be a string or block content.
 *
 * - For a **block** control the string is split into paragraphs on newlines
 *   (a `plainText` control collapses to one paragraph); block content is used
 *   as-is (cloned).
 * - For an **inline** control the string becomes a single run that inherits the
 *   placeholder's formatting (so the value matches the field it replaces). A
 *   richText control turns `\n` into a line break; a plainText control never
 *   gets one. Passing `BlockContent[]` to an inline control throws
 *   {@link ContentControlKindError} unless it is a single all-inline paragraph.
 *
 * Pass `{ all: true }` to fill **every** control matching `filter` (one logical
 * value that recurs under a shared tag — e.g. a name in the body, a running
 * header, and several table cells) instead of just the first.
 *
 * The control's properties, tag/alias, and lossless raw `w:sdtPr` are preserved.
 * When the control was showing its placeholder (`w:showingPlcHdr`), that flag is
 * cleared so Word doesn't render the new content as placeholder text.
 *
 * Throws {@link ContentControlNotFoundError} if nothing matches,
 * {@link ContentControlLockedError} if the lock forbids editing,
 * {@link ContentControlTypeError} for a typed (dropdown/date/…) control, and
 * {@link ContentControlBoundError} for a data-bound control. Pass
 * `{ force: true }` to override the guards.
 */
declare function setContentControlContent(doc: Document, filter: ContentControlFilter, replacement: string | BlockContent[], options?: {
    force?: boolean;
    includeHeadersFooters?: boolean;
    all?: boolean;
}): Document;
/**
 * Remove the first control matching `filter` — block-level OR inline (incl.
 * inside table cells and, with `includeHeadersFooters: true`, headers/footers).
 * With `keepContent: true` the control's content is unwrapped in place (the box goes
 * away, the content stays) — block content lifts to its block siblings, inline
 * content stays inline in the enclosing paragraph. Otherwise the control and
 * its content are deleted.
 *
 * Pass `{ all: true }` to remove **every** control matching `filter` (e.g. to
 * flatten a finished template by unwrapping all controls) instead of the first.
 *
 * Unwrapping a repeating-section (item) is refused unless `force`, since lifting
 * its blocks out would orphan the (w15) repeating structure.
 *
 * Throws {@link ContentControlNotFoundError} / {@link ContentControlLockedError}
 * as {@link setContentControlContent} does.
 */
declare function removeContentControl(doc: Document, filter: ContentControlFilter, options?: {
    force?: boolean;
    keepContent?: boolean;
    includeHeadersFooters?: boolean;
    all?: boolean;
}): Document;

/**
 * Typed value setters for block-level content controls — set a dropdown
 * selection, toggle a checkbox, or set a date. These produce both the visible
 * content (the run text Word shows) and the structured state inside the
 * captured raw `w:sdtPr` (dropdown `w:lastValue`, `w14:checked`, `w:date`'s
 * `w:fullDate`), patched in place so the rest of the control round-trips
 * verbatim. Use these instead of {@link setContentControlContent} for typed
 * controls, which that function refuses by design.
 *
 * Raw `w:sdtPr` is patched with targeted string edits (not a full re-serialize)
 * to preserve the `CT_SdtPr` element order and any unmodeled properties — the
 * same capture-and-replay contract used everywhere else for SDTs.
 */

/** A typed value to apply to a content control. */
type ContentControlValue = {
    kind: 'dropdown';
    value: string;
} | {
    kind: 'checkbox';
    checked: boolean;
} | {
    kind: 'date';
    date: string;
};
/** The control doesn't support the requested value kind, or the value is invalid. */
declare class ContentControlValueError extends Error {
    constructor(message: string);
}
/** Format an ISO date (yyyy-mm-dd) with a subset of OOXML date tokens. */
declare function formatSdtDate(iso: string, pattern?: string): string;
/**
 * Set a typed value (dropdown selection / checkbox / date) on the first control
 * matching `filter` — **block-level OR inline** (inline includes controls inside
 * table cells, and with `includeHeadersFooters: true`, headers/footers) — returning a new
 * {@link Document}. Updates both the visible content and the structured raw
 * state (dropdown `w:lastValue`, `w14:checked`, `w:date/@w:fullDate`), so the
 * result round-trips and Word shows the new value.
 *
 * Pass `{ all: true }` to set the value on **every** control matching `filter`
 * (a value shared across duplicated controls) instead of just the first.
 *
 * Throws `ContentControlNotFoundError` if nothing matches,
 * {@link ContentControlLockedError} if content-locked,
 * {@link ContentControlBoundError} if data-bound (the store would override the
 * write), and {@link ContentControlValueError} if the value doesn't fit the
 * control type. The lock/bound guards are overridable with `{ force: true }`.
 */
declare function setContentControlValue(doc: Document, filter: ContentControlFilter, value: ContentControlValue, options?: {
    force?: boolean;
    includeHeadersFooters?: boolean;
    all?: boolean;
}): Document;

export { ContentControlBoundError as C, type FindContentControlsOptions as F, type ContentControlFilter as a, type ContentControlInfo as b, ContentControlKindError as c, type ContentControlLocation as d, ContentControlLockedError as e, ContentControlNotFoundError as f, ContentControlTypeError as g, type ContentControlValue as h, ContentControlValueError as i, findContentControl as j, findContentControls as k, formatSdtDate as l, getContentControlText as m, setContentControlValue as n, removeContentControl as r, setContentControlContent as s };
