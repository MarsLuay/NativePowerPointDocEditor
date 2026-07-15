/**
 * Content-control (SDT) addressing — public surface without DocumentAgent.
 * Implementation lives under `../agent/` for historical reasons; this barrel
 * is the import path for editor adapters.
 * @packageDocumentation
 * @public
 */
import { b as ContentControlInfo, a as ContentControlFilter } from '../contentControlValues-CxwmLrCX.js';
export { C as ContentControlBoundError, c as ContentControlKindError, d as ContentControlLocation, e as ContentControlLockedError, f as ContentControlNotFoundError, g as ContentControlTypeError, h as ContentControlValue, i as ContentControlValueError, F as FindContentControlsOptions, j as findContentControl, k as findContentControls, l as formatSdtDate, m as getContentControlText, r as removeContentControl, s as setContentControlContent, n as setContentControlValue } from '../contentControlValues-CxwmLrCX.js';
import { Document } from '../types/document.js';
import { s as SdtType, t as SdtProperties } from '../content-B8ScSBzC.js';
import '../colors-C3vA7HUU.js';
import '../formatting-JhqWT_XM.js';
import '../lists-Bn29SzeS.js';
import '../watermark-D90356ZM.js';
import '../styles-2J4U-Lgk.js';
import '../docx/wrapTypes.js';

/**
 * Create a new inline content control (`w:sdt`) by wrapping a text span.
 *
 * Complements the discovery/edit functions in {@link ./contentControls}: where
 * those find and mutate existing controls, this wraps an exact run of text
 * inside a paragraph in a new control with a synthesized, Word-correct
 * `w:sdtPr`. Pure — the input {@link Document} is not mutated.
 */

/** A create request failed: the target couldn't be resolved, or the wrap is invalid. */
declare class ContentControlCreateError extends Error {
    constructor(message: string);
}
/**
 * Where to create a control: an exact text span inside a paragraph. The
 * paragraph is located by Word `w14:paraId`, and the chosen `occurrence` of
 * `text` is wrapped in an inline control — including inside a table cell, where
 * block-level controls aren't allowed.
 */
interface CreateContentControlTarget {
    /** Word `w14:paraId` of the paragraph containing the text. */
    paraId: string;
    /** Exact substring to wrap. */
    text: string;
    /** Which occurrence of `text` to wrap when it repeats (1-based; default 1). */
    occurrence?: number;
}
/** Modeled properties for a control created by {@link createContentControl}. */
interface NewContentControlProps {
    /** Control type (default `richText`). */
    sdtType?: SdtType;
    /** Developer identifier (`w:tag`). */
    tag?: string;
    /** Friendly name (`w:alias`). */
    alias?: string;
    /** Numeric id (`w:id`). Default: auto-assigned, unique across the document. */
    id?: number;
    /** Lock setting (`w:lock`). */
    lock?: SdtProperties['lock'];
    /** Dropdown/combobox list items. */
    listItems?: {
        displayText: string;
        value: string;
    }[];
    /** Date display format (`w:date/w:dateFormat`), for `date` controls. */
    dateFormat?: string;
    /** Initial checkbox state, for `checkbox` controls. */
    checked?: boolean;
    /** Whether the control starts in placeholder state (`w:showingPlcHdr`). */
    showingPlaceholder?: boolean;
}
/**
 * Wrap an exact text span inside a paragraph in a new inline content control
 * (`w:sdt`), returning a new {@link Document} and the created control's
 * {@link ContentControlInfo}. Pure — the input is not mutated. This is the form
 * needed inside table cells and mid-sentence, where block controls aren't
 * allowed: runs are split at the span boundaries (formatting preserved) and
 * interior fields/tabs/breaks are kept wholesale.
 *
 * The control's `w:sdtPr` is synthesized from `props`, and its `w:id` is
 * auto-assigned (unique across the document) when `props.id` is omitted, so the
 * control round-trips and `findContentControl(doc, { tag })` resolves it after a
 * save/reload.
 *
 * **Body only:** the search covers body paragraphs and block/table content —
 * paragraphs inside headers or footers are not reachable. Passing a `paraId`
 * from a header/footer part throws {@link ContentControlCreateError} with a
 * "No paragraph found" message.
 *
 * @throws {@link ContentControlCreateError} when the paragraph or text isn't
 * found, the span overlaps an existing control or crosses a non-run boundary,
 * the `sdtType` can't be synthesized, or a supplied `id` already exists.
 */
declare function createContentControl(doc: Document, target: CreateContentControlTarget, props?: NewContentControlProps): {
    doc: Document;
    control: ContentControlInfo;
};

/**
 * Repeating-section (w15:repeatingSection) support — add and remove repeated
 * items, the way Word's "+" affordance does. A repeating section is a block
 * content control whose `w:sdtPr` carries `<w15:repeatingSection>`; its direct
 * children are item controls each carrying `<w15:repeatingSectionItem>`.
 *
 * Adding clones an existing item (with a fresh, unique `w:id`) and inserts it
 * after; removing drops one item but keeps at least one. The w15 elements ride
 * in the captured raw `w:sdtPr` (they're unmodeled), so we detect and patch the
 * raw string rather than re-serializing.
 */

/** The control's raw `w:sdtPr` declares it a repeating section (the container). */
declare function isRepeatingSection(props: SdtProperties): boolean;
/** The control's raw `w:sdtPr` declares it a repeating-section item. */
declare function isRepeatingSectionItem(props: SdtProperties): boolean;
/** Raised when an operation targets something that isn't a repeating section/item. */
declare class RepeatingSectionError extends Error {
    constructor(message: string);
}
/**
 * Add a new repeating-section item, cloned from an existing one and inserted
 * after it. `afterIndex` is the item ordinal to clone/insert after (default:
 * the last item). Returns a new {@link Document}.
 */
declare function addRepeatingSectionItem(doc: Document, filter: ContentControlFilter, options?: {
    afterIndex?: number;
}): Document;
/**
 * Remove the repeating-section item at ordinal `index`. Keeps at least one item
 * (Word does not allow removing the last). Returns a new {@link Document}.
 */
declare function removeRepeatingSectionItem(doc: Document, filter: ContentControlFilter, index: number): Document;

export { ContentControlCreateError, ContentControlFilter, ContentControlInfo, type CreateContentControlTarget, type NewContentControlProps, RepeatingSectionError, addRepeatingSectionItem, createContentControl, isRepeatingSection, isRepeatingSectionItem, removeRepeatingSectionItem };
