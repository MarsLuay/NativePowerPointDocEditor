/**
 * @eigenpal/docx-editor-agents
 *
 * Word-like API for AI document review.
 *
 * @example
 * ```ts
 * const reviewer = await DocxReviewer.fromBuffer(buffer, 'AI Reviewer');
 *
 * // Read
 * const text = reviewer.getContentAsText();
 *
 * // Comment on a paragraph
 * reviewer.addComment(5, 'Liability cap seems too low.');
 *
 * // Replace text (creates tracked change)
 * reviewer.replace(5, '$50k', '$500k');
 *
 * // Or batch from LLM JSON response
 * reviewer.applyReview({
 *   comments: [{ paragraphIndex: 5, text: 'Too low.' }],
 *   proposals: [{ paragraphIndex: 5, search: '$50k', replaceWith: '$500k' }],
 * });
 *
 * const output = await reviewer.toBuffer();
 * ```
 *
 * @packageDocumentation
 * @public
 */
import { G as GetContentOptions, C as ContentBlock, a as CommentFilter, R as ReviewComment, b as ChangeFilter, c as ReviewChange, F as FoundMatch, S as SelectionInfo, A as AddCommentByParaIdOptions, d as ReplyOptions, P as ProposeChangeOptions, e as ApplyFormattingOptions, f as SetParagraphStyleOptions, g as PageContent, h as ScrollToParaIdOptions, i as ContentChangeEvent, j as SelectionChangeEvent } from './server-29plFYqd.js';
export { k as AgentToolDefinition, l as AgentToolResult, B as BatchError, m as BatchResult, n as BatchReviewOptions, D as DocxReviewer, o as agentTools, p as createReviewerBridge, q as executeToolCall, r as getToolSchemas } from './server-29plFYqd.js';

/**
 * Error classes for @eigenpal/docx-editor-agents
 */
declare class TextNotFoundError extends Error {
    constructor(search: string, paragraphIndex?: number);
}
declare class ChangeNotFoundError extends Error {
    constructor(id: number);
}
declare class CommentNotFoundError extends Error {
    constructor(id: number);
}

/**
 * Formal Word JS API parity contract.
 *
 * This file declares what subset of Microsoft Word's Office.js JavaScript API
 * (https://learn.microsoft.com/en-us/javascript/api/word) we mirror, **at the
 * type level**. It is checked at compile time: `EditorBridge` must satisfy
 * `WordCompatBridge` (verified via the static assertion at the bottom).
 *
 * The assertion is the source of truth. If you change the bridge surface and
 * forget to update parity, typecheck breaks.
 *
 * ## What we mirror (✓)
 *
 * | Word JS API                         | Our equivalent                        |
 * | ----------------------------------- | ------------------------------------- |
 * | `Range` (stable handle)             | `paraId: string`                      |
 * | `body.search(text)` → Ranges        | `findText(query, opts) → FoundMatch[]`|
 * | `range.insertComment(text)`         | `addComment({paraId, text, search?})` |
 * | `comment.reply(text)`               | `replyTo(commentId, {text})`          |
 * | `comment.resolved = true`           | `resolveComment(commentId)`           |
 * | `range.insertText(text, location)`  | `proposeChange({paraId, search, replaceWith})` (3 modes via empty-string semantics) |
 * | `document.getSelection()`           | `getSelection() → SelectionInfo|null` |
 * | `range.scrollIntoView()`            | `scrollTo(paraId)`                    |
 * | `commentCollection.getItems()`      | `getComments(filter?)`                |
 * | `body.paragraphs.getItems()`        | `getContent(opts?)`                   |
 * | `document.body.text`                | `getContentAsText(opts?)`             |
 * | `revisionCollection.getItems()`     | `getChanges(filter?)`                 |
 * | `Document.onContentChanged.add(...)`| `onContentChange(listener)`           |
 * | `Document.onSelectionChanged.add()` | `onSelectionChange(listener)`         |
 * | `Range.font.bold` / `italic` / `color` / `size` / `name` | `applyFormatting({paraId, search?, marks})` |
 * | `ParagraphFormat.style` / `applyStyle(...)` | `setParagraphStyle({paraId, styleId})` |
 * | `Range.insertBreak(BreakType.*)`    | `insertBreak({paraId, type})`         |
 *
 * ## Beyond Word's surface (paged-document affordances)
 *
 *  - `getPage(n)` / `getPages({from, to})` / `getTotalPages()` — Word's JS API
 *    doesn't model pages as first-class addressable units. We do, because the
 *    editor is paged. Backed by the layout-painter's page boundary state.
 *
 * ## Differences (intentional, documented)
 *
 *  - All "insertText" overloads collapse into `proposeChange` with empty-string
 *    semantics: replaceWith="" deletes; search="" inserts at paragraph end;
 *    both non-empty replaces. Word has separate `insertText(...,'Replace')`,
 *    `insertText(...,'Before')`, etc.; we found three modes were enough and
 *    serialize cleaner for LLM tool calls.
 *  - Tracked changes are always *suggestions* in our world. Word lets the
 *    range mutate directly; we always go through the tracked-change path so
 *    the human keeps the final say. (This matches the agent UX.)
 *  - Word's `Range.context.sync()` is unnecessary — every call is one PM
 *    transaction.
 *  - Word's `Range` is a stateful object. Ours is a plain `{paraId, search?}`
 *    JSON value, so it survives JSON.stringify and works for tool calls / MCP.
 *
 * ## Out of scope (gaps we deliberately don't implement)
 *
 *  - Paragraph creation (`body.insertParagraph`). Out of scope for v1.
 *  - Table mutation (insert row/col, delete cell). Read-only.
 *  - Headers / footers / sections.
 *  - `Range.getOoxml()` / `getHtml()`. Plain text only.
 *  - `customXmlParts` / `contentControls`.
 *  - Accept / reject tracked changes. Human-only by design.
 *
 * Future versions can grow these by extending this interface — typecheck
 * will then enforce the new contract.
 */

/**
 * The formal Word-JS-API parity surface. Each method maps to one or more
 * Word API methods (named in the JSDoc above each member).
 *
 * If you change the EditorBridge surface, this interface must change too —
 * the static assertion at the bottom of the file enforces it.
 */
interface WordCompatBridge {
    /** Word: `document.body.text` (stringified, indexed lines). */
    getContentAsText(options?: GetContentOptions): string;
    /** Word: `body.paragraphs.getItems()`. */
    getContent(options?: GetContentOptions): ContentBlock[];
    /** Word: `commentCollection.getItems()`. */
    getComments(filter?: CommentFilter): ReviewComment[];
    /** Word: `revisionCollection.getItems()`. */
    getChanges(filter?: ChangeFilter): ReviewChange[];
    /** Word: `body.search(text)` returning Range[]. */
    findText(query: string, options?: {
        caseSensitive?: boolean;
        limit?: number;
    }): FoundMatch[];
    /** Word: `document.getSelection()`. */
    getSelection(): SelectionInfo | null;
    /** Word: `range.insertComment(text)`. Anchored by paraId (Range handle). */
    addComment(options: AddCommentByParaIdOptions): number | null;
    /** Word: `comment.reply(text)`. */
    replyTo(commentId: number, options: ReplyOptions): number | null;
    /** Word: `comment.resolved = true`. */
    resolveComment(commentId: number): void;
    /**
     * Word: `range.insertText(text, location)` collapsed into one verb.
     *  - replacement: search non-empty, replaceWith non-empty
     *  - deletion:    search non-empty, replaceWith ""
     *  - insertion:   search "",        replaceWith non-empty (paragraph end)
     */
    proposeChange(options: ProposeChangeOptions): boolean;
    /**
     * Word: `range.font.bold = true` / `range.font.italic = true` / etc.
     * Applied to a paragraph or to a unique phrase within it. Direct edit,
     * not a tracked change.
     */
    applyFormatting(options: ApplyFormattingOptions): boolean;
    /** Word: `paragraph.styleBuiltIn = ...` / `paragraph.style = 'Heading 1'`. */
    setParagraphStyle(options: SetParagraphStyleOptions): boolean;
    /**
     * Read one rendered page (1-indexed). Word's JS API does not expose pages
     * as first-class objects; we do because the editor is paged.
     */
    getPage(pageNumber: number): PageContent | null;
    /** Read a range of rendered pages (1-indexed, inclusive). */
    getPages(options: {
        from: number;
        to: number;
    }): PageContent[];
    /** Total number of pages currently rendered. */
    getTotalPages(): number;
    /** Word: `range.scrollIntoView()`. */
    scrollTo(paraId: string, options?: ScrollToParaIdOptions): boolean;
    /** Word: `Document.onContentChanged.add(handler)`. Returns unsubscribe. */
    onContentChange(listener: (event: ContentChangeEvent) => void): () => void;
    /** Word: `Document.onSelectionChanged.add(handler)`. Returns unsubscribe. */
    onSelectionChange(listener: (event: SelectionChangeEvent) => void): () => void;
}

export { ChangeNotFoundError, CommentNotFoundError, ContentBlock, GetContentOptions, ReviewChange, ReviewComment, TextNotFoundError, type WordCompatBridge };
