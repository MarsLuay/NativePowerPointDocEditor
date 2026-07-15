/**
 * ProseMirror Plugins
 *
 * Selection tracker plugin for the DOCX editor.
 * Keymap plugins are now provided by the extension system.
 * @packageDocumentation
 * @public
 */
export { SelectionChangeCallback, SelectionContext, createSelectionTrackerPlugin, extractSelectionContext, getSelectionContext, selectionTrackerKey } from './selectionTracker.mjs';
import { PluginKey, EditorState, Transaction, Plugin } from 'prosemirror-state';
export { D as DocumentContext, c as createDocumentContextPlugin, a as createDocumentStylesPlugin, d as documentContextKey, b as documentStylesKey, g as getDefaultTableStyleId, e as getDocumentStyleResolver, f as getDocumentTheme } from '../../documentStyles-B96jYfta.mjs';
import { a as RevisionInfo } from '../../content-BZ9rYecc.mjs';
import '../../formatting-DFtuRFQY.mjs';
import '../../colors-C3vA7HUU.mjs';
import '../../styles-BNjUANte.mjs';
import '../../styleResolver-DCaiQV4u.mjs';
import '../../docx/wrapTypes.mjs';
import '../../lists-CyGxd5Y2.mjs';
import '../../watermark-D90356ZM.mjs';

/**
 * Suggestion-mode plugin state + shared meta keys.
 *
 * Kept in its own module to avoid an import cycle: every handler file needs
 * the plugin key + meta constants but should not transitively pull in the
 * plugin factory.
 */

interface SuggestionModeState {
    active: boolean;
    author: string;
}
declare const suggestionModeKey: PluginKey<SuggestionModeState>;

/**
 * Suggesting-mode toggle / set / query commands. Each dispatches a meta on
 * the plugin key — the plugin's `state.apply` merges it into
 * `SuggestionModeState`. No-ops gracefully when the plugin isn't mounted.
 */

/**
 * Toggle suggesting mode on/off. The mounted `createSuggestionModePlugin`
 * is required for the dispatch to have any effect — without it the meta
 * is silently dropped. Returns `false` (no-op) if the plugin is missing.
 *
 * @example
 * ```ts
 * import { toggleSuggestionMode } from '@eigenpal/docx-editor-core/prosemirror/plugins';
 * toggleSuggestionMode(view.state, view.dispatch);
 * ```
 */
declare function toggleSuggestionMode(state: EditorState, dispatch?: (tr: Transaction) => void): boolean;
/**
 * Set suggesting mode active state and (optionally) author. Author
 * tracks across every revision minted while the mode is on. The
 * mounted `createSuggestionModePlugin` is required.
 *
 * @example
 * ```ts
 * import { setSuggestionMode } from '@eigenpal/docx-editor-core/prosemirror/plugins';
 * setSuggestionMode(true, view.state, view.dispatch, 'Jane');
 * // ... typed text now wraps in <w:ins> with author="Jane"
 * setSuggestionMode(false, view.state, view.dispatch);
 * ```
 */
declare function setSuggestionMode(active: boolean, state: EditorState, dispatch?: (tr: Transaction) => void, author?: string): boolean;
/**
 * Check if suggestion mode is currently active.
 */
declare function isSuggestionModeActive(state: EditorState): boolean;

/**
 * Suggestion Mode Plugin
 *
 * When active, intercepts all text insertions and deletions, wrapping
 * them in tracked-change marks (insertion/deletion) instead of modifying
 * the document directly.
 *
 * - Typed text is marked as insertion (green underline)
 * - Deleted text is NOT removed — it's marked as deletion (red strikethrough)
 * - Text already marked as insertion by the current author is deleted
 *   normally (retracting your own suggestion)
 *
 * The implementation is split across this directory for readability:
 *   - `state.ts`       — plugin key, meta constants, shared types
 *   - `markAttrs.ts`   — fresh-attr minting + projection
 *   - `adjacency.ts`   — coalescing lookups (sibling, cross-block, cellMarker)
 *   - `handlers/`      — keyboard / input handlers (delete, insert, structural)
 *   - `commands.ts`    — toggle / set / isActive
 *   - this file        — `createSuggestionModePlugin` + public re-exports
 */

/**
 * Create the suggestion-mode ProseMirror plugin. **Must be mounted on
 * the editor view for `setSuggestionMode` and `toggleSuggestionMode`
 * to do anything** — both adapters (`@eigenpal/docx-editor-react`,
 * `@eigenpal/docx-editor-vue`) auto-mount this inside the `DocxEditor`
 * component, so consumers using the bundled components don't need to
 * register it themselves.
 *
 * When active, typed text gets the `insertion` mark, deleted text gets
 * the `deletion` mark (text stays in the doc; the painter strikes it
 * through), Enter sets `pPrIns` on the originating paragraph, and
 * Backspace at paragraph start sets `pPrDel` on the previous paragraph.
 * Author + adjacent same-author marks coalesce into one tracked change.
 *
 * @param initialActive - Whether suggesting mode starts on. Default `false`.
 * @param author - Author name attached to every minted revision. Default `'User'`.
 *
 * @example
 * ```ts
 * import { createSuggestionModePlugin } from '@eigenpal/docx-editor-core/prosemirror/plugins';
 *
 * const plugin = createSuggestionModePlugin(false, 'Jane');
 * EditorState.create({ doc, plugins: [plugin, ...other] });
 * ```
 */
declare function createSuggestionModePlugin(initialActive?: boolean, author?: string): Plugin;

/**
 * Single source of tracked-revision ids across the package.
 *
 * Why a shared module-level counter:
 *
 *   - Each `revisionId` is the OOXML `<w:ins w:id="…"/>` attribute. Two
 *     unrelated revisions emitted by the same author with the same id
 *     would silently collapse in the sidebar (grouped by id+author+date)
 *     and on accept (resolved by id), so collisions are observable.
 *   - Pre-refactor, three call sites each kept their own
 *     `Date.now() + offset` counter (suggestionMode.ts, table commands
 *     delete.ts, table commands insert.ts). Offsets made first-load
 *     collisions rare but not impossible — counters drift independently,
 *     and parallel-browser Playwright workers can start from identical
 *     `Date.now()` seeds.
 *
 * Routing every mint through this module guarantees within-realm
 * uniqueness even when callers interleave across plugins, commands, and
 * test harnesses.
 *
 * Re-exported as `@public` through `prosemirror/plugins/index.ts` so
 * adapter integrations (image insertion, custom commands) can mint
 * revision triples without reaching into the implementation directly.
 *
 * @packageDocumentation
 * @public
 */

/**
 * Build a fresh `RevisionInfo` triple from the active suggesting-mode
 * state. Returns `null` when suggesting mode is OFF — callers use this to
 * decide whether to track an edit or apply it directly.
 *
 * Shared by `suggestionMode.ts`'s text/paragraph handlers and by the
 * suggesting-aware table commands (`addRowBelow`, `deleteRow`, ...). One
 * source of truth for both the mint and the author/date fields.
 */
declare function makeRevisionInfo(state: EditorState): RevisionInfo | null;

export { RevisionInfo, createSuggestionModePlugin, isSuggestionModeActive, makeRevisionInfo, setSuggestionMode, suggestionModeKey, toggleSuggestionMode };
