/**
 * Selection Tracker Plugin
 *
 * Tracks selection changes and emits events for toolbar state updates.
 * Provides the current selection context including:
 * - Text formatting at cursor/selection
 * - Paragraph formatting
 * - Selection range information
 * @packageDocumentation
 * @public
 */
import { Plugin, EditorState, PluginKey } from 'prosemirror-state';
import { T as TextFormatting, P as ParagraphFormatting } from '../../formatting-JhqWT_XM.js';
import '../../colors-C3vA7HUU.js';

/**
 * Selection Tracker Plugin
 *
 * Tracks selection changes and emits events for toolbar state updates.
 * Provides the current selection context including:
 * - Text formatting at cursor/selection
 * - Paragraph formatting
 * - Selection range information
 * @packageDocumentation
 * @public
 */

/**
 * Selection context for toolbar state
 */
interface SelectionContext {
    /** Whether there's a non-collapsed selection */
    hasSelection: boolean;
    /** Whether selection spans multiple paragraphs */
    isMultiParagraph: boolean;
    /** Current text formatting at cursor/selection */
    textFormatting: TextFormatting;
    /** Current paragraph formatting */
    paragraphFormatting: ParagraphFormatting;
    /** Start paragraph index */
    startParagraphIndex: number;
    /** End paragraph index */
    endParagraphIndex: number;
    /** Whether cursor is in a list */
    inList: boolean;
    /** List type if in list */
    listType?: 'bullet' | 'numbered';
    /** List level (0-8) */
    listLevel?: number;
    /** Active comment IDs at cursor position */
    activeCommentIds: number[];
    /** Whether cursor is inside a tracked insertion */
    inInsertion: boolean;
    /** Whether cursor is inside a tracked deletion */
    inDeletion: boolean;
}
/**
 * Plugin key for accessing selection tracker state
 */
declare const selectionTrackerKey: PluginKey<SelectionContext>;
/**
 * Callback type for selection changes
 */
type SelectionChangeCallback = (context: SelectionContext) => void;
/**
 * Extract selection context from editor state
 */
declare function extractSelectionContext(state: EditorState): SelectionContext;
/**
 * Create selection tracker plugin
 */
declare function createSelectionTrackerPlugin(onSelectionChange?: SelectionChangeCallback): Plugin;
/**
 * Get current selection context from editor state
 */
declare function getSelectionContext(state: EditorState): SelectionContext | null;

export { type SelectionChangeCallback, type SelectionContext, createSelectionTrackerPlugin, extractSelectionContext, getSelectionContext, selectionTrackerKey };
