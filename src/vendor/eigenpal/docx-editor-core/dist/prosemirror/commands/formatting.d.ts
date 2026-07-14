/**
 * Text Formatting Commands — thin re-exports from extension system
 *
 * Toggle marks, set marks, clear formatting, hyperlinks.
 * All implementations live in extensions/marks/; this file re-exports
 * for backward compatibility.
 * @packageDocumentation
 * @public
 */
import { Command, EditorState } from 'prosemirror-state';
import { MarkType, Mark } from 'prosemirror-model';
import { T as TextFormatting } from '../../formatting-JhqWT_XM.js';
import { T as TextColorAttrs } from '../../marks-B25ktqF1.js';
import '../../colors-C3vA7HUU.js';

/**
 * Shared mark utility functions
 *
 * setMark, removeMark, isMarkActive, getMarkAttr, marksToTextFormatting, textFormattingToMarks, clearFormatting
 */

/**
 * Check if a mark is active in the current selection
 */
declare function isMarkActive(state: EditorState, markType: MarkType, attrs?: Record<string, unknown>): boolean;
/**
 * Get the current value of a mark attribute
 */
declare function getMarkAttr(state: EditorState, markType: MarkType, attr: string): unknown | null;
/**
 * Clear all text formatting (remove all marks)
 */
declare const clearFormatting: Command;
/**
 * Create a command that sets a mark on the selection
 */
declare function createSetMarkCommand(markType: MarkType, attrs?: Record<string, unknown>): Command;
/**
 * Create a command that removes a mark from the selection
 */
declare function createRemoveMarkCommand(markType: MarkType): Command;

declare function isHyperlinkActive(state: EditorState): boolean;
declare function getHyperlinkAttrs(state: EditorState): {
    href: string;
    tooltip?: string;
} | null;
declare function getSelectedText(state: EditorState): string;
/**
 * Resolve the hyperlink mark + contiguous range that surrounds the
 * current cursor. Used by edit/remove popup actions in both adapters.
 *
 * Resolution order for the mark itself:
 *   1. `$from.marks()` — the normal active-marks lookup
 *   2. `$from.nodeAfter`/`nodeBefore` marks — boundary positions don't
 *      report active marks via `marks()`
 *   3. (optional) text-node search by `fallbackHref` — last resort when
 *      the popup knows the href but the cursor sits at a gap
 *
 * The returned range walks the parent block grouping consecutive text
 * nodes that share the same href, and returns whichever range contains
 * the cursor.
 */
declare function findHyperlinkRangeAt(state: EditorState, fallbackHref?: string): {
    mark: Mark;
    start: number;
    end: number;
} | null;

/**
 * Text Formatting Commands — thin re-exports from extension system
 *
 * Toggle marks, set marks, clear formatting, hyperlinks.
 * All implementations live in extensions/marks/; this file re-exports
 * for backward compatibility.
 * @packageDocumentation
 * @public
 */

declare function textFormattingToMarks(formatting: TextFormatting): Mark[];
declare const toggleBold: Command;
declare const toggleItalic: Command;
declare const toggleUnderline: Command;
declare const toggleStrike: Command;
declare const toggleSuperscript: Command;
declare const toggleSubscript: Command;
declare function setTextColor(attrs: TextColorAttrs): Command;
declare const clearTextColor: Command;
declare function setHighlight(color: string): Command;
declare const clearHighlight: Command;
declare function setFontSize(size: number): Command;
declare const clearFontSize: Command;
declare function setFontFamily(fontName: string): Command;
declare const clearFontFamily: Command;
declare function setUnderlineStyle(style: string, color?: TextColorAttrs): Command;
declare function setHyperlink(href: string, tooltip?: string): Command;
declare const removeHyperlink: Command;
declare function insertHyperlink(text: string, href: string, tooltip?: string): Command;

export { clearFontFamily, clearFontSize, clearFormatting, clearHighlight, clearTextColor, createRemoveMarkCommand, createSetMarkCommand, findHyperlinkRangeAt, getHyperlinkAttrs, getMarkAttr, getSelectedText, insertHyperlink, isHyperlinkActive, isMarkActive, removeHyperlink, setFontFamily, setFontSize, setHighlight, setHyperlink, setTextColor, setUnderlineStyle, textFormattingToMarks, toggleBold, toggleItalic, toggleStrike, toggleSubscript, toggleSuperscript, toggleUnderline };
