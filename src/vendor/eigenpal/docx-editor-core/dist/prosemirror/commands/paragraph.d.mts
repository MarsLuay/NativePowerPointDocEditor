/**
 * Paragraph Formatting Commands — thin re-exports from extension system
 *
 * Alignment, line spacing, indentation, lists, paragraph styles.
 * All implementations live in extensions/; this file re-exports
 * for backward compatibility.
 * @packageDocumentation
 * @public
 */
import { EditorState, Command } from 'prosemirror-state';
import { P as ParagraphFormatting, T as TextFormatting, d as ParagraphAlignment, e as TabStop, i as TabStopAlignment, h as TabLeader, L as LineSpacingRule } from '../../formatting-DFtuRFQY.mjs';
import { N as NumberingMap } from '../../numberingParser-nSPZlzKr.mjs';
import '../../colors-C3vA7HUU.mjs';
import '../../lists-CyGxd5Y2.mjs';

interface ResolvedStyleAttrs {
    paragraphFormatting?: ParagraphFormatting;
    runFormatting?: TextFormatting;
    /**
     * Numbering definitions from the document package. When the applied style
     * carries a `w:numPr`, these resolve the numbering level into the list
     * marker attrs (template, per-level formats, counter key) so the painter
     * renders the style's numbering — e.g. "[Claim 1]" — instead of falling
     * back to a plain decimal marker.
     */
    numbering?: NumberingMap | null;
}
declare function getParagraphAlignment(state: EditorState): ParagraphAlignment | null;
declare function getParagraphTabs(state: EditorState): TabStop[] | null;
declare function getStyleId(state: EditorState): string | null;
declare function getParagraphBidi(state: EditorState): boolean;

declare function isInList(state: EditorState): boolean;
declare function getListInfo(state: EditorState): {
    numId: number;
    ilvl: number;
} | null;

/**
 * Paragraph Formatting Commands — thin re-exports from extension system
 *
 * Alignment, line spacing, indentation, lists, paragraph styles.
 * All implementations live in extensions/; this file re-exports
 * for backward compatibility.
 * @packageDocumentation
 * @public
 */

declare function setAlignment(alignment: ParagraphAlignment): Command;
declare const alignLeft: Command;
declare const alignCenter: Command;
declare const alignRight: Command;
declare const alignJustify: Command;
declare function setLineSpacing(value: number, rule?: LineSpacingRule): Command;
declare const singleSpacing: Command;
declare const oneAndHalfSpacing: Command;
declare const doubleSpacing: Command;
declare function increaseIndent(amount?: number): Command;
declare function decreaseIndent(amount?: number): Command;
declare function setIndentLeft(twips: number): Command;
declare function setIndentRight(twips: number): Command;
declare function setIndentFirstLine(twips: number, hanging?: boolean): Command;
declare const toggleBulletList: Command;
declare const toggleNumberedList: Command;
declare const increaseListLevel: Command;
declare const decreaseListLevel: Command;
declare const removeList: Command;
declare function setSpaceBefore(twips: number): Command;
declare function setSpaceAfter(twips: number): Command;
declare function applyStyle(styleId: string, resolvedAttrs?: ResolvedStyleAttrs): Command;
declare const clearStyle: Command;
declare function insertSectionBreak(breakType: 'nextPage' | 'continuous' | 'oddPage' | 'evenPage'): Command;
declare const removeSectionBreak: Command;
declare function addTabStop(position: number, alignment?: TabStopAlignment, leader?: TabLeader): Command;
declare function removeTabStop(position: number): Command;
declare const setRtl: Command;
declare const setLtr: Command;
declare const generateTOC: Command;

export { type ResolvedStyleAttrs, addTabStop, alignCenter, alignJustify, alignLeft, alignRight, applyStyle, clearStyle, decreaseIndent, decreaseListLevel, doubleSpacing, generateTOC, getListInfo, getParagraphAlignment, getParagraphBidi, getParagraphTabs, getStyleId, increaseIndent, increaseListLevel, insertSectionBreak, isInList, oneAndHalfSpacing, removeList, removeSectionBreak, removeTabStop, setAlignment, setIndentFirstLine, setIndentLeft, setIndentRight, setLineSpacing, setLtr, setRtl, setSpaceAfter, setSpaceBefore, singleSpacing, toggleBulletList, toggleNumberedList };
