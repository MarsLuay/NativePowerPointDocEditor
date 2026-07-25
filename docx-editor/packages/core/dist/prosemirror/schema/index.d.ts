/**
 * ProseMirror Schema for DOCX Editor
 *
 * Singleton ExtensionManager that builds the schema and initializes runtime.
 * Legacy code imports `schema` and commands from here; new code should use
 * ExtensionManager directly.
 * @packageDocumentation
 * @public
 */
import { Schema } from 'prosemirror-model';
import { a as ExtensionManager } from '../../types-RchZmPFN.js';
export { I as ImageAttrs, a as ImagePositionAttrs, P as ParagraphAttrs, T as TableAttrs, b as TableCellAttrs, c as TableRowAttrs } from '../../nodes-Bu9WifHU.js';
export { F as FontFamilyAttrs, a as FontSizeAttrs, H as HyperlinkAttrs, T as TextColorAttrs, U as UnderlineAttrs } from '../../marks-B25ktqF1.js';
import 'prosemirror-state';
import '../../colors-C3vA7HUU.js';
import '../../formatting-JhqWT_XM.js';
import '../../lists-Bn29SzeS.js';
import '../../docx/wrapTypes.js';
import '../../content-B8ScSBzC.js';
import '../../watermark-D90356ZM.js';

/**
 * ProseMirror Schema for DOCX Editor
 *
 * Singleton ExtensionManager that builds the schema and initializes runtime.
 * Legacy code imports `schema` and commands from here; new code should use
 * ExtensionManager directly.
 * @packageDocumentation
 * @public
 */

declare const singletonManager: ExtensionManager;
declare const schema: Schema<string, string>;
/**
 * Export types for convenience
 */
type DocxSchema = Schema<string, string>;
type DocxNode = ReturnType<typeof schema.node>;
type DocxMark = ReturnType<typeof schema.mark>;

export { type DocxMark, type DocxNode, type DocxSchema, schema, singletonManager };
