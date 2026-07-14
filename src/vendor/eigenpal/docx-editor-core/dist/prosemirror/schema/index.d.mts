/**
 * ProseMirror Schema for DOCX Editor
 *
 * Singleton ExtensionManager that builds the schema and initializes runtime.
 * Legacy code imports `schema` and commands from here; new code should use
 * ExtensionManager directly.
 * @packageDocumentation
 * @public
 */
import * as prosemirror_model from 'prosemirror-model';
import { a as ExtensionManager } from '../../types-RchZmPFN.mjs';
export { I as ImageAttrs, a as ImagePositionAttrs, P as ParagraphAttrs, T as TableAttrs, b as TableCellAttrs, c as TableRowAttrs } from '../../nodes-CHOFq5Ka.mjs';
export { F as FontFamilyAttrs, a as FontSizeAttrs, H as HyperlinkAttrs, T as TextColorAttrs, U as UnderlineAttrs } from '../../marks-DOMflBhB.mjs';
import 'prosemirror-state';
import '../../colors-C3vA7HUU.mjs';
import '../../formatting-DFtuRFQY.mjs';
import '../../lists-CyGxd5Y2.mjs';
import '../../docx/wrapTypes.mjs';
import '../../content-BZ9rYecc.mjs';
import '../../watermark-D90356ZM.mjs';

declare const singletonManager: ExtensionManager;
declare const schema: prosemirror_model.Schema<any, any>;
/**
 * Export types for convenience
 */
type DocxSchema = typeof schema;
type DocxNode = ReturnType<typeof schema.node>;
type DocxMark = ReturnType<typeof schema.mark>;

export { type DocxMark, type DocxNode, type DocxSchema, schema, singletonManager };
