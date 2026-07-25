/**
 * Template ProseMirror Plugin
 *
 * Simple plugin that finds template tags using regex and creates decorations.
 * No separate parsing layer - everything happens here.
 * @packageDocumentation
 * @public
 */
import * as prosemirror_state from 'prosemirror-state';
import { Plugin, PluginKey } from 'prosemirror-state';
import { DecorationSet, EditorView } from 'prosemirror-view';

/**
 * Template tag types
 */
type TagType = 'variable' | 'sectionStart' | 'sectionEnd' | 'invertedStart' | 'raw';
/**
 * A found template tag
 */
interface TemplateTag {
    id: string;
    type: TagType;
    name: string;
    rawTag: string;
    from: number;
    to: number;
    /** For sections: nested variable names */
    nestedVars?: string[];
    /** True if this variable is inside a section (shown in section's nested vars) */
    insideSection?: boolean;
}
/**
 * Plugin state
 */
interface TemplatePluginState {
    tags: TemplateTag[];
    decorations: DecorationSet;
    hoveredId?: string;
    selectedId?: string;
}
/**
 * Plugin key
 */
declare const templatePluginKey: PluginKey<TemplatePluginState>;
/**
 * Create the template plugin
 */
declare function createTemplatePlugin(): Plugin<TemplatePluginState>;
/**
 * Get tags from editor state
 */
declare function getTemplateTags(state: prosemirror_state.EditorState): TemplateTag[];
/**
 * Set hovered tag
 */
declare function setHoveredElement(view: EditorView, id: string | undefined): void;
/**
 * Set selected tag
 */
declare function setSelectedElement(view: EditorView, id: string | undefined): void;
/**
 * CSS styles for template decorations
 */
declare const TEMPLATE_DECORATION_STYLES = "\n.docx-template-tag {\n  cursor: pointer;\n  transition: background-color 0.1s;\n}\n\n.docx-template-tag:hover,\n.docx-template-tag.hovered {\n  filter: brightness(0.95);\n}\n\n.docx-template-tag.selected {\n  box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.5);\n}\n";

export { TEMPLATE_DECORATION_STYLES, type TagType, type TemplateTag, createTemplatePlugin, getTemplateTags, setHoveredElement, setSelectedElement, templatePluginKey };
