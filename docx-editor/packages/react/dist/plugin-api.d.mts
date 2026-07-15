/**
 * @npde/docx-editor-react/plugin-api
 *
 * Generic plugin interface and host component for integrating external
 * plugins with the editor. Pairs with the framework-agnostic plugin types
 * exported from `@npde/docx-editor-core/plugin-api`.
 *
 * @example
 * ```tsx
 * import { PluginHost, templatePlugin, type EditorPlugin } from '@npde/docx-editor-react/plugin-api';
 *
 * function MyEditor() {
 *   return (
 *     <PluginHost plugins={[templatePlugin]}>
 *       <DocxEditor document={doc} onChange={handleChange} />
 *     </PluginHost>
 *   );
 * }
 * ```
 *
 * @packageDocumentation
 * @public
 */
import { P as PluginHostProps, a as PluginHostRef, b as ReactEditorPlugin } from './types-D4kZpqv7.mjs';
export { E as EditorPlugin, c as PluginContext, R as ReactSidebarItem, S as SidebarItemRenderProps } from './types-D4kZpqv7.mjs';
import * as React from 'react';
export { RenderedDomContextImpl, createRenderedDomContext } from '@npde/docx-editor-core/plugin-api/RenderedDomContext';
import { TemplateTag } from '@npde/docx-editor-core/prosemirror/template/prosemirror-plugin';
export { TEMPLATE_DECORATION_STYLES, TagType, TemplateTag, createTemplatePlugin as createTemplateProseMirrorPlugin, getTemplateTags as getTemplatePluginTags, setHoveredElement, setSelectedElement, templatePluginKey } from '@npde/docx-editor-core/prosemirror/template/prosemirror-plugin';
export { PanelConfig, PluginPanelProps, PositionCoordinates, RenderedDomContext, SidebarItem, SidebarItemContext } from '@npde/docx-editor-core/plugin-api';
import 'prosemirror-view';

declare const PLUGIN_HOST_STYLES = "\n.plugin-host {\n  display: flex;\n  width: 100%;\n  height: 100%;\n  overflow: visible;\n  position: relative;\n}\n\n.plugin-host-editor {\n  flex: 1;\n  display: flex;\n  flex-direction: column;\n  min-width: 0;\n  overflow: visible;\n}\n\n\n.plugin-panels-left,\n.plugin-panels-right {\n  display: flex;\n  flex-direction: column;\n  flex-shrink: 0;\n  background: #f8f9fa;\n  border-color: #e9ecef;\n}\n\n.plugin-panels-left {\n  border-right: 1px solid #e9ecef;\n}\n\n.plugin-panels-right {\n  border-left: 1px solid #e9ecef;\n}\n\n.plugin-panels-bottom {\n  border-top: 1px solid #e9ecef;\n  background: #f8f9fa;\n}\n\n.plugin-panel {\n  position: relative;\n  display: flex;\n  flex-direction: column;\n  overflow: hidden;\n  transition: width 0.2s ease, height 0.2s ease;\n}\n\n.plugin-panel.collapsed {\n  overflow: visible;\n}\n\n.plugin-panel-toggle {\n  display: flex;\n  align-items: center;\n  gap: 4px;\n  padding: 6px 8px;\n  background: transparent;\n  border: none;\n  cursor: pointer;\n  font-size: 12px;\n  color: #6c757d;\n  white-space: nowrap;\n}\n\n.plugin-panel.collapsed .plugin-panel-toggle {\n  writing-mode: vertical-rl;\n  text-orientation: mixed;\n  flex-direction: column;\n  height: 100%;\n  padding: 8px 6px;\n}\n\n.plugin-panel-toggle:hover {\n  background: #e9ecef;\n  color: #495057;\n}\n\n.plugin-panel-toggle-icon {\n  font-weight: bold;\n  font-size: 14px;\n}\n\n.plugin-panel.collapsed .plugin-panel-toggle-icon {\n  transform: rotate(90deg);\n}\n\n.plugin-panel-toggle-label {\n  font-weight: 500;\n}\n\n.plugin-panel-content {\n  flex: 1;\n  overflow: auto;\n}\n\n/* Right panel rendered inside viewport - scrolls with content */\n.plugin-panel-in-viewport {\n  position: absolute;\n  top: 0;\n  /* Position is set dynamically via inline styles based on page edge */\n  width: 220px;\n  pointer-events: auto;\n  z-index: 10;\n  overflow: visible;\n}\n\n.plugin-panel-in-viewport.collapsed {\n  width: 32px;\n}\n\n.plugin-panel-in-viewport .plugin-panel-toggle {\n  position: sticky;\n  top: 0;\n  background: rgba(255, 255, 255, 0.95);\n  border-radius: 4px;\n  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);\n}\n\n.plugin-panel-in-viewport-content {\n  overflow: visible;\n  position: relative;\n}\n\n/* Plugin overlay container for rendering highlights/decorations */\n.plugin-overlays-container {\n  position: absolute;\n  top: 0;\n  left: 0;\n  right: 0;\n  bottom: 0;\n  pointer-events: none;\n  overflow: visible;\n  z-index: 5;\n}\n\n.plugin-overlay {\n  position: absolute;\n  top: 0;\n  left: 0;\n  right: 0;\n  bottom: 0;\n  pointer-events: none;\n}\n\n/* Individual overlay children manage their own pointer-events.\n   Do NOT set pointer-events: auto here \u2014 it overrides overlay containers\n   that need pointer-events: none to let clicks pass through to the editor. */\n";
/**
 * PluginHost Component
 *
 * Wraps the editor and provides:
 * - Plugin state management
 * - Panel rendering for each plugin
 * - CSS injection for plugin styles
 * - Callbacks for editor interaction
 */
declare const PluginHost: React.ForwardRefExoticComponent<PluginHostProps & React.RefAttributes<PluginHostRef>>;

/**
 * Template Plugin
 *
 * Docxtemplater template support as a plugin for the DOCX Editor.
 *
 * Features:
 * - Full docxtemplater syntax detection (variables, loops, conditionals)
 * - Sidebar annotation chips showing template structure (via getSidebarItems)
 * - Differentiated visual highlighting by element type
 *
 * @example
 * ```tsx
 * import { PluginHost, templatePlugin } from '@npde/docx-editor-react/plugin-api';
 *
 * function MyEditor() {
 *   return (
 *     <PluginHost plugins={[templatePlugin]}>
 *       <DocxEditor document={doc} onChange={handleChange} />
 *     </PluginHost>
 *   );
 * }
 * ```
 */

interface TemplatePluginState {
    tags: TemplateTag[];
    hoveredId?: string;
    selectedId?: string;
}
/**
 * Create the template plugin instance.
 */
declare function createPlugin(_options?: {
    /** @deprecated — panel is no longer used; template chips render in the unified sidebar */
    defaultCollapsed?: boolean;
    /** @deprecated */
    panelPosition?: 'left' | 'right';
    /** @deprecated */
    panelWidth?: number;
}): ReactEditorPlugin<TemplatePluginState>;
/**
 * Default template plugin instance.
 */
declare const templatePlugin: ReactEditorPlugin<TemplatePluginState>;

export { PLUGIN_HOST_STYLES, PluginHost, PluginHostProps, PluginHostRef, ReactEditorPlugin, createPlugin as createTemplatePlugin, templatePlugin };
