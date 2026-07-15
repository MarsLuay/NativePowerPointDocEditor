/**
 * Core Plugin API — Framework-Agnostic
 *
 * Exports the core plugin interfaces and types that can be used
 * by any framework adapter (React, Vue, etc.).
 *
 * @experimental Plugin API is still evolving. Breaking changes may
 * happen in minor releases until plugin authors stabilize the contract.
 * @packageDocumentation
 * @public
 */
export { EditorPluginCore, PanelConfig, PluginPanelProps, PositionCoordinates, RenderedDomContext, SidebarItem, SidebarItemContext } from './types.js';
import 'prosemirror-state';
import 'prosemirror-view';
import 'prosemirror-model';
