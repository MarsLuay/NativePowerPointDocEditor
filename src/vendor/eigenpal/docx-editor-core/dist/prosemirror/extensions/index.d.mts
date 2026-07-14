/**
 * Extension System — Barrel Export
 * @packageDocumentation
 * @public
 */
import { b as ExtensionDefinition, E as Extension, M as MarkExtensionDefinition, c as MarkExtension, d as NodeExtensionDefinition, N as NodeExtension, A as AnyExtension } from '../../types-RchZmPFN.mjs';
export { C as CommandMap, e as ExtensionConfig, f as ExtensionContext, a as ExtensionManager, g as ExtensionPriority, h as ExtensionRuntime, K as KeyboardShortcutMap, i as MarkExtensionConfig, j as NodeExtensionConfig, P as Priority } from '../../types-RchZmPFN.mjs';
import { SelectionChangeCallback } from '../plugins/selectionTracker.mjs';
import { EditorState, Transaction } from 'prosemirror-state';
export { TableCellExtension, TableHeaderExtension, TableNodeExtension, TableRowExtension } from './nodes/TableExtension.mjs';
export { T as TableContextInfo } from '../../borders-BH_GoYU3.mjs';
import 'prosemirror-model';
import '../../formatting-DFtuRFQY.mjs';
import '../../colors-C3vA7HUU.mjs';

/**
 * Extension Factory Functions
 *
 * Creates extension instances from definitions.
 * Each factory returns a function that accepts options and returns an extension instance.
 */

/**
 * Create a generic extension (plugins, commands, keymaps — no schema contribution)
 */
declare function createExtension<TOptions extends Record<string, unknown> = Record<string, unknown>>(def: ExtensionDefinition<TOptions>): (options?: Partial<TOptions>) => Extension;
/**
 * Create a node extension (contributes a NodeSpec to the schema)
 */
declare function createNodeExtension<TOptions extends Record<string, unknown> = Record<string, unknown>>(def: NodeExtensionDefinition<TOptions>): (options?: Partial<TOptions>) => NodeExtension;
/**
 * Create a mark extension (contributes a MarkSpec to the schema)
 */
declare function createMarkExtension<TOptions extends Record<string, unknown> = Record<string, unknown>>(def: MarkExtensionDefinition<TOptions>): (options?: Partial<TOptions>) => MarkExtension;

/**
 * StarterKit — bundles all extensions into a ready-to-use set
 *
 * Usage:
 *   const extensions = createStarterKit();
 *   const manager = new ExtensionManager(extensions);
 *   manager.buildSchema();
 *   manager.initializeRuntime();
 */

interface StarterKitOptions {
    /** Extensions to disable by name */
    disable?: string[];
    /** History depth (default: 100) */
    historyDepth?: number;
    /** History new group delay (default: 500) */
    historyNewGroupDelay?: number;
    /** Selection change callback */
    onSelectionChange?: SelectionChangeCallback;
}
/**
 * Create the full set of extensions for the DOCX editor
 */
declare function createStarterKit(options?: StarterKitOptions): AnyExtension[];

/**
 * Get the set of changed paragraph IDs from an EditorState
 */
declare function getChangedParagraphIds(state: EditorState): Set<string>;
/**
 * Check if structural changes (paragraph add/delete) occurred
 */
declare function hasStructuralChanges(state: EditorState): boolean;
/**
 * Check if any changes affected paragraphs without paraId
 */
declare function hasUntrackedChanges(state: EditorState): boolean;
/**
 * Create a transaction that clears the change tracker
 */
declare function clearTrackedChanges(state: EditorState): Transaction;
declare const ParagraphChangeTrackerExtension: (options?: Partial<{}> | undefined) => Extension;

export { AnyExtension, Extension, ExtensionDefinition, MarkExtension, MarkExtensionDefinition, NodeExtension, NodeExtensionDefinition, ParagraphChangeTrackerExtension, type StarterKitOptions, clearTrackedChanges, createExtension, createMarkExtension, createNodeExtension, createStarterKit, getChangedParagraphIds, hasStructuralChanges, hasUntrackedChanges };
