/**
 * Table Extension — 4 node specs + plugins + commands
 *
 * Uses separate NodeExtension instances for each table node type, plus an
 * Extension that registers the prosemirror-tables editing plugins, the
 * Backspace/Delete keymap chain, and the full 32-command surface.
 *
 * NodeSpecs (declarative attrs + parseDOM + toDOM) and the CSS-paste
 * helpers shared by td/th live in ./{specs,paste}.ts. The table-context
 * query / cell-navigation helpers live in ./context.ts. The plugin
 * runtime itself lives under ./commands/ — one file per command domain
 * (insert, delete, selection, borders, cellFormatting, sizing,
 * tableStyle), plus shared helpers and the active-cell decoration plugin.
 * @packageDocumentation
 * @public
 */
import { N as NodeExtension, E as Extension, A as AnyExtension } from '../../../types-RchZmPFN.mjs';
export { B as BorderPreset, a as BorderSpec, T as TableContextInfo, g as getTableContext, b as goToNextCell, c as goToPrevCell, i as isInTable } from '../../../borders-BH_GoYU3.mjs';
import 'prosemirror-model';
import 'prosemirror-state';
import '../../../colors-C3vA7HUU.mjs';

declare const TableNodeExtension: (options?: Partial<Record<string, unknown>> | undefined) => NodeExtension;
declare const TableRowExtension: (options?: Partial<Record<string, unknown>> | undefined) => NodeExtension;
declare const TableCellExtension: (options?: Partial<Record<string, unknown>> | undefined) => NodeExtension;
declare const TableHeaderExtension: (options?: Partial<Record<string, unknown>> | undefined) => NodeExtension;
declare const TablePluginExtension: (options?: Partial<Record<string, unknown>> | undefined) => Extension;
declare function createTableExtensions(): AnyExtension[];

export { TableCellExtension, TableHeaderExtension, TableNodeExtension, TablePluginExtension, TableRowExtension, createTableExtensions };
