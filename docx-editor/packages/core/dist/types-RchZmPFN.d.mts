import { Schema, NodeSpec, MarkSpec } from 'prosemirror-model';
import { Plugin, Command } from 'prosemirror-state';

/**
 * Extension Manager
 *
 * Two-phase initialization:
 * 1. buildSchema() — collects NodeSpecs/MarkSpecs from extensions → new Schema
 * 2. initializeRuntime() — calls onSchemaReady() on each extension, collects plugins/commands/keymaps
 */

declare class ExtensionManager {
    private extensions;
    private schema;
    private plugins;
    private commands;
    constructor(extensions: AnyExtension[]);
    /**
     * Phase 1: Build schema from node/mark extensions
     */
    buildSchema(): void;
    /**
     * Phase 2: Initialize runtime (plugins, commands, keymaps)
     * Must be called after buildSchema()
     */
    initializeRuntime(): void;
    /**
     * Get the built schema
     */
    getSchema(): Schema;
    /**
     * Get all plugins (raw + keymap merged)
     */
    getPlugins(): Plugin[];
    /**
     * Get the flat command registry
     */
    getCommands(): CommandMap;
    /**
     * Get a specific command by name
     */
    getCommand(name: string): ((...args: any[]) => Command) | undefined;
    /**
     * Lifecycle: destroy
     */
    destroy(): void;
}

/**
 * Extension System Type Definitions
 *
 * Tiptap-style extension architecture for ProseMirror.
 * Three extension types:
 * - Extension: plugins, commands, keymaps (no schema)
 * - NodeExtension: adds a node spec to the schema
 * - MarkExtension: adds a mark spec to the schema
 */

type ExtensionPriority = number;
declare const Priority: {
    readonly Highest: 0;
    readonly High: 50;
    readonly Default: 100;
    readonly Low: 150;
    readonly Lowest: 200;
};
interface ExtensionContext {
    schema: Schema;
    /**
     * The manager that owns this extension. Use this in runtime callbacks
     * (e.g. `handleKeyDown`) that need to dispatch commands, instead of
     * reaching back to the `singletonManager` export — the latter forms a
     * circular import that breaks when the package is consumed as a built
     * bundle.
     */
    manager: ExtensionManager;
}
type CommandMap = Record<string, (...args: any[]) => Command>;
type KeyboardShortcutMap = Record<string, Command>;
interface ExtensionRuntime {
    commands?: CommandMap;
    keyboardShortcuts?: KeyboardShortcutMap;
    plugins?: Plugin[];
}
interface ExtensionConfig {
    name: string;
    priority: ExtensionPriority;
    options: Record<string, unknown>;
}
interface NodeExtensionConfig extends ExtensionConfig {
    schemaNodeName: string;
    nodeSpec: NodeSpec;
}
interface MarkExtensionConfig extends ExtensionConfig {
    schemaMarkName: string;
    markSpec: MarkSpec;
}
interface Extension {
    type: 'extension';
    config: ExtensionConfig;
    onSchemaReady(ctx: ExtensionContext): ExtensionRuntime;
}
interface NodeExtension {
    type: 'node';
    config: NodeExtensionConfig;
    onSchemaReady(ctx: ExtensionContext): ExtensionRuntime;
}
interface MarkExtension {
    type: 'mark';
    config: MarkExtensionConfig;
    onSchemaReady(ctx: ExtensionContext): ExtensionRuntime;
}
type AnyExtension = Extension | NodeExtension | MarkExtension;
interface ExtensionDefinition<TOptions = Record<string, unknown>> {
    name: string;
    priority?: ExtensionPriority;
    defaultOptions?: TOptions;
    onSchemaReady(ctx: ExtensionContext, options: TOptions): ExtensionRuntime;
}
interface NodeExtensionDefinition<TOptions = Record<string, unknown>> {
    name: string;
    priority?: ExtensionPriority;
    defaultOptions?: TOptions;
    schemaNodeName: string;
    nodeSpec: NodeSpec | ((options: TOptions) => NodeSpec);
    onSchemaReady?(ctx: ExtensionContext, options: TOptions): ExtensionRuntime;
}
interface MarkExtensionDefinition<TOptions = Record<string, unknown>> {
    name: string;
    priority?: ExtensionPriority;
    defaultOptions?: TOptions;
    schemaMarkName: string;
    markSpec: MarkSpec | ((options: TOptions) => MarkSpec);
    onSchemaReady?(ctx: ExtensionContext, options: TOptions): ExtensionRuntime;
}

export { type AnyExtension as A, type CommandMap as C, type Extension as E, type KeyboardShortcutMap as K, type MarkExtensionDefinition as M, type NodeExtension as N, Priority as P, ExtensionManager as a, type ExtensionDefinition as b, type MarkExtension as c, type NodeExtensionDefinition as d, type ExtensionConfig as e, type ExtensionContext as f, type ExtensionPriority as g, type ExtensionRuntime as h, type MarkExtensionConfig as i, type NodeExtensionConfig as j };
