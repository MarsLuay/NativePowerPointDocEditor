/**
 * Core Plugin System
 *
 * Headless plugin system for extending DocumentAgent with custom
 * commands and exposing MCP tools for AI integration.
 *
 * @example
 * ```ts
 * import {
 *   pluginRegistry,
 *   docxtemplaterPlugin,
 *   type CorePlugin
 * } from '@eigenpal/docx-editor/core-plugins';
 *
 * // Register the docxtemplater plugin
 * pluginRegistry.register(docxtemplaterPlugin);
 *
 * // Get MCP tools for MCP server
 * const tools = pluginRegistry.getMcpTools();
 *
 * // Check available command handlers
 * const commandTypes = pluginRegistry.getCommandTypes();
 * ```
 * @packageDocumentation
 * @public
 */
import { C as CorePlugin } from './types-DDQWU7-j.js';
export { b as CommandHandler, d as CommandResult, E as ExtractCommand, J as JsonSchema, L as LoadedDocument, e as McpSession, f as McpToolAnnotations, g as McpToolContent, h as McpToolContext, M as McpToolDefinition, n as McpToolExample, i as McpToolHandler, j as McpToolResult, k as PluginCommand, b as PluginCommandHandler, l as PluginEvent, c as PluginEventListener, P as PluginOptions, a as PluginRegistrationResult, M as ToolDefinition, i as ToolHandler, j as ToolResult, T as TypedCommandHandler, Z as ZodSchemaLike, m as isZodSchema } from './types-DDQWU7-j.js';
export { P as PluginRegistry, c as createPluginRegistrar, p as pluginRegistry, r as registerPlugins } from './registry-egK81C_z.js';
import './types/document.js';
import './colors-C3vA7HUU.js';
import './formatting-JhqWT_XM.js';
import './lists-Bn29SzeS.js';
import './content-B8ScSBzC.js';
import './docx/wrapTypes.js';
import './watermark-D90356ZM.js';
import './styles-2J4U-Lgk.js';
import './types/agentApi.js';

/**
 * Docxtemplater Plugin
 *
 * Core plugin for template variable functionality using docxtemplater.
 *
 * **Command handlers** — `insertTemplateVariable` and `replaceWithTemplateVariable`
 * allow DocumentAgent to programmatically insert `{variable}` placeholders.
 *
 * @example
 * ```ts
 * import { pluginRegistry } from '@eigenpal/docx-editor/core-plugins';
 * import { docxtemplaterPlugin } from '@eigenpal/docx-editor/core-plugins/docxtemplater';
 *
 * pluginRegistry.register(docxtemplaterPlugin);
 * ```
 */

/**
 * Docxtemplater plugin for template variable functionality.
 *
 * Dependency validation is handled lazily by `processTemplate` at call time,
 * so no eager `initialize()` is needed.
 */
declare const docxtemplaterPlugin: CorePlugin;

export { CorePlugin, CorePlugin as Plugin, docxtemplaterPlugin };
