/**
 * Core Plugin System
 *
 * Headless plugin system for extending document tooling with custom
 * commands and exposing MCP tools for AI integration.
 *
 * @example
 * ```ts
 * import {
 *   pluginRegistry,
 *   docxtemplaterPlugin,
 *   type CorePlugin
 * } from '@npde/docx-editor/core-plugins';
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
import { C as CorePlugin } from './types-B-iZWEFz.mjs';
export { b as CommandHandler, d as CommandResult, E as ExtractCommand, J as JsonSchema, L as LoadedDocument, e as McpSession, f as McpToolAnnotations, g as McpToolContent, h as McpToolContext, M as McpToolDefinition, n as McpToolExample, i as McpToolHandler, j as McpToolResult, k as PluginCommand, b as PluginCommandHandler, l as PluginEvent, c as PluginEventListener, P as PluginOptions, a as PluginRegistrationResult, M as ToolDefinition, i as ToolHandler, j as ToolResult, T as TypedCommandHandler, Z as ZodSchemaLike, m as isZodSchema } from './types-B-iZWEFz.mjs';
export { P as PluginRegistry, c as createPluginRegistrar, p as pluginRegistry, r as registerPlugins } from './registry-B78PDzyN.mjs';
import './types/document.mjs';
import './colors-C3vA7HUU.mjs';
import './formatting-DFtuRFQY.mjs';
import './lists-CyGxd5Y2.mjs';
import './content-BZ9rYecc.mjs';
import './docx/wrapTypes.mjs';
import './watermark-D90356ZM.mjs';
import './styles-Diw0MASy.mjs';
import './types/agentApi.mjs';

/**
 * Docxtemplater Plugin
 *
 * Core plugin for template variable functionality using docxtemplater.
 *
 * **Command handlers** — `insertTemplateVariable` and `replaceWithTemplateVariable`
 * allow host code to programmatically insert `{variable}` placeholders.
 *
 * @example
 * ```ts
 * import { pluginRegistry } from '@npde/docx-editor/core-plugins';
 * import { docxtemplaterPlugin } from '@npde/docx-editor/core-plugins/docxtemplater';
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
