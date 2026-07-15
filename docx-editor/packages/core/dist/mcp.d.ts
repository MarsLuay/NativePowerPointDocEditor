/**
 * MCP Server Module
 *
 * Model Context Protocol server for exposing document editing tools to AI clients.
 *
 * @example
 * ```ts
 * import { createMcpServer, startStdioServer } from '@npde/docx-editor/mcp';
 *
 * // Create server instance
 * const server = createMcpServer({ debug: true });
 *
 * // List available tools
 * console.log(server.listTools());
 *
 * // Call a tool programmatically
 * const result = await server.handleToolCall('docx_load', { content: base64 });
 * ```
 * @packageDocumentation
 * @public
 */
import { M as McpToolDefinition, e as McpSession, J as JsonSchema } from './types-BRTT_5Lv.js';
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
 * Core MCP Tools
 *
 * Built-in MCP tools for document operations that are always available.
 * These provide basic document manipulation without requiring plugins.
 */

/**
 * Load a DOCX document from base64
 */
declare const loadDocumentTool: McpToolDefinition;
/**
 * Save a document to base64
 */
declare const saveDocumentTool: McpToolDefinition;
/**
 * Close a document
 */
declare const closeDocumentTool: McpToolDefinition;
/**
 * Get document information
 */
declare const getDocumentInfoTool: McpToolDefinition;
/**
 * Get document plain text
 */
declare const getDocumentTextTool: McpToolDefinition;
declare const coreMcpTools: McpToolDefinition[];

/**
 * MCP Server
 *
 * Model Context Protocol server that exposes document editing tools to AI clients.
 * Discovers and registers tools from the plugin system plus core built-in tools.
 *
 * @example
 * ```ts
 * import { createMcpServer, startStdioServer } from '@npde/docx-editor/mcp';
 * import { pluginRegistry, docxtemplaterPlugin } from '@npde/docx-editor/core-plugins';
 *
 * // Register plugins
 * pluginRegistry.register(docxtemplaterPlugin);
 *
 * // Start MCP server
 * startStdioServer();
 * ```
 */

/**
 * MCP Server configuration
 */
interface McpServerConfig {
    /** Server name */
    name?: string;
    /** Server version */
    version?: string;
    /** Include core tools (default: true) */
    includeCoreTools?: boolean;
    /** Enable debug logging */
    debug?: boolean;
    /** Custom tools to add */
    additionalTools?: McpToolDefinition[];
}
/**
 * MCP Server instance
 */
interface McpServer {
    /** All registered tools */
    tools: Map<string, McpToolDefinition>;
    /** Active session */
    session: McpSession;
    /** Handle a tool call */
    handleToolCall(toolName: string, input: unknown): Promise<unknown>;
    /** List available tools */
    listTools(): McpToolInfo[];
    /** Get server info */
    getInfo(): {
        name: string;
        version: string;
        toolCount: number;
    };
}
/**
 * Tool info for listing
 */
interface McpToolInfo {
    name: string;
    description: string;
    inputSchema: JsonSchema;
    category?: string;
}
/**
 * Create an MCP server instance
 *
 * @param config - Server configuration
 * @returns MCP server instance
 */
declare function createMcpServer(config?: McpServerConfig): McpServer;
/**
 * JSON-RPC request
 */
interface JsonRpcRequest {
    jsonrpc: '2.0';
    id: string | number;
    method: string;
    params?: unknown;
}
/**
 * JSON-RPC response
 */
interface JsonRpcResponse {
    jsonrpc: '2.0';
    id: string | number;
    result?: unknown;
    error?: {
        code: number;
        message: string;
        data?: unknown;
    };
}
/**
 * Handle a JSON-RPC request
 */
declare function handleJsonRpcRequest(server: McpServer, request: JsonRpcRequest): Promise<JsonRpcResponse>;
/**
 * Start the MCP server with stdio transport
 *
 * Reads JSON-RPC requests from stdin, writes responses to stdout.
 * This is the standard way to run an MCP server for Claude Desktop.
 */
declare function startStdioServer(config?: McpServerConfig): Promise<void>;

export { type McpServer, type McpServerConfig, type McpToolInfo, closeDocumentTool, coreMcpTools, createMcpServer, getDocumentInfoTool, getDocumentTextTool, handleJsonRpcRequest, loadDocumentTool, saveDocumentTool, startStdioServer };
