/**
 * @eigenpal/docx-editor-agents/mcp
 *
 * Model Context Protocol (MCP) server for the docx editor agent bridge.
 *
 * Two transports, same core:
 *   - `stdio`: classic MCP transport. Use `runStdioServer(bridge)` from a
 *     Node subprocess that Claude Desktop, Cursor, or any MCP-aware client
 *     will spawn. Newline-delimited JSON-RPC.
 *   - `direct`: call `new McpServer(bridge).handle(message)` if you have
 *     your own transport (WebSocket, `postMessage`, HTTP long-poll, etc.).
 *
 * The server is transport-agnostic and zero-dep. The stdio module reaches
 * for `process.stdin` / `process.stdout` only when you call
 * `runStdioServer` without explicit streams.
 *
 * @packageDocumentation
 * @public
 */
import { E as EditorBridge } from './server-29plFYqd.mjs';

/**
 * MCP wire protocol (subset) — JSON-RPC 2.0 framing + the message types we
 * actually implement. Zero dependencies. Pure functions; everything is unit-
 * testable without a transport.
 *
 * This is NOT a full MCP SDK. We implement only what the server needs:
 *   - initialize / initialized
 *   - tools/list
 *   - tools/call
 *   - notifications/cancelled (no-op, accepted)
 *
 * Spec reference: https://spec.modelcontextprotocol.io
 */
type JsonRpcId = string | number | null;
interface JsonRpcRequest {
    jsonrpc: '2.0';
    id: JsonRpcId;
    method: string;
    params?: unknown;
}
interface JsonRpcNotification {
    jsonrpc: '2.0';
    method: string;
    params?: unknown;
}
interface JsonRpcSuccess {
    jsonrpc: '2.0';
    id: JsonRpcId;
    result: unknown;
}
interface JsonRpcError {
    jsonrpc: '2.0';
    id: JsonRpcId;
    error: {
        code: number;
        message: string;
        data?: unknown;
    };
}
type JsonRpcResponse = JsonRpcSuccess | JsonRpcError;
type JsonRpcMessage = JsonRpcRequest | JsonRpcNotification | JsonRpcResponse;
/** Standard JSON-RPC error codes. We only ever emit JSON-RPC errors for
 * protocol-level problems; tool execution failures use MCP's `isError`
 * envelope inside a successful response, per spec. */
declare const ErrorCode: {
    readonly ParseError: -32700;
    readonly InvalidRequest: -32600;
    readonly MethodNotFound: -32601;
    readonly InvalidParams: -32602;
    readonly InternalError: -32603;
};
interface McpInitializeResult {
    protocolVersion: string;
    capabilities: {
        tools?: Record<string, unknown>;
    };
    serverInfo: {
        name: string;
        version: string;
    };
}
interface McpToolDescriptor {
    name: string;
    description: string;
    inputSchema: Record<string, unknown>;
}
interface McpToolsListResult {
    tools: McpToolDescriptor[];
}
interface McpToolsCallParams {
    name: string;
    arguments?: Record<string, unknown>;
}
interface McpContent {
    type: 'text';
    text: string;
}
interface McpToolsCallResult {
    content: McpContent[];
    isError?: boolean;
}
declare function isJsonRpcRequest(m: unknown): m is JsonRpcRequest;
declare function isJsonRpcNotification(m: unknown): m is JsonRpcNotification;
declare function makeSuccess(id: JsonRpcId, result: unknown): JsonRpcSuccess;
declare function makeError(id: JsonRpcId, code: number, message: string, data?: unknown): JsonRpcError;
interface ParseResult {
    /** Parsed messages. */
    messages: JsonRpcMessage[];
    /** Lines that failed to parse — caller should send a ParseError per line if it had a discernible id. */
    parseErrors: string[];
    /** Remaining buffer (no trailing newline yet). */
    rest: string;
}
/**
 * Parse newline-delimited JSON-RPC frames out of a buffer. Returns parsed
 * messages plus any leftover bytes. Tolerates blank lines.
 */
declare function parseFrames(buffer: string): ParseResult;
/** Encode a JSON-RPC message as a single newline-terminated frame. */
declare function encodeFrame(message: JsonRpcMessage): string;

/**
 * MCP server core. Transport-agnostic — accepts a JsonRpcRequest, returns
 * either a JsonRpcResponse or `null` (for notifications, which never reply).
 *
 * Wraps an EditorBridge: tools/list returns the bridge's tool schemas in MCP
 * shape; tools/call dispatches via executeToolCall and converts the
 * AgentToolResult into MCP CallToolResult content.
 */

interface McpServerOptions {
    /** Server name reported in `initialize` response. Default: `@eigenpal/docx-editor-agents`. */
    name?: string;
    /** Server version. Default: `0.0.0` (override at build time). */
    version?: string;
    /** MCP protocol version we claim to speak. Default: `2025-06-18`. */
    protocolVersion?: string;
}
declare class McpServer {
    private readonly bridge;
    private readonly opts;
    constructor(bridge: EditorBridge, options?: McpServerOptions);
    /**
     * Handle one inbound message. Returns the response to send back, or `null`
     * for notifications and other no-reply messages. Never throws.
     */
    handle(message: JsonRpcMessage): JsonRpcResponse | null;
    private handleInitialize;
    private handleToolsList;
    private handleToolsCall;
}

/**
 * Stdio transport for the MCP server. Reads newline-delimited JSON-RPC from
 * an input stream, dispatches via McpServer, writes responses to an output
 * stream. Pure stream handling — no Node-only assumptions beyond "Readable
 * has .on('data') and Writable has .write".
 *
 * For real Node usage:
 *   import { runStdioServer } from '@eigenpal/docx-editor-agents/mcp';
 *   runStdioServer(bridge);
 *
 * For tests, pass any EventEmitter-shaped Readable + a function-shaped
 * Writable; see __tests__/mcp/stdio.test.ts.
 */

/** Minimal duck-typed input stream — anything with `on('data', cb)` works. */
interface InputStream {
    on(event: 'data', listener: (chunk: Buffer | string) => void): unknown;
    on(event: 'end', listener: () => void): unknown;
    on(event: 'error', listener: (err: Error) => void): unknown;
}
/** Minimal duck-typed output stream — anything with a `write(string) => bool` works. */
interface OutputStream {
    write(chunk: string): boolean | void;
}
interface StdioServerOptions extends McpServerOptions {
    input?: InputStream;
    output?: OutputStream;
    /** Called with diagnostic strings (e.g. parse errors). Default: stderr. */
    log?: (msg: string) => void;
}
interface StdioServerHandle {
    /** Underlying server (for tests / introspection). */
    server: McpServer;
    /** Manually feed a raw chunk (used by tests; the live transport calls this internally). */
    feed: (chunk: string | Buffer) => void;
    /** Stop accepting input and reject further writes. Idempotent. */
    close: () => void;
}
/**
 * Wire an EditorBridge to a JSON-RPC stdio loop. Returns immediately; reading
 * happens via the stream listeners. Designed to be testable: pass in fake
 * streams, call `feed(...)` directly, then assert on what was written.
 */
declare function runStdioServer(bridge: EditorBridge, options?: StdioServerOptions): StdioServerHandle;

export { ErrorCode, type JsonRpcError, type JsonRpcId, type JsonRpcMessage, type JsonRpcNotification, type JsonRpcRequest, type JsonRpcResponse, type JsonRpcSuccess, type McpContent, type McpInitializeResult, McpServer, type McpServerOptions, type McpToolDescriptor, type McpToolsCallParams, type McpToolsCallResult, type McpToolsListResult, type StdioServerHandle, type StdioServerOptions, encodeFrame, isJsonRpcNotification, isJsonRpcRequest, makeError, makeSuccess, parseFrames, runStdioServer };
