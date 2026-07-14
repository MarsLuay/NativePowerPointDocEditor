/**
 * @eigenpal/docx-editor-agents/react
 *
 * React entry. Hooks, components, and types that need React as a peer
 * dependency. Pair with `/server` (or `/ai-sdk/server`) for the API route
 * that drives the LLM.
 *
 * @example
 * ```tsx
 * import { useDocxAgentTools } from '@eigenpal/docx-editor-agents/react';
 *
 * const { tools, executeToolCall, getContext } = useDocxAgentTools({
 *   editorRef,
 *   author: 'Assistant',
 * });
 * ```
 *
 * @packageDocumentation
 * @public
 */
import { s as EditorRefLike, l as AgentToolResult, r as getToolSchemas, k as AgentToolDefinition, t as AgentContextSnapshot } from './server-29plFYqd.mjs';
export { u as getToolDisplayName } from './server-29plFYqd.mjs';
import * as react from 'react';
import { ReactNode, CSSProperties } from 'react';
import { A as AgentMessage, a as AgentToolCall } from './agent-types-C8RvQB7n.mjs';

/**
 * useAgentChat — React hook that wires agent tools to a live DocxEditor.
 *
 * @example
 * ```tsx
 * import { useAgentChat } from '@eigenpal/docx-editor-agents/react';
 *
 * const { executeToolCall, toolSchemas } = useAgentChat({ editorRef, author: 'Assistant' });
 *
 * // Pass toolSchemas to your AI provider, execute tool calls on the client
 * const result = executeToolCall('add_comment', { paragraphIndex: 3, text: 'Fix this.' });
 * ```
 */

interface UseAgentChatOptions {
    /** Reference to the DocxEditor (must match EditorRefLike interface). */
    editorRef: React.RefObject<EditorRefLike | null>;
    /** Default author name for comments and changes. Default: 'AI' */
    author?: string;
}
interface UseAgentChatReturn {
    /** Execute a tool call through the bridge. */
    executeToolCall: (toolName: string, input: Record<string, unknown>) => AgentToolResult;
    /** Tool schemas in OpenAI function calling format. Pass to your AI provider. */
    toolSchemas: ReturnType<typeof getToolSchemas>;
}
/**
 * Hook that creates an EditorBridge and provides tool execution.
 */
declare function useAgentChat(options: UseAgentChatOptions): UseAgentChatReturn;

/**
 * useDocxAgentTools — wires the toolkit to a live DocxEditor for BYO chat
 * frameworks.
 *
 * Returns three things consumers need to plug an agent into the editor:
 *
 *  - `tools` — schemas in OpenAI function-calling format. Pass to `streamText({ tools })`,
 *    OpenAI's `tools` field, Anthropic's `tools`, or any provider that accepts
 *    that shape.
 *  - `executeToolCall` — the executor. Hand to AI SDK's `onToolCall`, or call
 *    yourself when you wire up tool calls manually.
 *  - `getContext` — snapshot of `{selection, currentPage, paragraphCount}` for
 *    system-prompt injection. Pass through `prepareRequestBody` (AI SDK) or
 *    inline into your own request body so the agent always knows what the
 *    user is looking at without an extra tool round-trip.
 *
 * Custom tools merge with the built-ins via the `tools` option. Names collide
 * → consumer wins (your override replaces the built-in by name).
 *
 * @example
 * ```tsx
 * const { tools, executeToolCall, getContext } = useDocxAgentTools({
 *   editorRef,
 *   author: 'Assistant',
 *   tools: {
 *     fetch_clause: {
 *       name: 'fetch_clause',
 *       description: 'Fetch a clause template by name.',
 *       inputSchema: { type: 'object', properties: { name: { type: 'string' } } },
 *       handler: async (input) => ({ success: true, data: await fetchTemplate(input.name) }),
 *     },
 *   },
 * });
 * ```
 */

interface UseDocxAgentToolsOptions {
    /** Reference to the DocxEditor (must match EditorRefLike). */
    editorRef: React.RefObject<EditorRefLike | null>;
    /** Default author name for comments / tracked changes. Default: 'AI'. */
    author?: string;
    /**
     * Optional consumer-defined tools to merge with the built-ins. Keyed by
     * tool name. A tool with the same name as a built-in **replaces** it.
     * Pass a stable reference (memoized or module-level) to avoid rebuilding
     * the tool list on every render.
     */
    tools?: Record<string, AgentToolDefinition<any>>;
    /**
     * Allow-list of built-in tool names to expose. When provided, only the
     * named tools are returned (custom tools from `tools` always pass).
     * Useful for read-only or scope-restricted agents:
     *
     * @example include: ['read_document', 'find_text', 'add_comment']
     */
    include?: readonly string[];
    /**
     * Block-list of built-in tool names to hide. Applied after `include`.
     * Use for agents that should not write tracked changes:
     *
     * @example exclude: ['suggest_change', 'apply_formatting', 'set_paragraph_style']
     */
    exclude?: readonly string[];
}
interface UseDocxAgentToolsReturn {
    /** Tool schemas in OpenAI function calling format — pass to your AI provider. */
    tools: ReturnType<typeof getToolSchemas>;
    /** Execute a tool call by name. Pass to AI SDK's `onToolCall`. */
    executeToolCall: (name: string, args: Record<string, unknown>) => AgentToolResult;
    /** Snapshot of the user's current view for system-prompt injection. */
    getContext: () => AgentContextSnapshot;
}
declare function useDocxAgentTools(options: UseDocxAgentToolsOptions): UseDocxAgentToolsReturn;

interface AgentPanelProps {
    /** Header title. Defaults to `'Assistant'`. Pass a localised string for i18n. */
    title?: string;
    /** Header icon node. Defaults to a sparkle SVG. */
    icon?: ReactNode;
    /** Close button aria-label. Defaults to `'Close panel'`. */
    closeLabel?: string;
    /** Resize handle aria-label. Defaults to `'Resize agent panel'`. */
    resizeHandleLabel?: string;
    /** Controlled width in pixels. Omit for uncontrolled (internal state + localStorage). */
    width?: number;
    /** Default width when uncontrolled. */
    defaultWidth?: number;
    /** Min drag width. */
    minWidth?: number;
    /** Max drag width. */
    maxWidth?: number;
    /** Width change callback (drag end and intermediate). */
    onWidthChange?: (w: number) => void;
    /** Header close button click. Omit to hide the close button. */
    onClose?: () => void;
    /** Panel content. Render whatever you want — a chat, tabs, settings, anything. */
    children: ReactNode;
    /** Optional class on the outer wrapper. */
    className?: string;
    /**
     * When `true`, the panel collapses to zero width with an ease-out
     * transition (the children are still mounted so chat state survives
     * close/reopen). The DocxEditor wrapper passes this when the user
     * toggles the panel — kept off by default for standalone usage.
     */
    closed?: boolean;
}
declare function AgentPanel({ title, icon, closeLabel, resizeHandleLabel, width: controlledWidth, defaultWidth, minWidth, maxWidth, onWidthChange, onClose, children, className, closed, }: AgentPanelProps): react.JSX.Element;

interface AgentChatLogProps {
    messages: AgentMessage[];
    /** Render thinking dots at the bottom of the list. */
    loading?: boolean;
    /** Render an error bubble after the last message. */
    error?: string | null;
    /** Shown when there are no messages and not loading. */
    emptyState?: ReactNode;
    /** "Assistant is thinking" aria-label. Default English. */
    thinkingLabel?: string;
    /** "Working… N steps" — pass for i18n. Default English. */
    workingLabel?: (count: number) => string;
    /** "N steps" — pass for i18n. Default English. */
    summaryLabel?: (count: number) => string;
    /** "+ N earlier steps" — pass for i18n. Default English. */
    earlierLabel?: (count: number) => string;
    /** Auto-scroll to bottom on new messages / loading toggles. Default: true. */
    autoScroll?: boolean;
    /**
     * Map a tool name to a friendly label for the per-message timeline.
     * Pass `getToolDisplayName` from `@eigenpal/docx-editor-agents/react` to
     * use the toolkit's registry-aware labels (e.g. "Adding comment").
     */
    humanizeToolName?: (name: string) => string;
    /** Cap the tool-call timeline to this many recent rows. Default 3. */
    maxVisibleCalls?: number;
    className?: string;
    style?: CSSProperties;
}
interface AgentTimelineProps {
    /** Tool calls in chronological order. */
    toolCalls: AgentToolCall[];
    /**
     * Whether the parent assistant turn is still streaming. While true, the
     * timeline is forced expanded and shows a spinner; on false it
     * auto-collapses to an "N steps" summary unless the user expanded it.
     */
    streaming?: boolean;
    /**
     * Cap the number of rendered call rows — older entries collapse into a
     * "+N earlier steps" header. Default 3.
     */
    maxVisibleCalls?: number;
    /**
     * Map a tool name to a friendly label. Defaults to a sentence-case
     * conversion of the snake_case name. Pass `getToolDisplayName` from
     * `@eigenpal/docx-editor-agents/react` to use the toolkit's registry.
     */
    humanizeName?: (name: string) => string;
    /** "Working… N steps" — pass for i18n. Default English. */
    workingLabel?: (count: number) => string;
    /** "N steps" — pass for i18n. Default English. */
    summaryLabel?: (count: number) => string;
    /** "+ N earlier steps" — pass for i18n. Default English. */
    earlierLabel?: (count: number) => string;
}
/**
 * Collapsible timeline of an assistant turn's tool calls. Lives above the
 * assistant text bubble. Auto-collapses when the turn finishes; click the
 * summary row to re-expand.
 */
declare function AgentTimeline({ toolCalls, streaming, maxVisibleCalls, humanizeName, workingLabel, summaryLabel, earlierLabel, }: AgentTimelineProps): react.JSX.Element | null;
declare function AgentChatLog({ messages, loading, error, emptyState, autoScroll, humanizeToolName, maxVisibleCalls, className, style, thinkingLabel, workingLabel, summaryLabel, earlierLabel, }: AgentChatLogProps): react.JSX.Element;
interface AgentComposerProps {
    value: string;
    onChange: (next: string) => void;
    onSubmit: () => void;
    disabled?: boolean;
    placeholder?: string;
    /** Send-button aria-label. Default `'Send'`. */
    sendLabel?: string;
    /** Small text under the input — typically a scope reminder. */
    footnote?: ReactNode;
    className?: string;
}
declare function AgentComposer({ value, onChange, onSubmit, disabled, placeholder, sendLabel, footnote, className, }: AgentComposerProps): react.JSX.Element;
interface AgentSuggestionChipProps {
    label: string;
    onClick: () => void;
    disabled?: boolean;
}
declare function AgentSuggestionChip({ label, onClick, disabled }: AgentSuggestionChipProps): react.JSX.Element;

export { AgentChatLog, type AgentChatLogProps, AgentComposer, type AgentComposerProps, AgentContextSnapshot, AgentMessage, AgentPanel, type AgentPanelProps, AgentSuggestionChip, type AgentSuggestionChipProps, AgentTimeline, type AgentTimelineProps, AgentToolCall, AgentToolDefinition, AgentToolResult, EditorRefLike, type UseAgentChatOptions, type UseAgentChatReturn, type UseDocxAgentToolsOptions, type UseDocxAgentToolsReturn, useAgentChat, useDocxAgentTools };
