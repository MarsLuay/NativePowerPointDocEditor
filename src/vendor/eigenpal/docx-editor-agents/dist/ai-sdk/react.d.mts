/**
 * @eigenpal/docx-editor-agents/ai-sdk/react
 *
 * Vercel AI SDK adapter (React side). Opt-in.
 *
 * Use this if you're driving the chat with `useChat` from `@ai-sdk/react`.
 * The library's `<AgentChatLog>` consumes a flat `AgentMessage[]` shape;
 * AI SDK's `useChat` produces `UIMessage[]` with structured `parts`.
 * `toAgentMessages()` is the bridge.
 *
 * @example
 * ```tsx
 * const chat = useChat({ ... });
 * const messages = useMemo(
 *   () => toAgentMessages(chat.messages, chat.status),
 *   [chat.messages, chat.status]
 * );
 * return <AgentChatLog messages={messages} />;
 * ```
 *
 * @packageDocumentation
 * @public
 */
import { A as AgentMessage } from '../agent-types-C8RvQB7n.mjs';
export { a as AgentToolCall } from '../agent-types-C8RvQB7n.mjs';

/**
 * Framework-agnostic Vercel AI SDK adapter logic. The React and Vue
 * subpaths re-export from here so consumers don't have to import a
 * cross-framework path.
 */

/** Minimal structural shape of a Vercel AI SDK `UIMessage`. */
interface AiSdkUIMessage {
    id: string;
    role: 'user' | 'assistant' | 'system';
    parts?: ReadonlyArray<{
        type: string;
        text?: string;
        toolCallId?: string;
        state?: string;
        input?: unknown;
        output?: unknown;
        errorText?: string;
    }>;
}
/**
 * Adapt AI SDK's `UIMessage[]` (from `useChat`) to the `AgentMessage[]`
 * shape `<AgentChatLog>` consumes.
 *
 * @param uiMessages - the `messages` array from `useChat`
 * @param status - the `status` from `useChat`. The last assistant
 *   message is marked `streaming` while the chat is still in flight.
 */
declare function toAgentMessages(uiMessages: ReadonlyArray<AiSdkUIMessage>, status: string): AgentMessage[];

export { AgentMessage, type AiSdkUIMessage, toAgentMessages };
