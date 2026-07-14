import { MaybeRef, Ref } from 'vue';
import { EditorRefLike } from '../../bridge';
import { getToolSchemas, AgentToolResult } from '../../tools';
export interface UseAgentBridgeOptions {
    /** Vue ref pointing at the DocxEditor instance (must match `EditorRefLike`). */
    editorRef: Ref<EditorRefLike | null | undefined>;
    /**
     * Default author for comments and tracked changes. Accepts a plain
     * string or a `Ref<string>`/computed — bridge rebuilds when the value
     * changes (matches React's `useMemo([editorRef, author])` shape).
     * Defaults to `'AI'`.
     */
    author?: MaybeRef<string>;
}
export interface UseAgentBridgeReturn {
    executeToolCall: (toolName: string, input: Record<string, unknown>) => AgentToolResult;
    toolSchemas: ReturnType<typeof getToolSchemas>;
}
export declare function useAgentBridge(options: UseAgentBridgeOptions): UseAgentBridgeReturn;
