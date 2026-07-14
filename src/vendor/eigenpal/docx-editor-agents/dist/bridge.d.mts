/**
 * @eigenpal/docx-editor-agents/bridge
 *
 * Editor bridge that connects agent tools to a live `DocxEditor` instance.
 * Framework-agnostic. The React adapter lives in
 * `@eigenpal/docx-editor-agents/react`.
 *
 * @example
 * ```ts
 * import { createEditorBridge } from '@eigenpal/docx-editor-agents/bridge';
 * const bridge = createEditorBridge(editorRef, 'Assistant');
 * bridge.addComment({ paragraphIndex: 3, text: 'Fix this.' });
 * ```
 *
 * @packageDocumentation
 * @public
 */
export { k as AgentToolDefinition, l as AgentToolResult, i as ContentChangeEvent, E as EditorBridge, s as EditorRefLike, v as ParagraphHighlightOptions, h as ScrollToParaIdOptions, j as SelectionChangeEvent, o as agentTools, w as createEditorBridge, p as createReviewerBridge, q as executeToolCall, r as getToolSchemas } from './server-29plFYqd.mjs';
