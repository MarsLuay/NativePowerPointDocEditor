/**
 * @eigenpal/docx-editor-agents/server
 *
 * Server entry for API routes, Node.js, serverless functions, and Workers.
 *
 * Import the toolkit here without pulling in React peer deps. Use this from
 * Next.js routes, a FastAPI bridge, a Cloudflare Worker, or any other
 * backend that streams an LLM call with tool definitions.
 *
 * @example
 * ```ts
 * import { getToolSchemas } from '@eigenpal/docx-editor-agents/server';
 * import { streamText, jsonSchema, convertToModelMessages } from 'ai';
 *
 * // `getToolSchemas()` returns OpenAI function-calling format. For Vercel
 * // AI SDK v5, adapt to `{ [name]: { description, inputSchema } }` once.
 * const tools = Object.fromEntries(
 *   getToolSchemas().map((s) => [
 *     s.function.name,
 *     { description: s.function.description, inputSchema: jsonSchema(s.function.parameters as never) },
 *   ])
 * );
 *
 * export async function POST(req: Request) {
 *   const { messages } = await req.json();
 *   const result = streamText({ model: 'openai/gpt-4o', messages: convertToModelMessages(messages), tools });
 *   return result.toUIMessageStreamResponse();
 * }
 * ```
 *
 * @packageDocumentation
 * @public
 */
export { t as AgentContextSnapshot, k as AgentToolDefinition, l as AgentToolResult, D as DocxReviewer, E as EditorBridge, S as SelectionInfo, p as createReviewerBridge, o as docxAgentTools, q as executeToolCall, u as getToolDisplayName, r as getToolSchemas } from './server-29plFYqd.mjs';
