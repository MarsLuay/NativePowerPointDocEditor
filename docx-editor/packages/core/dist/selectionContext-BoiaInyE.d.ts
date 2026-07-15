import { Document } from './types/document.js';
import { StyleInfo, AgentContext, Position, Range, AgentCommand, SelectionContext } from './types/agentApi.js';
import { S as SelectiveSaveOptions } from './selectiveSave-CWaPEv0B.js';
import { D as DocxInput } from './docxInput-DTbCa48g.js';
import { T as TextFormatting, P as ParagraphFormatting } from './formatting-JhqWT_XM.js';

/**
 * DocumentAgent - High-level fluent API for programmatic document manipulation
 *
 * Provides a convenient interface for:
 * - Reading document content and metadata
 * - Editing text with formatting
 * - Inserting tables, images, and hyperlinks
 * - Managing template variables
 * - Exporting to DOCX buffer
 *
 * All operations are immutable - they return a new DocumentAgent instance
 * or don't modify the underlying document.
 */

/**
 * Options for inserting text
 */
interface InsertTextOptions {
    /** Text formatting */
    formatting?: TextFormatting;
}
/**
 * Options for inserting table
 */
interface InsertTableOptions {
    /** Table data (2D array of strings) */
    data?: string[][];
    /** Whether first row is a header */
    hasHeader?: boolean;
}
/**
 * Options for inserting image
 */
interface InsertImageOptions {
    /** Image width in pixels */
    width?: number;
    /** Image height in pixels */
    height?: number;
    /** Alt text for accessibility */
    alt?: string;
}
/**
 * Options for inserting hyperlink
 */
interface InsertHyperlinkOptions {
    /** Display text (overrides selected text) */
    displayText?: string;
    /** Tooltip on hover */
    tooltip?: string;
}
/**
 * Formatted text segment
 */
interface FormattedTextSegment {
    /** Text content */
    text: string;
    /** Applied formatting */
    formatting?: TextFormatting;
    /** Is part of a hyperlink */
    isHyperlink?: boolean;
    /** Hyperlink URL if applicable */
    hyperlinkUrl?: string;
}
/**
 * DocumentAgent provides a fluent API for document manipulation
 *
 * @example
 * ```ts
 * const agent = new DocumentAgent(buffer);
 *
 * // Read operations
 * const text = agent.getText();
 * const wordCount = agent.getWordCount();
 * const variables = agent.getVariables();
 *
 * // Write operations (returns new agent)
 * const newAgent = agent
 *   .insertText({ paragraphIndex: 0, offset: 0 }, 'Hello ', { formatting: { bold: true } })
 *   .applyStyle({ paragraphIndex: 0, offset: 0 }, { paragraphIndex: 0, offset: 5 }, 'Heading1');
 *
 * // Export
 * const newBuffer = await newAgent.toBuffer();
 * ```
 */
declare class DocumentAgent {
    private _document;
    private _pendingVariables;
    /**
     * Create a new DocumentAgent
     *
     * @param source - Document object or ArrayBuffer to parse
     */
    constructor(source: Document | ArrayBuffer);
    /**
     * Create a DocumentAgent from a DOCX buffer (async)
     *
     * @param buffer - DOCX file as ArrayBuffer, Uint8Array, Blob, or File
     * @returns Promise resolving to DocumentAgent
     */
    static fromBuffer(buffer: DocxInput): Promise<DocumentAgent>;
    /**
     * Create a DocumentAgent from a Document object
     *
     * @param document - Parsed Document
     * @returns DocumentAgent
     */
    static fromDocument(document: Document): DocumentAgent;
    /**
     * Get the underlying document
     */
    getDocument(): Document;
    /**
     * Get plain text content of the document
     *
     * @returns All document text concatenated
     */
    getText(): string;
    /**
     * Get formatted text segments
     *
     * @returns Array of text segments with formatting info
     */
    getFormattedText(): FormattedTextSegment[];
    /**
     * Get detected template variables
     *
     * @returns Array of variable names (without braces)
     */
    getVariables(): string[];
    /**
     * Get available styles from the document
     *
     * @returns Array of style info
     */
    getStyles(): StyleInfo[];
    /**
     * Get approximate page count
     *
     * Note: This is an estimate based on content length.
     * Actual page count requires full layout computation.
     *
     * @returns Estimated page count
     */
    getPageCount(): number;
    /**
     * Get word count
     *
     * @returns Number of words in the document
     */
    getWordCount(): number;
    /**
     * Get character count
     *
     * @param includeSpaces - Whether to include whitespace
     * @returns Number of characters
     */
    getCharacterCount(includeSpaces?: boolean): number;
    /**
     * Get paragraph count
     *
     * @returns Number of paragraphs
     */
    getParagraphCount(): number;
    /**
     * Get table count
     *
     * @returns Number of tables
     */
    getTableCount(): number;
    /**
     * Get document context for AI agents
     *
     * @param outlineMaxChars - Max characters per paragraph in outline
     * @returns Agent context
     */
    getAgentContext(outlineMaxChars?: number): AgentContext;
    /**
     * Insert text at a position
     *
     * @param position - Where to insert
     * @param text - Text to insert
     * @param options - Insert options
     * @returns New DocumentAgent with text inserted
     */
    insertText(position: Position, text: string, options?: InsertTextOptions): DocumentAgent;
    /**
     * Replace text in a range
     *
     * @param range - Range to replace
     * @param text - Replacement text
     * @param options - Replace options
     * @returns New DocumentAgent with text replaced
     */
    replaceRange(range: Range, text: string, options?: InsertTextOptions): DocumentAgent;
    /**
     * Delete text in a range
     *
     * @param range - Range to delete
     * @returns New DocumentAgent with text deleted
     */
    deleteRange(range: Range): DocumentAgent;
    /**
     * Apply text formatting to a range
     *
     * @param range - Range to format
     * @param formatting - Formatting to apply
     * @returns New DocumentAgent with formatting applied
     */
    applyFormatting(range: Range, formatting: Partial<TextFormatting>): DocumentAgent;
    /**
     * Apply a named style to a paragraph
     *
     * @param paragraphIndex - Index of the paragraph
     * @param styleId - Style ID to apply
     * @returns New DocumentAgent with style applied
     */
    applyStyle(paragraphIndex: number, styleId: string): DocumentAgent;
    /**
     * Apply paragraph formatting
     *
     * @param paragraphIndex - Index of the paragraph
     * @param formatting - Formatting to apply
     * @returns New DocumentAgent with formatting applied
     */
    applyParagraphFormatting(paragraphIndex: number, formatting: Partial<ParagraphFormatting>): DocumentAgent;
    /**
     * Insert a table at a position
     *
     * @param position - Where to insert the table
     * @param rows - Number of rows
     * @param cols - Number of columns
     * @param options - Table options
     * @returns New DocumentAgent with table inserted
     */
    insertTable(position: Position, rows: number, cols: number, options?: InsertTableOptions): DocumentAgent;
    /**
     * Insert an image at a position
     *
     * @param position - Where to insert the image
     * @param src - Image source (base64 data URL or URL)
     * @param options - Image options
     * @returns New DocumentAgent with image inserted
     */
    insertImage(position: Position, src: string, options?: InsertImageOptions): DocumentAgent;
    /**
     * Insert a hyperlink
     *
     * @param range - Range to make into a hyperlink
     * @param url - URL of the hyperlink
     * @param options - Hyperlink options
     * @returns New DocumentAgent with hyperlink inserted
     */
    insertHyperlink(range: Range, url: string, options?: InsertHyperlinkOptions): DocumentAgent;
    /**
     * Remove a hyperlink but keep the text
     *
     * @param range - Range containing the hyperlink
     * @returns New DocumentAgent with hyperlink removed
     */
    removeHyperlink(range: Range): DocumentAgent;
    /**
     * Insert a paragraph break
     *
     * @param position - Where to break the paragraph
     * @returns New DocumentAgent with paragraph broken
     */
    insertParagraphBreak(position: Position): DocumentAgent;
    /**
     * Merge consecutive paragraphs
     *
     * @param startParagraphIndex - First paragraph index
     * @param count - Number of paragraphs to merge with the first
     * @returns New DocumentAgent with paragraphs merged
     */
    mergeParagraphs(startParagraphIndex: number, count: number): DocumentAgent;
    /**
     * Set a template variable value
     *
     * Note: Variables are not applied until `applyVariables()` is called
     *
     * @param name - Variable name (without braces)
     * @param value - Variable value
     * @returns This DocumentAgent (for chaining)
     */
    setVariable(name: string, value: string): DocumentAgent;
    /**
     * Set multiple template variables
     *
     * @param variables - Map of variable names to values
     * @returns This DocumentAgent (for chaining)
     */
    setVariables(variables: Record<string, string>): DocumentAgent;
    /**
     * Get pending variable values
     *
     * @returns Map of pending variable values
     */
    getPendingVariables(): Record<string, string>;
    /**
     * Clear pending variables
     *
     * @returns This DocumentAgent (for chaining)
     */
    clearPendingVariables(): DocumentAgent;
    /**
     * Apply all pending template variables
     *
     * Uses docxtemplater to substitute variables while preserving formatting.
     *
     * @param variables - Optional additional variables (merged with pending)
     * @returns New DocumentAgent with variables applied
     */
    applyVariables(variables?: Record<string, string>): Promise<DocumentAgent>;
    /**
     * Export document to DOCX ArrayBuffer
     *
     * @returns Promise resolving to DOCX file as ArrayBuffer
     */
    toBuffer(options?: {
        selective?: SelectiveSaveOptions;
    }): Promise<ArrayBuffer>;
    /**
     * Export document to Blob
     *
     * @param mimeType - MIME type for the blob
     * @returns Promise resolving to DOCX file as Blob
     */
    toBlob(mimeType?: string): Promise<Blob>;
    /**
     * Execute multiple commands in sequence
     *
     * @param commands - Commands to execute
     * @returns New DocumentAgent with all commands applied
     */
    executeCommands(commands: AgentCommand[]): DocumentAgent;
    /**
     * Execute a single command and return new agent
     */
    private _executeCommand;
    /**
     * Extract formatted text segments from a paragraph
     */
    private _extractParagraphSegments;
}
/**
 * Create a DocumentAgent from a DOCX buffer
 *
 * @param buffer - DOCX file as ArrayBuffer
 * @returns Promise resolving to DocumentAgent
 */
declare function createAgent(buffer: ArrayBuffer): Promise<DocumentAgent>;
/**
 * Create a DocumentAgent from a parsed Document
 *
 * @param document - Parsed Document
 * @returns DocumentAgent
 */
declare function createAgentFromDocument(document: Document): DocumentAgent;

/**
 * Command Executor
 *
 * Executes agent commands on a document immutably:
 * - Handles all command types from AgentCommand
 * - Preserves surrounding formatting
 * - Returns new document (immutable updates)
 *
 * Per-command handlers live under `executor/` by domain (text edits,
 * paragraph splits/merges, structural inserts, template variables).
 * Shared helpers (clone, lookup, text-range edits) live in
 * `executor/helpers.ts`.
 */

/**
 * Execute an agent command on a document
 * Returns a new document with the command applied (immutable)
 *
 * Dispatch order:
 * 1. Try plugin handlers first (allows plugins to override built-in commands)
 * 2. Fall back to built-in handlers
 *
 * @param doc - The document to modify
 * @param command - The command to execute
 * @returns New document with command applied
 */
declare function executeCommand(doc: Document, command: AgentCommand): Document;
/**
 * Execute multiple commands in sequence
 *
 * @param doc - The document to modify
 * @param commands - Commands to execute in order
 * @returns New document with all commands applied
 */
declare function executeCommands(doc: Document, commands: AgentCommand[]): Document;

/**
 * Agent Context Builder
 *
 * Generates context objects for AI/LLM consumption from DOCX documents.
 * The context provides a structured summary of the document that can be
 * used by AI agents to understand the document structure and content.
 *
 * All outputs are JSON serializable for easy transmission to AI backends.
 */

/**
 * Options for building agent context
 */
interface AgentContextOptions {
    /** Maximum characters per paragraph in outline (default: 100) */
    outlineMaxChars?: number;
    /** Maximum paragraphs to include in outline (default: 50) */
    maxOutlineParagraphs?: number;
    /** Include table content in context (default: false) */
    includeTableContent?: boolean;
    /** Include detailed formatting info (default: false) */
    includeFormatting?: boolean;
}
/**
 * Options for building selection context
 */
interface SelectionContextOptions$1 {
    /** Characters of context before/after selection (default: 200) */
    contextChars?: number;
    /** Include suggested actions (default: true) */
    includeSuggestions?: boolean;
}
/**
 * Build agent context from a document
 *
 * @param doc - The parsed document
 * @param options - Context building options
 * @returns AgentContext object (JSON serializable)
 */
declare function getAgentContext(doc: Document, options?: AgentContextOptions): AgentContext;
/**
 * Build selection context for AI operations
 *
 * @param doc - The parsed document
 * @param range - The selected range
 * @param options - Selection context options
 * @returns SelectionContext object (JSON serializable)
 */
declare function buildSelectionContext$1(doc: Document, range: Range, options?: SelectionContextOptions$1): SelectionContext;
/**
 * Get a simple document summary for quick context
 *
 * @param doc - The parsed document
 * @returns Summary string
 */
declare function getDocumentSummary(doc: Document): string;

/**
 * Selection Context Builder
 *
 * Builds rich context objects from document selections for AI operations.
 * Includes selected text, formatting, surrounding context, and suggested actions.
 */

/**
 * Options for building selection context
 */
interface SelectionContextOptions {
    /** Characters of context before selection (default: 200) */
    contextCharsBefore?: number;
    /** Characters of context after selection (default: 200) */
    contextCharsAfter?: number;
    /** Include suggested actions (default: true) */
    includeSuggestions?: boolean;
    /** Include document summary (default: true) */
    includeDocumentSummary?: boolean;
    /** Maximum suggested actions (default: 8) */
    maxSuggestions?: number;
}
/**
 * Extended selection context with additional details
 */
interface ExtendedSelectionContext extends SelectionContext {
    /** Document summary for additional context */
    documentSummary?: string;
    /** Selection word count */
    wordCount?: number;
    /** Selection character count */
    characterCount?: number;
    /** Is selection multi-paragraph */
    isMultiParagraph?: boolean;
    /** Selected paragraph indices */
    paragraphIndices?: number[];
    /** Language detection hint */
    detectedLanguage?: string;
    /** Content type hints */
    contentType?: 'prose' | 'list' | 'heading' | 'table' | 'mixed';
}
/**
 * Selection formatting summary
 */
interface FormattingSummary {
    /** Predominant formatting */
    predominant: Partial<TextFormatting>;
    /** Is formatting consistent across selection */
    isConsistent: boolean;
    /** All formatting found */
    allFormatting: Partial<TextFormatting>[];
}
/**
 * Build selection context for AI operations
 *
 * @param doc - The parsed document
 * @param range - The selected range
 * @param options - Selection context options
 * @returns SelectionContext object
 */
declare function buildSelectionContext(doc: Document, range: Range, options?: SelectionContextOptions): SelectionContext;
/**
 * Build extended selection context with additional details
 *
 * @param doc - The parsed document
 * @param range - The selected range
 * @param options - Selection context options
 * @returns ExtendedSelectionContext object
 */
declare function buildExtendedSelectionContext(doc: Document, range: Range, options?: SelectionContextOptions): ExtendedSelectionContext;
/**
 * Get formatting summary for a selection
 *
 * @param doc - The parsed document
 * @param range - The selected range
 * @returns FormattingSummary object
 */
declare function getSelectionFormattingSummary(doc: Document, range: Range): FormattingSummary;

export { type AgentContextOptions as A, DocumentAgent as D, type ExtendedSelectionContext as E, type FormattedTextSegment as F, type InsertHyperlinkOptions as I, type SelectionContextOptions$1 as S, type FormattingSummary as a, type InsertImageOptions as b, type InsertTableOptions as c, type InsertTextOptions as d, type SelectionContextOptions as e, buildExtendedSelectionContext as f, buildSelectionContext as g, buildSelectionContext$1 as h, createAgent as i, createAgentFromDocument as j, executeCommand as k, executeCommands as l, getAgentContext as m, getDocumentSummary as n, getSelectionFormattingSummary as o };
