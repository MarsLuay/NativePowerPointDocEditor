/**
 * Agent API Types
 *
 * TypeScript interfaces for the agent API:
 * - Position and Range types
 * - Command types for document manipulation
 * - Context types for AI agents
 * @packageDocumentation
 * @public
 */
import { T as TextFormatting, P as ParagraphFormatting } from '../formatting-JhqWT_XM.js';
import '../colors-C3vA7HUU.js';

/**
 * Agent API Types
 *
 * TypeScript interfaces for the agent API:
 * - Position and Range types
 * - Command types for document manipulation
 * - Context types for AI agents
 * @packageDocumentation
 * @public
 */

/**
 * Position within a document
 */
interface Position {
    /** Index of the paragraph (0-indexed) */
    paragraphIndex: number;
    /** Offset within the paragraph in characters */
    offset: number;
    /** Optional: Content index within paragraph (run, hyperlink, etc.) */
    contentIndex?: number;
    /** Optional: Section index */
    sectionIndex?: number;
}
/**
 * Range within a document
 */
interface Range {
    /** Start position */
    start: Position;
    /** End position */
    end: Position;
    /** Whether the range is collapsed (cursor position) */
    collapsed?: boolean;
}
/**
 * Create a collapsed range (cursor) at a position
 */
declare function createCollapsedRange(position: Position): Range;
/**
 * Create a range from two positions
 */
declare function createRange(start: Position, end: Position): Range;
/**
 * Check if a position is within a range
 */
declare function isPositionInRange(position: Position, range: Range): boolean;
/**
 * Compare two positions
 * Returns: -1 if a < b, 0 if equal, 1 if a > b
 */
declare function comparePositions(a: Position, b: Position): -1 | 0 | 1;
/**
 * Base command interface
 */
interface BaseCommand {
    /** Command type */
    type: string;
    /** Unique command ID (for undo tracking) */
    id?: string;
}
/**
 * Insert text at a position
 */
interface InsertTextCommand extends BaseCommand {
    type: 'insertText';
    /** Position to insert at */
    position: Position;
    /** Text to insert */
    text: string;
    /** Optional formatting for the inserted text */
    formatting?: TextFormatting;
}
/**
 * Replace text in a range
 */
interface ReplaceTextCommand extends BaseCommand {
    type: 'replaceText';
    /** Range to replace */
    range: Range;
    /** Replacement text */
    text: string;
    /** Optional formatting for the new text */
    formatting?: TextFormatting;
}
/**
 * Delete text in a range
 */
interface DeleteTextCommand extends BaseCommand {
    type: 'deleteText';
    /** Range to delete */
    range: Range;
}
/**
 * Apply formatting to a range
 */
interface FormatTextCommand extends BaseCommand {
    type: 'formatText';
    /** Range to format */
    range: Range;
    /** Formatting to apply */
    formatting: Partial<TextFormatting>;
}
/**
 * Apply paragraph formatting
 */
interface FormatParagraphCommand extends BaseCommand {
    type: 'formatParagraph';
    /** Paragraph index */
    paragraphIndex: number;
    /** Formatting to apply */
    formatting: Partial<ParagraphFormatting>;
}
/**
 * Apply a named style to a paragraph
 */
interface ApplyStyleCommand extends BaseCommand {
    type: 'applyStyle';
    /** Paragraph index */
    paragraphIndex: number;
    /** Style ID to apply */
    styleId: string;
}
/**
 * Insert a table at a position
 */
interface InsertTableCommand extends BaseCommand {
    type: 'insertTable';
    /** Position to insert at */
    position: Position;
    /** Number of rows */
    rows: number;
    /** Number of columns */
    columns: number;
    /** Optional table data */
    data?: string[][];
    /** Optional header row */
    hasHeader?: boolean;
}
/**
 * Insert an image at a position
 */
interface InsertImageCommand extends BaseCommand {
    type: 'insertImage';
    /** Position to insert at */
    position: Position;
    /** Image source (base64 or URL) */
    src: string;
    /** Image width in pixels */
    width?: number;
    /** Image height in pixels */
    height?: number;
    /** Alt text */
    alt?: string;
}
/**
 * Insert a hyperlink at a range
 */
interface InsertHyperlinkCommand extends BaseCommand {
    type: 'insertHyperlink';
    /** Range to make into a hyperlink */
    range: Range;
    /** URL of the hyperlink */
    url: string;
    /** Display text (replaces range text if provided) */
    displayText?: string;
    /** Tooltip */
    tooltip?: string;
}
/**
 * Remove a hyperlink but keep the text
 */
interface RemoveHyperlinkCommand extends BaseCommand {
    type: 'removeHyperlink';
    /** Range containing the hyperlink */
    range: Range;
}
/**
 * Insert a paragraph break
 */
interface InsertParagraphBreakCommand extends BaseCommand {
    type: 'insertParagraphBreak';
    /** Position to break at */
    position: Position;
}
/**
 * Merge paragraphs
 */
interface MergeParagraphsCommand extends BaseCommand {
    type: 'mergeParagraphs';
    /** First paragraph index */
    paragraphIndex: number;
    /** Number of paragraphs to merge with */
    count: number;
}
/**
 * Split a paragraph
 */
interface SplitParagraphCommand extends BaseCommand {
    type: 'splitParagraph';
    /** Position to split at */
    position: Position;
}
/**
 * Set template variable value
 */
interface SetVariableCommand extends BaseCommand {
    type: 'setVariable';
    /** Variable name */
    name: string;
    /** Variable value */
    value: string;
}
/**
 * Apply all template variables
 */
interface ApplyVariablesCommand extends BaseCommand {
    type: 'applyVariables';
    /** Variable values */
    values: Record<string, string>;
}
/**
 * Union of all command types
 */
type AgentCommand = InsertTextCommand | ReplaceTextCommand | DeleteTextCommand | FormatTextCommand | FormatParagraphCommand | ApplyStyleCommand | InsertTableCommand | InsertImageCommand | InsertHyperlinkCommand | RemoveHyperlinkCommand | InsertParagraphBreakCommand | MergeParagraphsCommand | SplitParagraphCommand | SetVariableCommand | ApplyVariablesCommand;
/**
 * Get command type
 */
type CommandType = AgentCommand['type'];
/**
 * Document context for AI agents
 */
interface AgentContext {
    /** Total paragraph count */
    paragraphCount: number;
    /** Total word count (approximate) */
    wordCount: number;
    /** Total character count */
    characterCount: number;
    /** Detected template variables */
    variables: string[];
    /** Variable count */
    variableCount: number;
    /** Available styles */
    availableStyles: StyleInfo[];
    /** Content outline (first N chars per paragraph) */
    outline: ParagraphOutline[];
    /** Document sections info */
    sections: SectionInfo[];
    /** Has tables */
    hasTables: boolean;
    /** Has images */
    hasImages: boolean;
    /** Has hyperlinks */
    hasHyperlinks: boolean;
    /** Document language */
    language?: string;
}
/**
 * Style information for context
 */
interface StyleInfo {
    /** Style ID */
    id: string;
    /** Display name */
    name: string;
    /** Style type */
    type: 'paragraph' | 'character' | 'table';
    /** Is built-in style */
    builtIn?: boolean;
}
/**
 * Paragraph outline for context
 */
interface ParagraphOutline {
    /** Paragraph index */
    index: number;
    /** First N characters */
    preview: string;
    /** Paragraph style */
    style?: string;
    /** Is heading */
    isHeading?: boolean;
    /** Heading level (1-9) */
    headingLevel?: number;
    /** Is list item */
    isListItem?: boolean;
    /** Is empty paragraph */
    isEmpty?: boolean;
}
/**
 * Section information
 */
interface SectionInfo {
    /** Section index */
    index: number;
    /** Number of paragraphs */
    paragraphCount: number;
    /** Page size */
    pageSize?: {
        width: number;
        height: number;
    };
    /** Is landscape */
    isLandscape?: boolean;
    /** Has header */
    hasHeader?: boolean;
    /** Has footer */
    hasFooter?: boolean;
}
/**
 * Context about the current selection
 */
interface SelectionContext {
    /** Selected text */
    selectedText: string;
    /** Selection range */
    range: Range;
    /** Current formatting of selection */
    formatting: Partial<TextFormatting>;
    /** Current paragraph formatting */
    paragraphFormatting: Partial<ParagraphFormatting>;
    /** Text before selection (context) */
    textBefore: string;
    /** Text after selection (context) */
    textAfter: string;
    /** Paragraph containing selection */
    paragraph: ParagraphContext;
    /** Is selection within a table */
    inTable?: boolean;
    /** Is selection within a hyperlink */
    inHyperlink?: boolean;
    /** Suggested actions based on selection */
    suggestedActions?: SuggestedAction[];
}
/**
 * Paragraph context for selection
 */
interface ParagraphContext {
    /** Paragraph index */
    index: number;
    /** Full paragraph text */
    fullText: string;
    /** Paragraph style */
    style?: string;
    /** Word count */
    wordCount: number;
}
/**
 * Suggested action for context menu
 */
interface SuggestedAction {
    /** Action ID */
    id: string;
    /** Display label */
    label: string;
    /** Description */
    description?: string;
    /** Icon name */
    icon?: string;
    /** Priority (higher = more prominent) */
    priority?: number;
}
/**
 * Response from an agent action
 */
interface AgentResponse {
    /** Success status */
    success: boolean;
    /** New text to insert (for rewrite/expand/etc.) */
    newText?: string;
    /** New formatted content */
    newContent?: AgentContent[];
    /** Commands to execute */
    commands?: AgentCommand[];
    /** Error message if failed */
    error?: string;
    /** Warning messages */
    warnings?: string[];
    /** Metadata about the response */
    metadata?: Record<string, unknown>;
}
/**
 * Content block in agent response
 */
interface AgentContent {
    /** Content type */
    type: 'text' | 'paragraph' | 'table' | 'image';
    /** Text content */
    text?: string;
    /** Formatting */
    formatting?: Partial<TextFormatting>;
    /** Paragraph formatting */
    paragraphFormatting?: Partial<ParagraphFormatting>;
    /** Table data (for table type) */
    tableData?: string[][];
    /** Image src (for image type) */
    imageSrc?: string;
}
/**
 * AI action types for context menu
 */
type AIAction = 'askAI' | 'rewrite' | 'expand' | 'summarize' | 'translate' | 'explain' | 'fixGrammar' | 'makeFormal' | 'makeCasual' | 'custom';
/**
 * AI action request
 */
interface AIActionRequest {
    /** Action type */
    action: AIAction;
    /** Selection context */
    context: SelectionContext;
    /** Custom prompt (for 'custom' action) */
    customPrompt?: string;
    /** Target language (for 'translate' action) */
    targetLanguage?: string;
    /** Additional options */
    options?: Record<string, unknown>;
}
/**
 * Get action label
 */
declare function getActionLabel(action: AIAction): string;
/**
 * Get action description
 */
declare function getActionDescription(action: AIAction): string;
/**
 * Default AI actions for context menu
 */
declare const DEFAULT_AI_ACTIONS: AIAction[];
/**
 * Command handler function type
 */
type CommandHandler<T extends AgentCommand = AgentCommand> = (command: T) => Promise<boolean>;
/**
 * AI request handler function type
 */
type AIRequestHandler = (request: AIActionRequest) => Promise<AgentResponse>;
/**
 * Create a command with generated ID
 */
declare function createCommand<T extends AgentCommand>(command: Omit<T, 'id'>): T;
declare const _default: {
    createCollapsedRange: typeof createCollapsedRange;
    createRange: typeof createRange;
    isPositionInRange: typeof isPositionInRange;
    comparePositions: typeof comparePositions;
    getActionLabel: typeof getActionLabel;
    getActionDescription: typeof getActionDescription;
    createCommand: typeof createCommand;
    DEFAULT_AI_ACTIONS: AIAction[];
};

export { type AIAction, type AIActionRequest, type AIRequestHandler, type AgentCommand, type AgentContent, type AgentContext, type AgentResponse, type ApplyStyleCommand, type ApplyVariablesCommand, type BaseCommand, type CommandHandler, type CommandType, DEFAULT_AI_ACTIONS, type DeleteTextCommand, type FormatParagraphCommand, type FormatTextCommand, type InsertHyperlinkCommand, type InsertImageCommand, type InsertParagraphBreakCommand, type InsertTableCommand, type InsertTextCommand, type MergeParagraphsCommand, type ParagraphContext, type ParagraphOutline, type Position, type Range, type RemoveHyperlinkCommand, type ReplaceTextCommand, type SectionInfo, type SelectionContext, type SetVariableCommand, type SplitParagraphCommand, type StyleInfo, type SuggestedAction, comparePositions, createCollapsedRange, createCommand, createRange, _default as default, getActionDescription, getActionLabel, isPositionInRange };
