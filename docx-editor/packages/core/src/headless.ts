/**
 * @npde/docx-editor-core/headless
 *
 * Headless aggregate for Node.js scripts, CLI tools, and server-side
 * processing. Prefer smaller subpaths (`./docx`, `./contentControls`,
 * `./utils`) for new code — they tree-shake better.
 *
 * Application-level OOXML edits in the Obsidian plugin live under plugin
 * `src/ai` + packers (`parseDocx` / `exportDocxBuffer` / rezip). This barrel
 * exposes the Document model and packers, not DocumentAgent.
 *
 * @example
 * ```ts
 * import { parseDocx, exportDocxBuffer, getBodyText } from '@npde/docx-editor-core/headless';
 *
 * const buffer = fs.readFileSync('input.docx');
 * const doc = await parseDocx(buffer);
 * console.log('Plain text:', getBodyText(doc.package.document));
 * const output = await exportDocxBuffer(doc);
 * fs.writeFileSync('output.docx', Buffer.from(output));
 * ```
 * @packageDocumentation
 * @public
 */

// ============================================================================
// VERSION
// ============================================================================

export const VERSION = '0.0.2';

// ============================================================================
// TEXT UTILITIES
// ============================================================================

export {
  getParagraphText,
  getRunText,
  getHyperlinkText,
  getTableText,
  getBodyText,
  countWords,
  countCharacters,
  getBodyWordCount,
  getBodyCharacterCount,
  getTextBefore,
  getTextAfter,
  getFormattingAtPosition,
  isPositionInHyperlink,
  getHyperlinkAtPosition,
  isHeadingStyle,
  parseHeadingLevel,
  hasImages,
  hasHyperlinks,
  hasTables,
  getParagraphs,
  getParagraphAtIndex,
  getBlockIndexForParagraph,
} from './contentControls/text-utils';

// ============================================================================
// CONTENT CONTROLS (SDT)
// ============================================================================

export {
  findContentControls,
  findContentControl,
  getContentControlText,
  setContentControlContent,
  removeContentControl,
  ContentControlNotFoundError,
  ContentControlLockedError,
  ContentControlTypeError,
  ContentControlBoundError,
  ContentControlKindError,
  type ContentControlFilter,
  type ContentControlInfo,
  type ContentControlLocation,
  type FindContentControlsOptions,
} from './contentControls/contentControls';
export {
  createContentControl,
  ContentControlCreateError,
  type CreateContentControlTarget,
  type NewContentControlProps,
} from './contentControls/createContentControl';
export {
  setContentControlValue,
  formatSdtDate,
  ContentControlValueError,
  type ContentControlValue,
} from './contentControls/contentControlValues';
export {
  addRepeatingSectionItem,
  removeRepeatingSectionItem,
  isRepeatingSection,
  isRepeatingSectionItem,
  RepeatingSectionError,
} from './contentControls/repeatingSection';

// ============================================================================
// PARSER / SERIALIZER
// ============================================================================

export { parseDocx } from './docx/parser';
export {
  serializeDocument as serializeDocx,
  serializeDocumentBody,
} from './docx/serializer/documentSerializer';
export { serializeSectionProperties } from './docx/serializer/sectionPropertiesSerializer';
export { repackDocx, createDocx, updateMultipleFiles } from './docx/rezip';
export { getDocumentWatermark, setDocumentWatermark } from './docx/watermarkApi';
export { attemptSelectiveSave } from './docx/selectiveSave';
export { buildPatchedDocumentXml, validatePatchSafety } from './docx/selectiveXmlPatch';

// ============================================================================
// TEMPLATE PROCESSING
// ============================================================================

export {
  processTemplate,
  processTemplateDetailed,
  processTemplateAsBlob,
  processTemplateAdvanced,
  getTemplateTags,
  validateTemplate,
  getMissingVariables,
  previewTemplate,
  createTemplateProcessor,
  type ProcessTemplateOptions,
  type ProcessTemplateResult,
  type TemplateError,
} from './utils/processTemplate';

// ============================================================================
// VARIABLE DETECTION
// ============================================================================

export {
  detectVariables,
  detectVariablesDetailed,
  detectVariablesInBody,
  detectVariablesInParagraph,
  extractVariablesFromText,
  hasTemplateVariables,
  isValidVariableName,
  sanitizeVariableName,
  formatVariable,
  parseVariable,
  replaceVariables,
  removeVariables,
  documentHasVariables,
  type VariableDetectionResult,
  type VariableOccurrence,
} from './utils/variableDetector';

// ============================================================================
// DOCUMENT CREATION
// ============================================================================

export {
  createEmptyDocument,
  createDocumentWithText,
  type CreateEmptyDocumentOptions,
} from './utils/createDocument';

// ============================================================================
// UTILITIES
// ============================================================================

export {
  twipsToPixels,
  pixelsToTwips,
  formatPx,
  emuToPixels,
  pointsToPixels,
  pointsToHalfPoints,
  halfPointsToPixels,
  pixelsToEmu,
  emuToTwips,
  twipsToEmu,
} from './utils/units';

export { mapHexToHighlightName } from './utils/highlightColors';

export {
  resolveColor,
  resolveHighlightColor,
  resolveShadingColor,
  parseColorString,
  createThemeColor,
  createRgbColor,
  darkenColor,
  lightenColor,
  blendColors,
  getContrastingColor,
  isBlack,
  isWhite,
  colorsEqual,
} from './utils/colorResolver';

// ============================================================================
// PLUGIN SYSTEM
// ============================================================================

export {
  pluginRegistry,
  PluginRegistry,
  registerPlugins,
  createPluginRegistrar,
  isZodSchema,
  type CorePlugin,
  type Plugin,
  type PluginCommand,
  type CommandHandler,
  type PluginCommandHandler,
  type CommandResult,
  type PluginOptions,
  type PluginRegistrationResult,
  type McpToolDefinition,
  type ToolDefinition,
  type McpToolHandler,
  type ToolHandler,
  type McpToolResult,
  type ToolResult,
  type McpToolContent,
  type McpToolContext,
  type McpToolAnnotations,
  type McpSession,
  type LoadedDocument,
  type JsonSchema,
  type ZodSchemaLike,
  type PluginEvent,
  type PluginEventListener,
} from './core-plugins';

// ============================================================================
// TYPES
// ============================================================================

// Document types
export type {
  Document,
  DocxPackage,
  DocumentBody,
  BlockContent,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  TextContent,
  Table,
  TableRow,
  TableCell,
  Image,
  Hyperlink,
  Theme,
  Style,
  StyleDefinitions,
  TextFormatting,
  ParagraphFormatting,
  SectionProperties,
  Footnote,
  Endnote,
  ListLevel,
  NumberingDefinitions,
  Relationship,
  // Track changes & comments
  Comment,
  CommentRangeStart,
  CommentRangeEnd,
  TrackedChangeInfo,
  TrackedRunChange,
  Insertion,
  Deletion,
  MoveFrom,
  MoveTo,
} from './types/document';

// Agent API types
export type {
  AIAction,
  AIActionRequest,
  AgentResponse,
  AgentContext,
  SelectionContext,
  Range,
  Position,
  ParagraphContext,
  ParagraphOutline,
  SectionInfo,
  StyleInfo,
  SuggestedAction,
  AgentCommand,
  InsertTextCommand,
  ReplaceTextCommand,
  DeleteTextCommand,
  FormatTextCommand,
  FormatParagraphCommand,
  InsertTableCommand,
  InsertImageCommand,
  InsertHyperlinkCommand,
  SetVariableCommand,
  ApplyStyleCommand,
  ApplyVariablesCommand,
} from './types/agentApi';

// API functions
export {
  createCollapsedRange,
  createRange,
  isPositionInRange,
  comparePositions,
  getActionLabel,
  getActionDescription,
  createCommand,
  DEFAULT_AI_ACTIONS,
} from './types/agentApi';
