/**
 * @eigenpal/docx-editor-core/headless
 *
 * Headless aggregate for Node.js scripts, CLI tools, and server-side
 * processing. Same surface as the default `.` entry, named to make the
 * "no DOM" intent explicit. Prefer the smaller subpaths (`./docx`,
 * `./agent`, `./utils`, etc.) for new code — they tree-shake better.
 *
 * @example
 * ```ts
 * import { DocumentAgent, parseDocx, pluginRegistry } from '@eigenpal/docx-editor-core/headless';
 * import { docxtemplaterPlugin } from '@eigenpal/docx-editor-core/core-plugins';
 *
 * // Register plugins
 * pluginRegistry.register(docxtemplaterPlugin);
 *
 * // Load and manipulate document
 * const buffer = fs.readFileSync('template.docx');
 * const agent = await DocumentAgent.fromBuffer(buffer);
 *
 * // Get document info
 * console.log('Word count:', agent.getWordCount());
 * console.log('Variables:', agent.getVariables());
 *
 * // Edit document
 * const newAgent = agent
 *   .insertText({ paragraphIndex: 0, offset: 0 }, 'Hello ')
 *   .applyStyle(0, 'Heading1');
 *
 * // Apply template variables
 * const finalAgent = await newAgent.applyVariables({
 *   customer_name: 'Jane Doe',
 *   date: '2024-02-15',
 * });
 *
 * // Export
 * const output = await finalAgent.toBuffer();
 * fs.writeFileSync('output.docx', Buffer.from(output));
 * ```
 * @packageDocumentation
 * @public
 */
export { A as AgentContextOptions, S as ContextSelectionOptions, D as DocumentAgent, E as ExtendedSelectionContext, F as FormattedTextSegment, a as FormattingSummary, I as InsertHyperlinkOptions, b as InsertImageOptions, c as InsertTableOptions, d as InsertTextOptions, e as SelectionContextOptions, f as buildExtendedSelectionContext, g as buildSelectionContext, h as buildSelectionContextFromContext, i as createAgent, j as createAgentFromDocument, k as executeCommand, l as executeCommands, m as getAgentContext, n as getDocumentSummary, o as getSelectionFormattingSummary } from './selectionContext-BJZZW1JA.js';
export { ContentControlCreateError, CreateContentControlTarget, NewContentControlProps, RepeatingSectionError, addRepeatingSectionItem, countCharacters, countWords, createContentControl, getBlockIndexForParagraph, getBodyCharacterCount, getBodyText, getBodyWordCount, getFormattingAtPosition, getHyperlinkAtPosition, getHyperlinkText, getParagraphAtIndex, getParagraphText, getParagraphs, getRunText, getTableText, getTextAfter, getTextBefore, hasHyperlinks, hasImages, hasTables, isHeadingStyle, isPositionInHyperlink, isRepeatingSection, isRepeatingSectionItem, parseHeadingLevel, removeRepeatingSectionItem } from './agent/index.js';
export { C as ContentControlBoundError, a as ContentControlFilter, b as ContentControlInfo, c as ContentControlKindError, d as ContentControlLocation, e as ContentControlLockedError, f as ContentControlNotFoundError, g as ContentControlTypeError, h as ContentControlValue, i as ContentControlValueError, F as FindContentControlsOptions, j as findContentControl, k as findContentControls, l as formatSdtDate, m as getContentControlText, r as removeContentControl, s as setContentControlContent, n as setContentControlValue } from './contentControlValues-CxwmLrCX.js';
export { parseDocx } from './docx/parser.js';
export { s as serializeDocumentBody, a as serializeDocx, b as serializeSectionProperties } from './sectionPropertiesSerializer-CTqBBOMh.js';
export { createDocx, default as repackDocx, updateMultipleFiles } from './docx/rezip.js';
import { Document } from './types/document.js';
export { DocxPackage } from './types/document.js';
import { W as Watermark } from './watermark-D90356ZM.js';
export { a as attemptSelectiveSave } from './selectiveSave-CWaPEv0B.js';
export { b as buildPatchedDocumentXml, v as validatePatchSafety } from './selectiveXmlPatch-ypkxlTD_.js';
export { C as CreateEmptyDocumentOptions, P as ProcessTemplateOptions, a as ProcessTemplateResult, T as TemplateError, b as blendColors, c as colorsEqual, d as createDocumentWithText, e as createEmptyDocument, f as createRgbColor, g as createTemplateProcessor, h as createThemeColor, i as darkenColor, j as getContrastingColor, k as getMissingVariables, l as getTemplateTags, m as isBlack, n as isWhite, o as lightenColor, p as parseColorString, q as previewTemplate, r as processTemplate, s as processTemplateAdvanced, t as processTemplateAsBlob, u as processTemplateDetailed, v as resolveColor, w as resolveHighlightColor, x as resolveShadingColor, y as validateTemplate } from './colorResolver-CcFVVA_s.js';
export { V as VariableDetectionResult, a as VariableOccurrence, d as detectVariables, b as detectVariablesDetailed, c as detectVariablesInBody, e as detectVariablesInParagraph, f as documentHasVariables, g as extractVariablesFromText, h as formatVariable, i as hasTemplateVariables, j as isValidVariableName, p as parseVariable, r as removeVariables, k as replaceVariables, s as sanitizeVariableName } from './variableDetector-mtLOiIZG.js';
export { emuToPixels, emuToTwips, formatPx, halfPointsToPixels, pixelsToEmu, pixelsToTwips, pointsToHalfPoints, pointsToPixels, twipsToEmu, twipsToPixels } from './utils/units.js';
export { mapHexToHighlightName } from './utils/highlightColors.js';
export { b as CommandHandler, d as CommandResult, C as CorePlugin, J as JsonSchema, L as LoadedDocument, e as McpSession, f as McpToolAnnotations, g as McpToolContent, h as McpToolContext, M as McpToolDefinition, i as McpToolHandler, j as McpToolResult, C as Plugin, k as PluginCommand, b as PluginCommandHandler, l as PluginEvent, c as PluginEventListener, P as PluginOptions, a as PluginRegistrationResult, M as ToolDefinition, i as ToolHandler, j as ToolResult, Z as ZodSchemaLike, m as isZodSchema } from './types-DDQWU7-j.js';
export { P as PluginRegistry, c as createPluginRegistrar, p as pluginRegistry, r as registerPlugins } from './registry-egK81C_z.js';
export { AIAction, AIActionRequest, AgentCommand, AgentContext, AgentResponse, ApplyStyleCommand, ApplyVariablesCommand, DEFAULT_AI_ACTIONS, DeleteTextCommand, FormatParagraphCommand, FormatTextCommand, InsertHyperlinkCommand, InsertImageCommand, InsertTableCommand, InsertTextCommand, ParagraphContext, ParagraphOutline, Position, Range, ReplaceTextCommand, SectionInfo, SelectionContext, SetVariableCommand, StyleInfo, SuggestedAction, comparePositions, createCollapsedRange, createCommand, createRange, getActionDescription, getActionLabel, isPositionInRange } from './types/agentApi.js';
export { B as BlockContent, C as Comment, f as CommentRangeEnd, g as CommentRangeStart, h as Deletion, D as DocumentBody, E as Endnote, F as Footnote, H as Hyperlink, I as Image, i as Insertion, M as MoveFrom, j as MoveTo, P as Paragraph, k as ParagraphContent, l as Run, m as RunContent, S as SectionProperties, T as Table, n as TableCell, o as TableRow, p as TextContent, q as TrackedChangeInfo, r as TrackedRunChange } from './content-B8ScSBzC.js';
export { L as ListLevel, a as NumberingDefinitions } from './lists-Bn29SzeS.js';
export { P as ParagraphFormatting, T as TextFormatting } from './formatting-JhqWT_XM.js';
export { R as Relationship, a as Style, S as StyleDefinitions, T as Theme } from './styles-BGGjYqnf.js';
import './docxInput-DTbCa48g.js';
import './colors-C3vA7HUU.js';
import './docx/wrapTypes.js';
import 'jszip';

/**
 * Watermark document API
 *
 * Platform-agnostic helpers for reading and applying a document watermark.
 * Shared by the React and Vue adapters (and the imperative ref API) so the
 * "Design → Watermark" behavior stays identical across frameworks.
 *
 * A watermark lives on `HeaderFooter.watermark`. MS Word repeats the same
 * watermark across the default, first-page, and even-page headers of every
 * section, so `setDocumentWatermark`:
 *
 * 1. Applies the watermark (a per-header copy) to every existing header, and
 * 2. Creates the header parts a section needs but lacks, so the watermark
 *    still shows on title pages (`w:titlePg`) and even pages
 *    (`w:evenAndOddHeaders`) and on documents that had no header at all.
 *
 * To avoid breaking header inheritance (a section that omits a header
 * reference inherits the previous section's header — Word's "link to
 * previous"), a missing `first`/`even` header part is only created when **no**
 * header of that type exists anywhere in the document, i.e. there is nothing to
 * inherit. All updates are immutable — a new `Document` is returned so the
 * change lands in the host's undo/redo history.
 */

/** Read the document's watermark (the first header that carries one). */
declare function getDocumentWatermark(doc: Document | null | undefined): Watermark | undefined;
/**
 * Return a new `Document` with the watermark applied to all headers, or removed
 * when `watermark` is null. Creates the header parts a section needs but lacks
 * (default for a headerless doc; first/even for title/even pages) so the
 * watermark shows on every page MS Word would show it.
 */
declare function setDocumentWatermark(doc: Document, watermark: Watermark | null): Document;

/**
 * @eigenpal/docx-editor-core/headless
 *
 * Headless aggregate for Node.js scripts, CLI tools, and server-side
 * processing. Same surface as the default `.` entry, named to make the
 * "no DOM" intent explicit. Prefer the smaller subpaths (`./docx`,
 * `./agent`, `./utils`, etc.) for new code — they tree-shake better.
 *
 * @example
 * ```ts
 * import { DocumentAgent, parseDocx, pluginRegistry } from '@eigenpal/docx-editor-core/headless';
 * import { docxtemplaterPlugin } from '@eigenpal/docx-editor-core/core-plugins';
 *
 * // Register plugins
 * pluginRegistry.register(docxtemplaterPlugin);
 *
 * // Load and manipulate document
 * const buffer = fs.readFileSync('template.docx');
 * const agent = await DocumentAgent.fromBuffer(buffer);
 *
 * // Get document info
 * console.log('Word count:', agent.getWordCount());
 * console.log('Variables:', agent.getVariables());
 *
 * // Edit document
 * const newAgent = agent
 *   .insertText({ paragraphIndex: 0, offset: 0 }, 'Hello ')
 *   .applyStyle(0, 'Heading1');
 *
 * // Apply template variables
 * const finalAgent = await newAgent.applyVariables({
 *   customer_name: 'Jane Doe',
 *   date: '2024-02-15',
 * });
 *
 * // Export
 * const output = await finalAgent.toBuffer();
 * fs.writeFileSync('output.docx', Buffer.from(output));
 * ```
 * @packageDocumentation
 * @public
 */
declare const VERSION = "0.0.2";

export { Document, VERSION, getDocumentWatermark, setDocumentWatermark };
