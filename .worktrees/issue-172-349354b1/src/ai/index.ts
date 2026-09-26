export { AI_ERROR_CODES, AiError, createAiError, isAiErrorDetail } from './errors';
export type { AiErrorCode, AiErrorDetail } from './errors';
export { AiCore } from './aiCore';
export { buildCapabilityManifest } from './capabilities';
export { OP_CATALOG, OP_IDS } from './opCatalog';
export { getOpDefinition, listOpDefinitions, validateDocumentOp, validateDocumentOps } from './opRegistry';
export { validateAgainstSchema } from './schemaValidation';
export { createNpdeAiApi } from './pluginApi';
export type { AiDocumentSession, NpdeAiApi, NpdeAiApiInfo } from './pluginApi';
export type { ExportPdfOptions, ExportPdfResult, ExportPdfConflict } from './pptxExportPdf';
export type { DocxExportPdfOptions, DocxExportPdfResult, DocxExportPdfConflict } from './docxExportPdf';
export {
	buildSlideSvgElementForExport,
	exportPresentationToPdfBytes,
	resolveExportPdfOutputTarget,
	writeExportPdfArtifact,
} from './pptxExportPdf';
export { AI_COMMAND_IDS, AI_LEGACY_COMMAND_IDS, AI_COMMAND_ID_LIST } from './aiCommandIds';
export { registerAiCommands } from './registerAiCommands';
export {
	copyJsonToClipboard,
	getActiveDocumentPath,
	parseApplyRequest,
	parseDescribeRequest,
	parseSaveRequest,
	parseUndoRequest,
	parseRedoRequest,
	parseValidateRequest,
	readClipboardJson,
	resolveDocumentPath,
} from './commandProtocol';
export { getAiManifestPath, removeCapabilitiesManifest, writeCapabilitiesManifest } from './manifestWriter';
export { NPDE_AI_SKILL_PATH } from './npdeSkill';
export { writeAiSkill } from './skillWriter';
export type {
	ApplyOptions,
	ApplyPreviewChange,
	ApplyResult,
	CapabilityManifest,
	CapabilitySurfaces,
	ClipboardCommandSpec,
	DescribeResult,
	DocumentOp,
	JsonSchema,
	OpDefinition,
	OpNamespace,
	OpStatus,
	StableIdRules,
} from './types';
export { createAiRuntime } from './aiRuntime';
export type { AiRuntime, DocxViewAgentBridge, PptxViewAgentBridge } from './aiRuntime';
export { PptxDocumentService } from './pptxDocumentService';
export { describePptxFromEngine } from './pptxDescribe';
export { DocxDocumentService } from './docxDocumentService';
export { describeDocxFromBuffer } from './docxDescribe';
export type { DocxDescribeSnapshot, DocxDescribedBlock, DocxDescribedRun, DocxDescribeScope } from './docxDescribe';
export {
	parseDocumentSectionLayouts,
	parseParagraphLayout,
	patchDocumentSectionLayout,
	patchParagraphLayout,
} from './docxLayout';
export type {
	DocxParagraphAlignment,
	DocxParagraphIndent,
	DocxParagraphLayout,
	DocxParagraphSpacing,
	DocxParagraphTab,
	DocxSectionLayout,
	DocxSectionLayoutPatch,
	DocxSectionMargins,
	DocxSectionPageSize,
} from './docxLayout';
export { parseDocumentBody } from './docxOoxml';
export { AI_EDIT_UNDO_LABEL, aiUndoStore } from './aiUndoStore';
