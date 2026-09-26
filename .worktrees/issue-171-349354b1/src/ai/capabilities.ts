import {
	EDITABLE_POWERPOINT_EXTENSIONS,
	LEGACY_POWERPOINT_EXTENSIONS,
	MACRO_ENABLED_POWERPOINT_EXTENSIONS,
} from '../powerpoint/extensions';
import { AI_COMMAND_IDS, AI_LEGACY_COMMAND_IDS } from './aiCommandIds';
import {
	AI_API_VERSION,
	CAPABILITY_SCHEMA_VERSION,
	PLUGIN_ID,
	type CapabilityManifest,
	type ClipboardCommandSpec,
	type CapabilityLimitations,
	type StableIdRules,
} from './types';
import { listOpDefinitions } from './opRegistry';
import { OP_EXAMPLES } from './opExamples';

const CAPABILITY_LIMITATIONS: CapabilityLimitations = {
	pptxFormats: {
		editable: [...EDITABLE_POWERPOINT_EXTENSIONS],
		viewOnly: [...MACRO_ENABLED_POWERPOINT_EXTENSIONS],
		unsupported: [...LEGACY_POWERPOINT_EXTENSIONS],
		notes: [
			'Legacy binary .ppt/.pps/.pot files do not open in the editor or agent API — convert to .pptx first.',
			'Macro-enabled .pptm/.ppsm/.potm and decks with vbaProject.bin open view-only; describe works, apply/save do not.',
			'Modern Open XML variants (.pptx/.ppsx/.potx) are fully editable when no embedded macros are present.',
		],
	},
	pptxRuntime: {
		preferredEngine: 'wasm-gc',
		fallbackEngine: 'js',
		wasmGcRequirement: 'Chromium 119+ WebAssembly GC (Obsidian desktop installer 1.5.8+).',
		fallbackLimits: [
			'Pure-JS pptx-svg backend loads when Wasm GC init fails (older installers, some mobile WebViews).',
			'JS fallback is slower and may not render every deck feature; export validation can still fail on advanced OOXML.',
			'Per-slide reparse prefers restore_slide_ooxml (pptx-svg >= 0.6.0); older builds reinit the whole deck on each slide edit.',
		],
	},
};

const STABLE_ID_RULES: StableIdRules = {
	pptxShape: 'slide:<slideIndex>/shape:<shapeIndex>',
	pptxParagraph: 'slide:<slideIndex>/shape:<shapeIndex>/p:<paragraphIndex>',
	pptxRun: 'slide:<slideIndex>/shape:<shapeIndex>/p:<paragraphIndex>/r:<runIndex>',
	docxBlock: 'body/p[<index>] | body/tbl[<index>]/tr[<row>]/tc[<col>]',
	docxRun: 'body/p[<index>]/r[<runIndex>]',
	docxTextPosition: '{ blockId: paragraph id, offset: 0-based char offset in paragraph plain text, runId?: optional anchor run }',
	docxTextRange: '{ start: docxTextPosition, end: docxTextPosition } — same part; end block/run must not precede start',
	editableRule: 'PPTX shape indices must be integers >= 0. Negative indices are inherited placeholders.',
};

const CLIPBOARD_COMMANDS: Record<string, ClipboardCommandSpec> = {
	'npde-ai-capabilities': {
		input: '{} or empty clipboard',
		output: 'CapabilityManifest',
		notes: 'Ignores clipboard input; copies the live manifest.',
	},
	'npde-ai-describe': {
		input: '{ "path"?: "vault/file.pptx" }',
		output: 'DescribeResult',
		notes: 'Omit path to use the active PPTX/DOCX file.',
	},
	'npde-ai-apply': {
		input: '{ "path"?: "vault/file.pptx", "ops": DocumentOp[], "dryRun"?: boolean }',
		output: 'ApplyResult',
		notes: 'Omit path to use the active PPTX/DOCX file.',
	},
	'npde-ai-validate': {
		input: '{ "ops": DocumentOp[] }',
		output: '{ ok: boolean, errors: AiErrorDetail[] }',
	},
	'npde-ai-save': {
		input: '{ "path"?: "vault/file.pptx" }',
		output: '{ ok: boolean, errors: AiErrorDetail[] }',
		notes: 'Runs full save validation for headless sessions.',
	},
	'npde-ai-undo': {
		input: '{ "path"?: "vault/file.pptx" }',
		output: '{ ok: boolean, errors: AiErrorDetail[] }',
		notes: 'Undoes the latest agent edit. Open views use editor undo; headless sessions use the agent undo stack.',
	},
	'npde-ai-redo': {
		input: '{ "path"?: "vault/file.pptx" }',
		output: '{ ok: boolean, errors: AiErrorDetail[] }',
		notes: 'Redoes the latest undone agent edit.',
	},
};

export interface BuildCapabilityManifestOptions {
	pluginVersion: string;
	enabled: boolean;
}

const PLUGIN_API_METHODS: Record<string, ClipboardCommandSpec> = {
	createDocument: {
		input: '{ "path": "folder/file.docx", "kind": "docx"|"pptx", "paragraphs"?: string[], "overwrite"?: boolean }',
		output: '{ ok: boolean, path?: string, errors: AiErrorDetail[] }',
		notes: 'Creates a blank DOCX/PPTX in the vault (same packages as New DOCX/PPTX). For DOCX, optional paragraphs fills the body via docx.replaceBodyParagraphs semantics. Does not open a leaf.',
	},
	exportPdf: {
		input: '{ "path": "vault/file.potx"|"vault/file.pptx", "outputPath"?: "vault/file.pdf", "slideIndices"?: number[], "conflict"?: "replace"|"keep-both", "scale"?: number }',
		output: '{ ok: boolean, path?: string, bytes?: number, slideCount?: number, errors: AiErrorDetail[] }',
		notes: 'Rasters PPTX/POTX/PPSX slides to a vault PDF via the NPDE SVG renderer (same path as Export → Whole deck PDF). Page size uses p:sldSz EMUs (true inches). Default output is <basename>.pdf beside the source; conflict defaults to keep-both. Does not call Microsoft PowerPoint.',
	},
	exportDocxPdf: {
		input: '{ "path": "vault/file.docx", "outputPath"?: "vault/file.pdf", "conflict"?: "replace"|"keep-both" }',
		output: '{ ok: boolean, path?: string, bytes?: number, errors: AiErrorDetail[] }',
		notes: 'Exports the rendered NPDE DOCX pages to a vault PDF. The DOCX must be open and rendered in an active NPDE view so pagination matches the editor preview.',
	},
};

export function buildCapabilityManifest(options: BuildCapabilityManifestOptions): CapabilityManifest {
	return {
		schemaVersion: CAPABILITY_SCHEMA_VERSION,
		apiVersion: AI_API_VERSION,
		pluginId: PLUGIN_ID,
		pluginVersion: options.pluginVersion,
		generatedAt: new Date().toISOString(),
		enabled: options.enabled,
		stableIdRules: STABLE_ID_RULES,
		limitations: CAPABILITY_LIMITATIONS,
		surfaces: {
			pluginApi: `app.plugins.plugins['${PLUGIN_ID}'].ai`,
			capabilitiesPath: `.obsidian/plugins/${PLUGIN_ID}/ai/capabilities.json`,
			commands: [
				AI_COMMAND_IDS.capabilities,
				AI_COMMAND_IDS.describe,
				AI_COMMAND_IDS.apply,
				AI_COMMAND_IDS.validate,
				AI_COMMAND_IDS.save,
				AI_COMMAND_IDS.undo,
				AI_COMMAND_IDS.redo,
			],
			legacyCommands: [
				AI_LEGACY_COMMAND_IDS.capabilities,
				AI_LEGACY_COMMAND_IDS.describe,
				AI_LEGACY_COMMAND_IDS.apply,
			],
			clipboardCommands: CLIPBOARD_COMMANDS,
			pluginApiMethods: PLUGIN_API_METHODS,
		},
		operations: listOpDefinitions().map((operation) => ({
			...operation,
			...(OP_EXAMPLES[operation.id] ? { example: OP_EXAMPLES[operation.id] } : {}),
		})),
	};
}
