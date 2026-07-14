import type { AiErrorDetail } from './errors';

export const AI_API_VERSION = 1;
export const CAPABILITY_SCHEMA_VERSION = 2;
export const PLUGIN_ID = 'native-powerpoint-doc-editor';

export type OpNamespace = 'pptx' | 'docx';
export type OpStatus = 'planned' | 'implemented';

export interface JsonSchema {
	type?: 'object' | 'string' | 'number' | 'integer' | 'boolean' | 'array' | 'null';
	properties?: Record<string, JsonSchema>;
	required?: string[];
	items?: JsonSchema;
	enum?: Array<string | number | boolean | null>;
	minimum?: number;
	description?: string;
	additionalProperties?: boolean;
}

export interface OpDefinition {
	id: string;
	namespace: OpNamespace;
	featureArea: string;
	description: string;
	status: OpStatus;
	parameters: JsonSchema;
	example?: DocumentOp;
}

export type DocumentOp = { op: string } & Record<string, unknown>;

export interface ApplyOptions {
	dryRun?: boolean;
}

export interface ApplyPreviewChange {
	id: string;
	field: string;
	before: unknown;
	after: unknown;
}

export interface ApplyResult {
	ok: boolean;
	dryRun?: boolean;
	changed?: string[];
	undoLabel?: string;
	canUndo?: boolean;
	canRedo?: boolean;
	preview?: ApplyPreviewChange[];
	warnings: string[];
	errors: AiErrorDetail[];
}

export interface DescribeResult {
	ok: boolean;
	errors: AiErrorDetail[];
	snapshot?: unknown;
}

export interface StableIdRules {
	pptxShape: string;
	pptxParagraph: string;
	pptxRun: string;
	docxBlock: string;
	docxRun: string;
	editableRule: string;
}

export interface ClipboardCommandSpec {
	input: string;
	output: string;
	notes?: string;
}

export interface CapabilitySurfaces {
	pluginApi: string;
	capabilitiesPath: string;
	commands: string[];
	legacyCommands?: string[];
	clipboardCommands: Record<string, ClipboardCommandSpec>;
}

export interface PptxFormatSupport {
	editable: string[];
	viewOnly: string[];
	unsupported: string[];
	notes: string[];
}

export interface PptxRuntimeSupport {
	preferredEngine: 'wasm-gc';
	fallbackEngine: 'js';
	wasmGcRequirement: string;
	fallbackLimits: string[];
}

export interface CapabilityLimitations {
	pptxFormats: PptxFormatSupport;
	pptxRuntime: PptxRuntimeSupport;
}

export interface CapabilityManifest {
	schemaVersion: number;
	apiVersion: number;
	pluginId: string;
	pluginVersion: string;
	generatedAt: string;
	enabled: boolean;
	stableIdRules: StableIdRules;
	limitations: CapabilityLimitations;
	surfaces: CapabilitySurfaces;
	operations: OpDefinition[];
}
