import { App, Component, FileView, Modal, Platform, TFile, WorkspaceLeaf, normalizePath } from 'obsidian';
import type { Translations } from './docx/runtime/contract';
import type { I18nService } from './i18n/I18nService';
import { showI18nNotice } from './i18n/notify';
import { loadDocxEditorChunk } from './docxEditorLoader';
import {
	getAvailableNumberedPath as resolveNumberedArtifactPath,
	resolveArtifactConflict,
	type ArtifactConflictChoice,
} from './export/artifactPaths';
import { buildSiblingPath, getVaultFolderPrefix } from './vault/paths';
import { renameFileToSiblingName } from './vault/renameFlow';
import { findHiddenDocxText, type HiddenTextFinding } from './docxHiddenTextScanner';
import { ensureDocxDefaultStyles } from './docxStyleDefaults';
import { extractDocxMarkdown, extractDocxText } from './docxTextExtractor';
import { isElement, isHTMLElement, isNode } from './domGuards';
import { scheduleIdleWork } from './idleSchedule';
import { createLoadTrace, monotonicNow, type LoadTrace } from './loadTrace';
import {
	logLifecycleStep,
	startOpenHeartbeat,
	traceSyncStep,
} from './debugInstrumentation';
import { debugLog, errorLog, infoLog, warnLog } from './logger';
import { aiUndoStore } from './ai/aiUndoStore';
import { closeModalDomScope, loadModalDomScope, openModalDomScope } from './modalDomScope';
import {
	normalizeEditorThemePreference,
	type EditorThemeResolution,
	type EditorThemePreference,
} from './settings';
import type { DocxReactMount } from './DocxReactMount';
import type { DocxReactViewHandle, DocxReactViewProps } from './DocxReactView';
import type { DocumentWordCount } from './documentWordCount';
import {
	bindPopoverDismissHandlers,
	configureMenuItemButton,
	createMenuItem,
	createPopoverShell,
} from './menuControls';
import {
	DOCX_EDITOR_MENUBAR_SELECTOR,
	DOCX_EDITOR_MENU_BUTTON_SELECTOR,
	DOCX_EDITOR_MENU_DROPDOWN_SELECTOR,
	DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR,
	DOCX_EDITOR_MENU_ROOT_SELECTOR,
	DOCX_EDITOR_ROOT_SELECTOR,
	EDITOR_CHROME_MENU_ITEMS,
	markEditorChromeMenuButton,
	markEditorChromeMenuItem,
	markEditorChromeNoToolbarTooltip,
	stampDocxEditorChromeRegions,
} from './docxEditorChromeMarkers';
import { getDocxEditorChromeRegionSelector } from './editorChromeRegions';

import { neutralizeToolbarButtonTooltipSources } from './docxToolbarTooltip';
import { VIEW_TYPE_DOCX } from './docxViewConstants';
import {
	createDocxEditorAdapter,
	type Disposable,
	type DocxEditorAdapterController,
} from './docx/adapter/DocxEditorAdapter';
import { createDetachedDocxEditorChromeElement } from './docxEditorChromeDom';
import { DocxAgentReloadGuard, type DocxReloadIdentity } from './docx/DocxAgentReloadGuard';

export { VIEW_TYPE_DOCX };

type UnsavedDocxChoice = 'save' | 'discard' | 'cancel';
type DocxPathChoice = string | null;
type DocxExportFormatId = 'pdf' | 'docx' | 'html' | 'txt' | 'md' | 'rtf';
type DocxExportChoice = { name: string; format: DocxExportFormatId } | null;
type DocxConflictChoice = 'overwrite' | 'cancel';
type ExistingFileChoice = 'replace' | 'copy';
const AGENT_RELOAD_READY_TIMEOUT_MS = 15_000;
type EditorOptionSearchActionId =
	| 'save'
	| 'save-as'
	| 'duplicate'
	| 'paste'
	| 'paste-without-formatting'
	| 'export-pdf'
	| 'export-docx'
	| 'export-html'
	| 'export-txt'
	| 'export-md'
	| 'export-rtf'
	| 'find'
	| 'find-replace'
	| 'insert-image'
	| 'custom-table'
	| 'import-font'
	| 'find-hidden-text'
	| 'page-setup'
	| 'page-break'
	| 'table-of-contents'
	| 'left-to-right'
	| 'right-to-left'
	| 'mode-editing'
	| 'mode-suggesting'
	| 'mode-viewing'
	| 'zoom-75'
	| 'zoom-100'
	| 'zoom-125';

interface DocxFileSignature {
	path: string;
	mtime: number;
	size: number;
}

type DocxSaveOrigin = DocxReloadIdentity;

interface DocxExportFormat {
	id: DocxExportFormatId;
	label: string;
	extension: string;
}

interface EditorOptionSearchBaseItem {
	id: string;
	label: string;
	keywords: readonly string[];
}

interface EditorOptionSearchActionItem extends EditorOptionSearchBaseItem {
	kind: 'action';
	actionId: EditorOptionSearchActionId;
}

interface EditorOptionSearchControlQueryItem extends EditorOptionSearchBaseItem {
	kind: 'control-query';
	labels: readonly string[];
}

interface EditorOptionSearchControlItem extends EditorOptionSearchBaseItem {
	kind: 'control';
	element: HTMLElement;
}

type EditorOptionSearchItem =
	| EditorOptionSearchActionItem
	| EditorOptionSearchControlQueryItem
	| EditorOptionSearchControlItem;

const DOCX_EXPORT_FORMATS: readonly DocxExportFormat[] = [
	{ id: 'pdf', label: 'PDF document (.pdf)', extension: 'pdf' },
	{ id: 'docx', label: 'Word document (.docx)', extension: 'docx' },
	{ id: 'html', label: 'Web page (.html)', extension: 'html' },
	{ id: 'txt', label: 'Plain text (.txt)', extension: 'txt' },
	{ id: 'md', label: 'Markdown (.md)', extension: 'md' },
	{ id: 'rtf', label: 'Rich Text Format (.rtf)', extension: 'rtf' },
];
const KNOWN_EXPORT_EXTENSION_PATTERN = /\.(?:docx|pdf|html?|txt|md|markdown|rtf)$/i;
const DEFAULT_EXPORT_FORMAT: DocxExportFormatId = 'pdf';
const EDITOR_OPTION_SEARCH_BASE_ITEMS: readonly (EditorOptionSearchActionItem | EditorOptionSearchControlQueryItem)[] = [
	{ kind: 'action', id: 'action:save', actionId: 'save', label: 'Save', keywords: ['write', 'autosave'] },
	{ kind: 'action', id: 'action:save-as', actionId: 'save-as', label: 'Save as...', keywords: ['copy', 'new docx'] },
	{ kind: 'action', id: 'action:duplicate', actionId: 'duplicate', label: 'Duplicate current DOCX', keywords: ['copy', 'clone'] },
	{ kind: 'action', id: 'action:paste', actionId: 'paste', label: 'Paste', keywords: ['edit', 'clipboard', 'command v', 'cmd v', 'ctrl v'] },
	{ kind: 'action', id: 'action:paste-without-formatting', actionId: 'paste-without-formatting', label: 'Paste without formatting', keywords: ['edit', 'clipboard', 'plain text', 'paste as text'] },
	{ kind: 'action', id: 'action:export-pdf', actionId: 'export-pdf', label: 'Export as PDF', keywords: ['pdf', 'export as'] },
	{ kind: 'action', id: 'action:export-docx', actionId: 'export-docx', label: 'Export as DOCX', keywords: ['word', 'export as'] },
	{ kind: 'action', id: 'action:export-html', actionId: 'export-html', label: 'Export as HTML', keywords: ['web page', 'export as'] },
	{ kind: 'action', id: 'action:export-txt', actionId: 'export-txt', label: 'Export as plain text', keywords: ['txt', 'text', 'export as'] },
	{ kind: 'action', id: 'action:export-md', actionId: 'export-md', label: 'Export as Markdown', keywords: ['md', 'export as'] },
	{ kind: 'action', id: 'action:export-rtf', actionId: 'export-rtf', label: 'Export as RTF', keywords: ['rich text', 'export as'] },
	{ kind: 'action', id: 'action:find', actionId: 'find', label: 'Find in document', keywords: ['search', 'text'] },
	{ kind: 'action', id: 'action:find-replace', actionId: 'find-replace', label: 'Find and replace', keywords: ['replace', 'search'] },
	{ kind: 'action', id: 'action:insert-image', actionId: 'insert-image', label: 'Insert image', keywords: ['picture', 'photo', 'media'] },
	{ kind: 'action', id: 'action:custom-table', actionId: 'custom-table', label: 'Custom table', keywords: ['insert table', 'rows', 'columns'] },
	{ kind: 'action', id: 'action:import-font', actionId: 'import-font', label: 'Import font', keywords: ['typeface', 'ttf', 'otf', 'woff'] },
	{ kind: 'action', id: 'action:find-hidden-text', actionId: 'find-hidden-text', label: 'Find hidden text', keywords: ['security', 'prompt injection', 'invisible'] },
	{ kind: 'action', id: 'action:page-setup', actionId: 'page-setup', label: 'Page setup', keywords: ['paper', 'margins', 'orientation'] },
	{ kind: 'action', id: 'action:page-break', actionId: 'page-break', label: 'Page break', keywords: ['insert page'] },
	{ kind: 'action', id: 'action:table-of-contents', actionId: 'table-of-contents', label: 'Table of contents', keywords: ['toc', 'outline'] },
	{ kind: 'action', id: 'action:left-to-right', actionId: 'left-to-right', label: 'Left to right', keywords: ['ltr', 'direction'] },
	{ kind: 'action', id: 'action:right-to-left', actionId: 'right-to-left', label: 'Right to left', keywords: ['rtl', 'direction'] },
	{ kind: 'action', id: 'action:mode-editing', actionId: 'mode-editing', label: 'Editing mode', keywords: ['edit'] },
	{ kind: 'action', id: 'action:mode-suggesting', actionId: 'mode-suggesting', label: 'Suggesting mode', keywords: ['suggest', 'review'] },
	{ kind: 'action', id: 'action:mode-viewing', actionId: 'mode-viewing', label: 'Viewing mode', keywords: ['view', 'read only'] },
	{ kind: 'action', id: 'action:zoom-75', actionId: 'zoom-75', label: 'Zoom 75%', keywords: ['small'] },
	{ kind: 'action', id: 'action:zoom-100', actionId: 'zoom-100', label: 'Zoom 100%', keywords: ['actual size', 'normal'] },
	{ kind: 'action', id: 'action:zoom-125', actionId: 'zoom-125', label: 'Zoom 125%', keywords: ['large'] },
	{ kind: 'control-query', id: 'control:bold', label: 'Bold', labels: ['Bold', 'Bold (Ctrl+B)', 'format bold', 'format_bold'], keywords: ['ctrl b', 'text formatting'] },
	{ kind: 'control-query', id: 'control:italic', label: 'Italic', labels: ['Italic', 'Italic (Ctrl+I)', 'format italic', 'format_italic'], keywords: ['ctrl i', 'text formatting'] },
	{ kind: 'control-query', id: 'control:underline', label: 'Underline', labels: ['Underline', 'Underline (Ctrl+U)', 'format underlined', 'format_underlined'], keywords: ['ctrl u', 'text formatting'] },
	{ kind: 'control-query', id: 'control:strikethrough', label: 'Strikethrough', labels: ['Strikethrough', 'strikethrough s', 'strikethrough_s'], keywords: ['strike', 'text formatting'] },
	{ kind: 'control-query', id: 'control:superscript', label: 'Superscript', labels: ['Superscript'], keywords: ['script', 'raise text'] },
	{ kind: 'control-query', id: 'control:subscript', label: 'Subscript', labels: ['Subscript'], keywords: ['script', 'lower text'] },
	{ kind: 'control-query', id: 'control:clear-formatting', label: 'Clear formatting', labels: ['Clear formatting', 'format clear', 'format_clear'], keywords: ['remove formatting'] },
	{ kind: 'control-query', id: 'control:insert-link', label: 'Insert link', labels: ['Insert link', 'Insert link (Ctrl+K)', 'link'], keywords: ['hyperlink', 'url', 'ctrl k'] },
	{ kind: 'control-query', id: 'control:undo', label: 'Undo', labels: ['Undo', 'Undo (Ctrl+Z)'], keywords: ['history', 'ctrl z'] },
	{ kind: 'control-query', id: 'control:redo', label: 'Redo', labels: ['Redo', 'Redo (Ctrl+Y)'], keywords: ['history', 'ctrl y'] },
	{ kind: 'control-query', id: 'control:bullet-list', label: 'Bullet list', labels: ['Bullet List', 'format list bulleted', 'format_list_bulleted'], keywords: ['bullets', 'list'] },
	{ kind: 'control-query', id: 'control:numbered-list', label: 'Numbered list', labels: ['Numbered List', 'format list numbered', 'format_list_numbered'], keywords: ['numbers', 'ordered list'] },
	{ kind: 'control-query', id: 'control:decrease-indent', label: 'Decrease indent', labels: ['Decrease Indent', 'format indent decrease', 'format_indent_decrease'], keywords: ['outdent', 'list'] },
	{ kind: 'control-query', id: 'control:increase-indent', label: 'Increase indent', labels: ['Increase Indent', 'format indent increase', 'format_indent_increase'], keywords: ['indent', 'list'] },
	{ kind: 'control-query', id: 'control:align-left', label: 'Align left', labels: ['Align Left', 'format align left', 'format_align_left'], keywords: ['alignment'] },
	{ kind: 'control-query', id: 'control:align-center', label: 'Center align', labels: ['Center', 'Align Center', 'format align center', 'format_align_center'], keywords: ['alignment'] },
	{ kind: 'control-query', id: 'control:align-right', label: 'Align right', labels: ['Align Right', 'format align right', 'format_align_right'], keywords: ['alignment'] },
	{ kind: 'control-query', id: 'control:justify', label: 'Justify', labels: ['Justify', 'format align justify', 'format_align_justify'], keywords: ['alignment'] },
	{ kind: 'control-query', id: 'control:font-family', label: 'Font family', labels: ['Font family', 'Font', 'Select font family'], keywords: ['typeface'] },
	{ kind: 'control-query', id: 'control:font-size', label: 'Font size', labels: ['Font size', 'Select font size'], keywords: ['text size', 'point size'] },
	{ kind: 'control-query', id: 'control:font-size-decrease', label: 'Decrease font size', labels: ['Decrease font size', 'font size decrease'], keywords: ['smaller text'] },
	{ kind: 'control-query', id: 'control:font-size-increase', label: 'Increase font size', labels: ['Increase font size', 'font size increase'], keywords: ['larger text'] },
	{ kind: 'control-query', id: 'control:font-color', label: 'Font color', labels: ['Font Color', 'Text color'], keywords: ['text color'] },
	{ kind: 'control-query', id: 'control:highlight-color', label: 'Text highlight color', labels: ['Text Highlight Color', 'Highlight Color'], keywords: ['highlight', 'background color'] },
	{ kind: 'control-query', id: 'control:line-spacing', label: 'Line spacing', labels: ['Line spacing', 'Line spacing: Single', 'Line spacing: Double'], keywords: ['paragraph spacing'] },
];

function getExportFormat(formatId: DocxExportFormatId): DocxExportFormat {
	return DOCX_EXPORT_FORMATS.find(format => format.id === formatId) ?? DOCX_EXPORT_FORMATS[0]!;
}

function getExportBaseName(name: string): string {
	return name
		.trim()
		.replace(/[\\/]/g, '-')
		.replace(KNOWN_EXPORT_EXTENSION_PATTERN, '')
		.trim();
}

function withExportExtension(name: string, formatId: DocxExportFormatId): string {
	const baseName = getExportBaseName(name);
	if (!baseName) {
		return '';
	}

	return `${baseName}.${getExportFormat(formatId).extension}`;
}

function escapeHtml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}

function createPlainTextHtml(text: string, title: string): string {
	const paragraphs = text
		.split(/\n{2,}/)
		.map(paragraph => paragraph.trim())
		.filter(Boolean);
	const body = paragraphs.length > 0
		? paragraphs.map(paragraph => `<p>${escapeHtml(paragraph).replace(/\n/g, '<br>')}</p>`).join('\n')
		: '<p></p>';

	return [
		'<!doctype html>',
		'<html lang="en">',
		'<head>',
		'<meta charset="utf-8">',
		'<meta name="viewport" content="width=device-width, initial-scale=1">',
		`<title>${escapeHtml(title)}</title>`,
		'<style>body{font-family:Arial,Helvetica,sans-serif;line-height:1.5;margin:48px;max-width:760px;}p{margin:0 0 1em;}</style>',
		'</head>',
		'<body>',
		'<main>',
		body,
		'</main>',
		'</body>',
		'</html>',
		'',
	].join('\n');
}

function toSignedRtfUnit(unit: number): number {
	// RTF \uN expects a signed 16-bit integer, so code units above 32767
	// must be expressed in their negative two's-complement form.
	return unit > 32767 ? unit - 65536 : unit;
}

function escapeRtf(value: string): string {
	let result = '';

	// Iterating with for..of yields whole code points, so astral characters
	// (emoji, anything above U+FFFF) arrive as a single `char` here.
	for (const char of value) {
		if (char === '\\') {
			result += '\\\\';
			continue;
		}
		if (char === '{') {
			result += '\\{';
			continue;
		}
		if (char === '}') {
			result += '\\}';
			continue;
		}
		if (char === '\n') {
			result += '\\par\n';
			continue;
		}

		const codePoint = char.codePointAt(0) ?? 0;
		if (codePoint <= 127) {
			result += char;
			continue;
		}

		if (codePoint > 0xFFFF) {
			// RTF has no astral escape; emit the UTF-16 surrogate pair so
			// readers reconstruct the original code point.
			const offset = codePoint - 0x10000;
			const highSurrogate = 0xD800 + (offset >> 10);
			const lowSurrogate = 0xDC00 + (offset & 0x3FF);
			result += `\\u${toSignedRtfUnit(highSurrogate)}?\\u${toSignedRtfUnit(lowSurrogate)}?`;
			continue;
		}

		result += `\\u${toSignedRtfUnit(codePoint)}?`;
	}

	return result;
}

function createPlainTextRtf(text: string): string {
	return `{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Helvetica;}}\\f0\\fs24\n${escapeRtf(text)}\n}\n`;
}

function getBinaryExportContent(content: ArrayBuffer | ArrayBufferView | string): ArrayBuffer | null {
	if (typeof content === 'string') {
		return null;
	}

	if (content instanceof ArrayBuffer || Object.prototype.toString.call(content) === '[object ArrayBuffer]') {
		return content as ArrayBuffer;
	}

	if (ArrayBuffer.isView(content)) {
		const copy = new Uint8Array(content.byteLength);
		copy.set(new Uint8Array(content.buffer, content.byteOffset, content.byteLength));
		return copy.buffer;
	}

	return null;
}

class UnsavedDocxModal extends Modal {
	private resolveChoice: (choice: UnsavedDocxChoice) => void;
	private resolved = false;
	private domScope?: Component;

	constructor(
		app: App,
		private fileName: string,
		resolveChoice: (choice: UnsavedDocxChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		this.domScope = openModalDomScope();
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'Save changes?' });
		contentEl.createEl('p', { text: `${this.fileName} has unsaved changes.` });

		const buttonRow = contentEl.createDiv({ cls: 'native-powerpoint-doc-editor-unsaved-actions' });
		const cancelButton = buttonRow.createEl('button', { text: 'Cancel' });
		const discardButton = buttonRow.createEl('button', { text: 'Discard' });
		const saveButton = buttonRow.createEl('button', { text: 'Save' });
		saveButton.addClass('mod-cta');

		this.domScope.registerDomEvent(cancelButton, 'click', () => this.choose('cancel'));
		this.domScope.registerDomEvent(discardButton, 'click', () => this.choose('discard'));
		this.domScope.registerDomEvent(saveButton, 'click', () => this.choose('save'));
		loadModalDomScope(this.domScope);
	}

	onClose() {
		closeModalDomScope(this.domScope);
		this.domScope = undefined;
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose('cancel');
		}
	}

	private choose(choice: UnsavedDocxChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

class DocxPathModal extends Modal {
	private resolveChoice: (choice: DocxPathChoice) => void;
	private resolved = false;
	private domScope?: Component;

	constructor(
		app: App,
		private title: string,
		private description: string,
		private initialPath: string,
		private actionLabel: string,
		resolveChoice: (choice: DocxPathChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		this.domScope = openModalDomScope();
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: this.title });
		contentEl.createEl('p', { text: this.description });

		const form = contentEl.createEl('form', { cls: 'native-powerpoint-doc-editor-path-form' });
		const input = form.createEl('input', {
			cls: 'native-powerpoint-doc-editor-path-input',
			type: 'text',
		});
		input.value = this.initialPath;
		input.setAttribute('spellcheck', 'false');

		const buttonRow = form.createDiv({ cls: 'native-powerpoint-doc-editor-unsaved-actions' });
		const cancelButton = buttonRow.createEl('button', { text: 'Cancel', type: 'button' });
		const saveButton = buttonRow.createEl('button', { text: this.actionLabel, type: 'submit' });
		saveButton.addClass('mod-cta');

		this.domScope.registerDomEvent(cancelButton, 'click', () => this.choose(null));
		this.domScope.registerDomEvent(form, 'submit', (evt) => {
			evt.preventDefault();
			this.choose(input.value);
		});

		input.focus();
		input.select();
		loadModalDomScope(this.domScope);
	}

	onClose() {
		closeModalDomScope(this.domScope);
		this.domScope = undefined;
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose(null);
		}
	}

	private choose(choice: DocxPathChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

class DocxExportModal extends Modal {
	private resolveChoice: (choice: DocxExportChoice) => void;
	private resolved = false;
	private domScope?: Component;

	constructor(
		app: App,
		private initialName: string,
		private initialFormat: DocxExportFormatId,
		resolveChoice: (choice: DocxExportChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		this.domScope = openModalDomScope();
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'Export as' });
		contentEl.createEl('p', { text: 'Export a copy next to the original file. If the file already exists, you can replace it or keep both.' });

		const form = contentEl.createEl('form', { cls: 'native-powerpoint-doc-editor-path-form' });
		const formatLabel = form.createEl('label', { cls: 'native-powerpoint-doc-editor-export-field' });
		formatLabel.createSpan({ text: 'Format' });
		const formatSelect = formatLabel.createEl('select', { cls: 'native-powerpoint-doc-editor-path-select' });
		for (const format of DOCX_EXPORT_FORMATS) {
			const option = formatSelect.createEl('option', { text: format.label, value: format.id });
			option.selected = format.id === this.initialFormat;
		}

		const nameLabel = form.createEl('label', { cls: 'native-powerpoint-doc-editor-export-field' });
		nameLabel.createSpan({ text: 'File name' });
		const input = nameLabel.createEl('input', {
			cls: 'native-powerpoint-doc-editor-path-input',
			type: 'text',
		});
		input.value = withExportExtension(this.initialName, this.initialFormat);
		input.setAttribute('spellcheck', 'false');

		this.domScope.registerDomEvent(formatSelect, 'change', () => {
			input.value = withExportExtension(input.value, formatSelect.value as DocxExportFormatId);
		});

		const buttonRow = form.createDiv({ cls: 'native-powerpoint-doc-editor-unsaved-actions' });
		const cancelButton = buttonRow.createEl('button', { text: 'Cancel', type: 'button' });
		const exportButton = buttonRow.createEl('button', { text: 'Export', type: 'submit' });
		exportButton.addClass('mod-cta');

		this.domScope.registerDomEvent(cancelButton, 'click', () => this.choose(null));
		this.domScope.registerDomEvent(form, 'submit', (evt) => {
			evt.preventDefault();
			this.choose({
				name: input.value,
				format: formatSelect.value as DocxExportFormatId,
			});
		});

		input.focus();
		input.select();
		loadModalDomScope(this.domScope);
	}

	onClose() {
		closeModalDomScope(this.domScope);
		this.domScope = undefined;
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose(null);
		}
	}

	private choose(choice: DocxExportChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

class ExistingFileModal extends Modal {
	private resolveChoice: (choice: ExistingFileChoice) => void;
	private resolved = false;
	private domScope?: Component;

	constructor(
		app: App,
		private filePath: string,
		resolveChoice: (choice: ExistingFileChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		this.domScope = openModalDomScope();
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'This file already exists.' });
		contentEl.createEl('p', { text: 'Replace it?' });
		contentEl.createEl('p', { text: this.filePath });

		const buttonRow = contentEl.createDiv({ cls: 'native-powerpoint-doc-editor-unsaved-actions' });
		const noButton = buttonRow.createEl('button', { text: 'No' });
		const yesButton = buttonRow.createEl('button', { text: 'Yes' });
		yesButton.addClass('mod-warning');

		this.domScope.registerDomEvent(noButton, 'click', () => this.choose('copy'));
		this.domScope.registerDomEvent(yesButton, 'click', () => this.choose('replace'));
		loadModalDomScope(this.domScope);
	}

	onClose() {
		closeModalDomScope(this.domScope);
		this.domScope = undefined;
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose('copy');
		}
	}

	private choose(choice: ExistingFileChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

class HiddenTextScanModal extends Modal {
	constructor(
		app: App,
		private fileName: string,
		private findings: HiddenTextFinding[],
		private partsScanned: number,
	) {
		super(app);
	}

	onOpen() {
		const { contentEl } = this;
		this.modalEl.addClass('native-powerpoint-doc-editor-hidden-text-shell');
		contentEl.empty();
		contentEl.addClass('native-powerpoint-doc-editor-hidden-text-modal');
		contentEl.createEl('h2', { text: 'Find hidden text' });

		if (this.findings.length === 0) {
			contentEl.createEl('p', {
				cls: 'native-powerpoint-doc-editor-hidden-text-summary native-powerpoint-doc-editor-hidden-text-empty',
				text: `No hidden, white, or tiny text was found in ${this.fileName}. Scanned ${this.partsScanned} document part(s).`,
			});
			return;
		}

		contentEl.createEl('p', {
			cls: 'native-powerpoint-doc-editor-hidden-text-summary',
			text: `Found ${this.findings.length} suspicious hidden text item(s) in ${this.fileName}. Review before pasting this document into an AI tool.`,
		});

		const list = contentEl.createDiv({ cls: 'native-powerpoint-doc-editor-hidden-text-results' });
		for (const finding of this.findings) {
			const resultEl = list.createDiv({ cls: 'native-powerpoint-doc-editor-hidden-text-result' });
			const header = resultEl.createDiv({ cls: 'native-powerpoint-doc-editor-hidden-text-header' });
			header.createSpan({
				cls: 'native-powerpoint-doc-editor-hidden-text-location',
				text: `${finding.partLabel}, paragraph ${finding.paragraphNumber}`,
			});
			header.createSpan({
				cls: 'native-powerpoint-doc-editor-hidden-text-path',
				text: finding.partPath,
			});

			const reasons = resultEl.createDiv({ cls: 'native-powerpoint-doc-editor-hidden-text-reasons' });
			for (const reason of finding.reasons) {
				reasons.createSpan({ cls: 'native-powerpoint-doc-editor-hidden-text-reason', text: reason });
			}
			for (const signal of finding.promptInjectionSignals) {
				reasons.createSpan({ cls: 'native-powerpoint-doc-editor-hidden-text-reason mod-warning', text: `Prompt-like text: ${signal}` });
			}

			resultEl.createEl('pre', {
				cls: 'native-powerpoint-doc-editor-hidden-text-snippet',
				text: finding.text,
			});
		}
	}

	onClose() {
		this.modalEl.removeClass('native-powerpoint-doc-editor-hidden-text-shell');
		this.contentEl.empty();
	}
}

class ExternalDocxChangeModal extends Modal {
	private resolveChoice: (choice: DocxConflictChoice) => void;
	private resolved = false;
	private domScope?: Component;

	constructor(
		app: App,
		private fileName: string,
		resolveChoice: (choice: DocxConflictChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		this.domScope = openModalDomScope();
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'File changed on disk' });
		contentEl.createEl('p', {
			text: `${this.fileName} was modified outside Native PowerPoint Doc Editor after it was opened.`,
		});
		contentEl.createEl('p', {
			text: 'Saving now will overwrite those outside changes. Cancel and use save as... Or duplicate current docx if you want to keep both versions.',
		});

		const buttonRow = contentEl.createDiv({ cls: 'native-powerpoint-doc-editor-unsaved-actions' });
		const cancelButton = buttonRow.createEl('button', { text: 'Cancel save' });
		const overwriteButton = buttonRow.createEl('button', { text: 'Overwrite anyway' });
		overwriteButton.addClass('mod-warning');

		this.domScope.registerDomEvent(cancelButton, 'click', () => this.choose('cancel'));
		this.domScope.registerDomEvent(overwriteButton, 'click', () => this.choose('overwrite'));
		loadModalDomScope(this.domScope);
	}

	onClose() {
		closeModalDomScope(this.domScope);
		this.domScope = undefined;
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose('cancel');
		}
	}

	private choose(choice: DocxConflictChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

function normalizeMenuText(text: string): string {
	return text.replace(/\s+/g, ' ').trim().toLowerCase();
}

function toTitleCase(value: string): string {
	return value.replace(/\b[a-z]/g, char => char.toUpperCase());
}

function cleanEditorOptionLabel(rawLabel: string): string {
	let label = rawLabel
		.replace(/[_-]+/g, ' ')
		.replace(/\s+/g, ' ')
		.trim();

	label = label
		.replace(/\s*\((?:ctrl|cmd|command|shift|alt|option|⌘)[^)]+\)\s*$/i, '')
		.replace(/\s*(?:ctrl|cmd|command|⌘)\s*\+.*$/i, '')
		.replace(/^format\s+/i, '')
		.trim();

	return /^[a-z0-9 ]+$/.test(label) ? toTitleCase(label) : label;
}

function getEditorOptionRawLabel(element: HTMLElement): string {
	return element.getAttribute('data-native-powerpoint-doc-editor-search-label')
		?? element.getAttribute('aria-label')
		?? element.dataset.nativePowerPointDocEditorToolbarLabel
		?? element.dataset.nativePowerPointDocEditorNativeTitle
		?? element.dataset.nativePowerPointDocEditorTooltipTitle
		?? element.getAttribute('title')
		?? element.textContent
		?? element.dataset.testid
		?? '';
}

function getEditorOptionControlLabel(element: HTMLElement): string {
	return cleanEditorOptionLabel(getEditorOptionRawLabel(element));
}

function getEditorOptionControlKeywords(element: HTMLElement, label: string): string[] {
	const keywords = new Set<string>();
	const testId = element.dataset.testid;
	const rawText = element.textContent?.trim();
	const groupLabel = element.closest<HTMLElement>('[role="group"][aria-label]')?.getAttribute('aria-label');
	const nativeTitle = element.dataset.nativePowerPointDocEditorNativeTitle ?? element.dataset.nativePowerPointDocEditorTooltipTitle ?? element.getAttribute('title');

	for (const value of [label, testId, rawText, groupLabel, nativeTitle]) {
		const normalized = cleanEditorOptionLabel(value ?? '');
		if (normalized) {
			keywords.add(normalized);
		}
	}

	return Array.from(keywords);
}

function isTopLevelEditorMenuButton(element: HTMLElement): boolean {
	const parent = element.parentElement;
	return Boolean(
		parent
		&& parent.parentElement?.getAttribute('role') === 'menubar'
		&& parent.querySelector(DOCX_EDITOR_MENU_BUTTON_SELECTOR) === element
	);
}

function getActiveEditorMenubar(hostEl: HTMLElement): HTMLElement | null {
	const menubars = hostEl.querySelectorAll<HTMLElement>(DOCX_EDITOR_MENUBAR_SELECTOR);
	if (menubars.length === 0) {
		return null;
	}

	return menubars[menubars.length - 1] ?? null;
}

function dedupeEditorChromeMenuItem(hostEl: HTMLElement, selector: string): HTMLElement | null {
	const menubar = getActiveEditorMenubar(hostEl);
	const matches = Array.from(hostEl.querySelectorAll<HTMLElement>(selector));
	if (matches.length === 0) {
		return null;
	}

	const inActiveMenubar = menubar
		? matches.filter((item) => menubar.contains(item))
		: [];
	const preferred = inActiveMenubar[0] ?? matches[matches.length - 1] ?? null;
	if (!preferred) {
		return null;
	}

	for (const item of matches) {
		if (item !== preferred) {
			item.remove();
		}
	}

	return preferred;
}

/**
 * `dataset.nativePowerPoint…` maps to `data-native-power-point-…` (capital P in Point
 * inserts a hyphen). Selectors / stamps use one-word `powerpoint`. Always set/read the
 * attribute string explicitly when matching CSS or querySelector.
 */
function setDocxEditorDataAttr(element: HTMLElement, suffix: string, value = 'true') {
	element.setAttribute(`data-native-powerpoint-doc-editor-${suffix}`, value);
	// Clear the mis-mapped PowerPoint dataset form if an older build wrote it.
	element.removeAttribute(`data-native-power-point-doc-editor-${suffix}`);
}

function getDocxEditorDataAttrSelector(suffix: string): string {
	return `[data-native-powerpoint-doc-editor-${suffix}]`;
}

/** Keep one injected File/Insert menu row; drop clones from failed dedupe runs. */
function takeUniqueMenuItem(
	dropdown: HTMLElement,
	suffix: string,
	className: string,
): HTMLElement | null {
	const matches = Array.from(dropdown.querySelectorAll<HTMLElement>(
		`${getDocxEditorDataAttrSelector(suffix)}, .${className}, [data-native-power-point-doc-editor-${suffix}]`,
	)).filter((item) => item.parentElement === dropdown);

	const preferred = matches[0] ?? null;
	for (const item of matches) {
		if (item !== preferred) {
			item.remove();
		}
	}
	if (preferred) {
		setDocxEditorDataAttr(preferred, suffix);
		preferred.addClass(className);
	}
	return preferred;
}

function ensureEditorChromeMenuItemPosition(
	menubar: HTMLElement,
	item: HTMLElement,
	insertBeforeSelector?: string,
) {
	const anchor = insertBeforeSelector
		? menubar.querySelector<HTMLElement>(insertBeforeSelector)
		: null;

	if (item.parentElement !== menubar) {
		if (anchor) {
			menubar.insertBefore(item, anchor);
		} else {
			menubar.appendChild(item);
		}
		return;
	}

	if (anchor && item.nextElementSibling !== anchor) {
		menubar.insertBefore(item, anchor);
	}
}

function shouldSkipEditorOptionControl(element: HTMLElement): boolean {
	return Boolean(
		element.closest('.native-powerpoint-doc-editor-option-search-menu')
		|| element.closest('.native-powerpoint-doc-editor-edit-menu')
		|| element.closest('[data-native-powerpoint-doc-editor-edit-menu-item]')
		|| element.closest('[data-native-powerpoint-doc-editor-search-menu-item]')
		|| element.closest('[data-native-powerpoint-doc-editor-settings-menu-item]')
		|| element.closest('[data-native-powerpoint-doc-editor-no-toolbar-tooltip]')
		|| element.closest(`${getDocxEditorChromeRegionSelector('header')} input`)
		|| isTopLevelEditorMenuButton(element)
	);
}

function isSearchableEditorOptionLabel(label: string): boolean {
	const normalizedLabel = normalizeMenuText(label);
	return normalizedLabel.length > 1 && !['file', 'edit', 'format', 'insert', 'search', 'settings', 'help'].includes(normalizedLabel);
}

function isVisibleEditorOptionControl(element: HTMLElement): boolean {
	const style = window.getComputedStyle(element);
	return style.display !== 'none'
		&& style.visibility !== 'hidden'
		&& style.pointerEvents !== 'none'
		&& element.getClientRects().length > 0;
}

function textStartsWithMenuLabel(text: string, label: string): boolean {
	const normalizedLabel = normalizeMenuText(label);
	if (!normalizedLabel || !text.startsWith(normalizedLabel)) {
		return false;
	}

	const suffix = text.slice(normalizedLabel.length);
	return suffix === '' || /^\s|^ctrl|^cmd|^⌘/.test(suffix);
}

function isPrimaryFindShortcut(evt: KeyboardEvent): boolean {
	const key = evt.key.toLowerCase();
	const isMacFind = evt.metaKey && !evt.ctrlKey;
	const isNonMacFind = evt.ctrlKey && !evt.metaKey && !Platform.isMacOS;
	const hasPrimaryModifier = isMacFind || isNonMacFind;
	return key === 'f' && hasPrimaryModifier && !evt.altKey && !evt.shiftKey;
}

function getEditorMenuLabels(i18n: I18nService) {
	return {
		file: normalizeMenuText(i18n.t('docx:toolbar.file')),
		edit: 'edit',
		format: normalizeMenuText(i18n.t('docx:toolbar.format')),
		insert: normalizeMenuText(i18n.t('docx:toolbar.insert')),
		help: normalizeMenuText(i18n.t('docx:toolbar.help')),
		save: [
			i18n.t('docx:toolbar.save'),
			i18n.t('common:actions.save'),
		],
		pageSetup: [i18n.t('docx:toolbar.pageSetup'), 'Page setup'],
		pageBreak: [i18n.t('docx:toolbar.pageBreak'), 'Page break'],
		tableOfContents: [i18n.t('docx:toolbar.tableOfContents'), 'Table of contents'],
		leftToRight: [i18n.t('docx:toolbar.leftToRight'), 'Left to right'],
		rightToLeft: [i18n.t('docx:toolbar.rightToLeft'), 'Right to left'],
	};
}

function shouldHandleEditorSaveClick(target: EventTarget | null, saveLabels: string[]) {
	if (!isElement(target)) {
		return false;
	}

	let candidate: Element | null = target;
	while (candidate && candidate !== activeDocument.body) {
		if (isHTMLElement(candidate)) {
			const text = normalizeMenuText(candidate.textContent ?? '');
			if (saveLabels.some((label) => textStartsWithMenuLabel(text, label))) {
				return true;
			}
		}

		candidate = candidate.parentElement;
	}

	return false;
}

export class DocxView extends FileView {
	private hostEl: HTMLDivElement | null = null;
	private reactMount: DocxReactMount | null = null;
	private reactMountLoading = false;
	private editorMountScheduled = false;
	private activeLeafListenerRegistered = false;
	private buffer: ArrayBuffer | null = null;
	private error: string | null = null;
	private isLoading = false;
	private isDirty = false;
	private documentSession = 0;
	private readonly agentReloadGuard = new DocxAgentReloadGuard();
	private lastKnownFileSignature: DocxFileSignature | null = null;
	private backupCreatedForOpenFile = false;
	private reserveReviewSidebar = false;
	private hostResizeObserver: ResizeObserver | null = null;
	private editorAdapter: DocxEditorAdapterController | null = null;
	private editorChromeObserver: Disposable | null = null;
	private editorChromeSyncQueued = false;
	private editorChromeSyncing = false;
	private optionSearchPopoverEl: HTMLElement | null = null;
	private optionSearchCleanup: (() => void) | null = null;
	private editorEditPopoverEl: HTMLElement | null = null;
	private editorEditCleanup: (() => void) | null = null;
	private openLoadTrace: LoadTrace | null = null;
	private stopOpenHeartbeat: (() => void) | null = null;
	private editorChromeObserversRegistered = false;
	private lastEditorChromeReconciledSession = -1;

	constructor(
		leaf: WorkspaceLeaf,
		private getAuthorName: () => string,
		private getEditorTheme: () => EditorThemePreference,
		private getResolvedEditorTheme: () => EditorThemeResolution,
		private getEditorLocale: () => Translations | undefined,
		private getPluginI18n: () => I18nService | null,
		private getShowRuler: () => boolean,
		private getAutosave: () => boolean,
		private getCreateBackupsBeforeSave: () => boolean,
		private getDefaultZoom: () => number,
		private onWordCountChange: (wordCount: DocumentWordCount) => void,
		private onWordCountClear: () => void,
	) {
		super(leaf);
	}

	getViewType() {
		return VIEW_TYPE_DOCX;
	}

	getDisplayText() {
		return this.file?.basename ?? 'DOCX';
	}

	private resolveEditorMenuLabels() {
		const i18n = this.getPluginI18n();
		if (!i18n) {
			throw new Error('Plugin i18n is not initialized');
		}
		return getEditorMenuLabels(i18n);
	}

	private showNotice(key: string, values?: Record<string, string | number | boolean>) {
		showI18nNotice(this.getPluginI18n(), key, values);
	}

	private docxT(key: string): string {
		return this.getPluginI18n()?.t(key) ?? key;
	}

	getIcon() {
		return 'file-text';
	}

	canAcceptExtension(extension: string) {
		return extension.toLowerCase() === 'docx';
	}

	private beginDocumentSession(): number {
		this.documentSession += 1;
		return this.documentSession;
	}

	private createSaveOrigin(file = this.file): DocxSaveOrigin {
		if (!file) {
			throw new Error('No docx file is open.');
		}
		return { documentSession: this.documentSession, filePath: file.path };
	}

	private isCurrentSaveOrigin(origin: DocxSaveOrigin): boolean {
		return this.documentSession === origin.documentSession && this.file?.path === origin.filePath;
	}

	private assertCurrentSaveOrigin(origin: DocxSaveOrigin): void {
		if (!this.isCurrentSaveOrigin(origin)) {
			throw new Error('Discarded save from a stale DOCX editor session.');
		}
	}

	private failAgentReload(origin: DocxSaveOrigin, error: unknown): void {
		if (!this.isCurrentSaveOrigin(origin)) {
			return;
		}
		const message = error instanceof Error ? error.message : String(error);
		this.agentReloadGuard.fail(origin, error instanceof Error ? error : new Error(message));
		this.error = `Could not safely reload ${this.file?.name ?? 'the DOCX'} after the AI edit: ${message}`;
		this.isLoading = false;
		this.isDirty = false;
		this.reactMount?.unmount();
		this.reactMount = null;
		this.reactMountLoading = false;
		this.hostEl?.empty();
		this.hostEl?.createDiv({ cls: 'native-powerpoint-doc-editor-editor-load-error', text: this.error });
		this.finishOpenLoadTrace('agent-reload-failed', {
			file: origin.filePath,
			message,
		});
		errorLog('agent', this.error, error);
	}

	private beginOpenLoadTrace(phase: string, data?: Record<string, unknown>) {
		if (!this.openLoadTrace) {
			this.openLoadTrace = createLoadTrace('docx-open', {
				file: this.file?.path,
				documentSession: this.documentSession,
			});
			this.startOpenHeartbeat();
		}

		this.openLoadTrace.mark(phase, data);
	}

	private markOpenLoadPhase(phase: string, data?: Record<string, unknown>) {
		this.openLoadTrace?.mark(phase, data);
	}

	private finishOpenLoadTrace(message: string, data?: Record<string, unknown>) {
		this.openLoadTrace?.finish(message, data);
		this.openLoadTrace = null;
		this.stopOpenHeartbeat?.();
		this.stopOpenHeartbeat = null;
	}

	private startOpenHeartbeat() {
		this.stopOpenHeartbeat?.();
		this.stopOpenHeartbeat = startOpenHeartbeat('docx-open', () => ({
			file: this.file?.path,
			documentSession: this.documentSession,
			hasReactMount: Boolean(this.reactMount),
			isLoading: this.isLoading,
			chromeObserversRegistered: this.editorChromeObserversRegistered,
		}));
	}

	async onOpen() {
		try {
			this.beginOpenLoadTrace('view-onOpen-start');
			logLifecycleStep('view-onOpen', { file: this.file?.path });
			debugLog('view', 'Opening DOCX view');
			traceSyncStep('view-onOpen:empty-content', () => {
				this.contentEl.empty();
				this.hostEl = this.contentEl.createDiv({ cls: 'native-powerpoint-doc-editor-host' });
				this.editorAdapter = createDocxEditorAdapter(this.hostEl);
			}, { file: this.file?.path });
			traceSyncStep('view-onOpen:apply-theme', () => this.applyThemeClass());
			traceSyncStep('view-onOpen:prepare-host', () => this.prepareViewHost());
			traceSyncStep('view-onOpen:host-metrics', () => this.registerHostMetrics());
			traceSyncStep('view-onOpen:save-interceptor', () => this.registerEditorSaveInterceptor());
			traceSyncStep('view-onOpen:copy-interceptor', () => this.registerEditorListAwareCopyInterceptor());
			traceSyncStep('view-onOpen:save-shortcut', () => this.registerSaveShortcut());
			traceSyncStep('view-onOpen:agent-undo-shortcut', () => this.registerAgentUndoShortcut());
			traceSyncStep('view-onOpen:find-shortcut', () => this.registerFindShortcut());
			traceSyncStep('view-onOpen:scroll-guard', () => this.registerEditorDropdownScrollGuard());
			traceSyncStep('view-onOpen:active-leaf', () => this.registerActiveLeafMounting());
			this.markOpenLoadPhase('view-onOpen-ready');
			this.render();
		} catch (openError) {
			const message = openError instanceof Error ? openError.message : String(openError);
			this.finishOpenLoadTrace('view-open-failed', {
				file: this.file?.path,
				message,
			});
			throw openError;
		}
	}

	private registerEditorChromeCustomization() {
		if (this.editorChromeObserversRegistered || !this.editorAdapter || !this.reactMount) {
			return;
		}

		this.editorChromeObserversRegistered = true;
		this.markOpenLoadPhase('editor-chrome-observers-start');
		logLifecycleStep('editor-chrome-observers:start', { file: this.file?.path });
		this.syncEditorChromeCustomizations(true);
		this.editorChromeObserver = this.editorAdapter.observeChrome(() => {
			this.scheduleEditorChromeSync();
		});
		this.register(() => {
			this.editorChromeObserver?.dispose();
			this.editorChromeObserver = null;
		});
		this.markOpenLoadPhase('editor-chrome-observers-ready');
		logLifecycleStep('editor-chrome-observers:ready', { file: this.file?.path });
	}

	private syncEditorChromeCustomizations(traceSteps = false) {
		const run = traceSteps
			? <T,>(step: string, fn: () => T) => traceSyncStep(step, fn)
			: <T,>(_step: string, fn: () => T) => fn();

		run('editor-chrome:stamp-regions', () => {
			if (this.hostEl) {
				// Stamp from vendor hooks first so root/toolbar/menubar attrs exist
				// before the rest of chrome sync queries them.
				stampDocxEditorChromeRegions(this.hostEl);
			}
		});
		run('editor-chrome:sync-vendor-dark-class', () => this.syncVendorEditorDarkClass());
		run('editor-chrome:dedupe-menu-items', () => this.dedupeEditorChromeMenuItems());
		run('editor-chrome:remove-titles', () => this.removeNativeButtonTitles());
		run('editor-chrome:sync-toolbar-tooltip-metadata', () => this.syncToolbarTooltipMetadata());
		run('editor-chrome:remove-help-menu', () => this.removeEditorHelpMenu());
		run('editor-chrome:edit-menu', () => this.addEditorEditMenuButton());
		run('editor-chrome:search-menu', () => this.addEditorSearchMenuButton());
		run('editor-chrome:settings-menu', () => this.addEditorSettingsMenuButton());
		run('editor-chrome:export-menu', () => this.addEditorFileExportAsMenuItem());
		run('editor-chrome:insert-menu', () => this.addEditorInsertMenuItems());
		run('editor-chrome:normalize-menu-items', () => this.normalizeNativeEditorMenuActionItems());
	}

	private syncVendorEditorDarkClass() {
		if (!this.hostEl) {
			return;
		}

		const wantDark = this.getResolvedEditorTheme() === 'dark';
		const roots = this.hostEl.querySelectorAll<HTMLElement>('.docx-editor-root.docx-editor, [data-testid="docx-editor"]');
		roots.forEach((root) => {
			root.classList.toggle('dark', wantDark);
			// Vendor `.dark` sets --doc-caret light for an inverted page. Obsidian
			// keeps Word-white pages (filter: none) — pin caret to document ink.
			// SelectionOverlay already paints caret as #000000; CSS also remaps --doc-caret.
			root.setCssProps({ '--doc-caret': '#000000' });
		});

		const sampleRoot = roots.item(0);
		const sampleCaret = sampleRoot?.querySelector<HTMLElement>(
			'[data-testid="caret"], [data-native-powerpoint-doc-editor-caret]',
		) ?? null;
		const rootStyles = sampleRoot ? getComputedStyle(sampleRoot) : null;
		const caretStyles = sampleCaret ? getComputedStyle(sampleCaret) : null;
		debugLog('editor', 'DOCX caret pinned to page ink', {
			file: this.file?.path,
			wantDark,
			roots: roots.length,
			docCaretVar: rootStyles?.getPropertyValue('--doc-caret').trim() || null,
			caretBackground: caretStyles?.backgroundColor || null,
			rootHasDarkClass: sampleRoot?.classList.contains('dark') ?? false,
		});
	}

	private scheduleEditorChromeSync() {
		if (this.editorChromeSyncQueued || this.editorChromeSyncing) {
			return;
		}

		this.editorChromeSyncQueued = true;
		window.requestAnimationFrame(() => {
			this.editorChromeSyncQueued = false;
			this.runEditorChromeSync();
		});
	}

	private runEditorChromeSync() {
		if (!this.hostEl || this.editorChromeSyncing) {
			return;
		}

		this.editorChromeSyncing = true;
		try {
			this.syncEditorChromeCustomizations(false);
		} finally {
			this.editorChromeSyncing = false;
		}
	}

	private reconcileEditorChromeAfterViewReady() {
		// React's root can commit after the initial mount shell and observer setup.
		// Reconcile once the vendor confirms that its editor view is ready, instead
		// of waiting for a later menu interaction to trigger a chrome mutation.
		this.runEditorChromeSync();

		if (!this.hostEl || this.lastEditorChromeReconciledSession === this.documentSession) {
			return;
		}

		this.lastEditorChromeReconciledSession = this.documentSession;
		debugLog('lifecycle', 'DOCX chrome reconciled after editor view ready', {
			file: this.file?.path,
			documentSession: this.documentSession,
			roots: this.hostEl.querySelectorAll(DOCX_EDITOR_ROOT_SELECTOR).length,
			menubars: this.hostEl.querySelectorAll(DOCX_EDITOR_MENUBAR_SELECTOR).length,
			menuButtons: this.hostEl.querySelectorAll(DOCX_EDITOR_MENU_BUTTON_SELECTOR).length,
			openMenus: this.hostEl.querySelectorAll(DOCX_EDITOR_MENU_DROPDOWN_SELECTOR).length,
		});
	}

	private registerActiveLeafMounting() {
		if (this.activeLeafListenerRegistered) {
			return;
		}

		this.activeLeafListenerRegistered = true;
		this.registerEvent(this.app.workspace.on('active-leaf-change', () => {
			// PPTX↔DOCX leaves must re-bind colorMode from live resolved theme —
			// inactive DOCX trees kept stale light chrome after system-mode switches.
			if (this.isLeafActive()) {
				this.refreshSettings();
				return;
			}
			// Await flush before remount so empty-para font/size edits land on
			// disk before another tab's focus triggers vault-buffer reload.
			void (async () => {
				try {
					await this.getReactHandle()?.flushPendingSave?.();
				} catch (error) {
					warnLog('save', 'DOCX flush on tab deactivate failed', {
						file: this.file?.path ?? null,
						error: error instanceof Error ? error.message : String(error),
					});
				}
				this.render();
			})();
		}));
		this.app.workspace.onLayoutReady(() => {
			this.render();
		});
		const cancelStartupRender = scheduleIdleWork(() => {
			this.render();
		}, { timeout: 0 });
		this.register(() => cancelStartupRender());
	}

	private isLeafActive(): boolean {
		if (this.app.workspace.getActiveViewOfType(DocxView) === this) {
			return true;
		}
		if (!this.file) {
			return false;
		}
		const activeFile = this.app.workspace.getActiveFile();
		if (!activeFile || activeFile.path !== this.file.path) {
			return false;
		}
		const docxLeaves = this.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX);
		return docxLeaves.length === 1 && docxLeaves[0]?.view === this;
	}

	private renderInactivePlaceholder() {
		if (!this.hostEl || this.reactMount) {
			return;
		}

		this.hostEl.empty();
		this.hostEl.createDiv({
			cls: 'native-powerpoint-doc-editor-editor-inactive',
			text: 'Activate this tab to load the DOCX editor.',
		});
	}

	private scheduleEditorMount() {
		if (this.editorMountScheduled || this.reactMount || this.reactMountLoading || !this.hostEl) {
			return;
		}

		this.editorMountScheduled = true;
		this.markOpenLoadPhase('editor-mount-idle-scheduled', { idleTimeoutMs: 16 });
		const cancelIdle = scheduleIdleWork(() => {
			this.editorMountScheduled = false;
			if (!this.isLeafActive() || !this.hostEl) {
				this.markOpenLoadPhase('editor-mount-idle-skipped', {
					isLeafActive: this.isLeafActive(),
					hasHost: Boolean(this.hostEl),
				});
				return;
			}

			this.markOpenLoadPhase('editor-mount-idle-fired');
			void this.ensureReactMount();
		}, { timeout: 16 });

		this.register(() => cancelIdle());
	}

	async onClose() {
		debugLog('view', 'Closing DOCX view', { file: this.file?.path });
		if (!await this.promptToSaveIfDirty()) {
			warnLog('view', 'Canceled DOCX view close because unsaved changes were kept', { file: this.file?.path });
			return;
		}
		this.finishOpenLoadTrace('view-closed-before-ready', { file: this.file?.path });
		this.beginDocumentSession();
		this.agentReloadGuard.clear(new Error('DOCX view closed before the agent reload completed.'));
		this.reactMount?.unmount();
		this.onWordCountClear();
		this.reactMount = null;
		this.reactMountLoading = false;
		this.hostResizeObserver?.disconnect();
		this.hostResizeObserver = null;
		this.editorChromeObserver?.dispose();
		this.editorChromeObserver = null;
		this.editorAdapter = null;
		this.editorChromeSyncQueued = false;
		this.editorChromeSyncing = false;
		this.closeEditorOptionSearchMenu();
		this.closeEditorEditMenu();
		this.hostEl = null;
		this.buffer = null;
		this.error = null;
		this.isDirty = false;
		this.lastKnownFileSignature = null;
		this.backupCreatedForOpenFile = false;
		this.reserveReviewSidebar = false;
		this.editorChromeObserversRegistered = false;
	}

	async onLoadFile(file: TFile) {
		logLifecycleStep('file-onLoadFile', { file: file.path, bytes: file.stat.size });
		const hadOpenLoadTrace = Boolean(this.openLoadTrace);
		this.beginOpenLoadTrace('file-load-start', {
			file: file.path,
			bytes: file.stat.size,
			mtime: file.stat.mtime,
		});
		infoLog('file', `Loading ${file.path}`, {
			mtime: file.stat.mtime,
			size: file.stat.size,
		});
		if (!await this.promptToSaveIfDirty()) {
			if (!hadOpenLoadTrace) {
				this.finishOpenLoadTrace('file-load-canceled', { file: file.path });
			}
			warnLog('file', `Canceled loading ${file.path} because the current DOCX has unsaved changes`);
			return;
		}
		const origin: DocxSaveOrigin = {
			documentSession: this.beginDocumentSession(),
			filePath: file.path,
		};
		this.agentReloadGuard.clear(new Error('A different DOCX file began loading.'));
		this.isLoading = true;
		this.onWordCountClear();
		this.error = null;
		this.buffer = null;
		this.isDirty = false;
		this.lastKnownFileSignature = null;
		this.backupCreatedForOpenFile = false;
		this.reserveReviewSidebar = false;
		this.render();

		try {
			const readStartedAt = monotonicNow();
			const sourceBuffer = await this.app.vault.readBinary(file);
			if (!this.isCurrentSaveOrigin(origin)) {
				return;
			}
			this.markOpenLoadPhase('vault-readBinary-complete', {
				bytes: sourceBuffer.byteLength,
				durationMs: Math.round((monotonicNow() - readStartedAt) * 10) / 10,
			});

			const styleStartedAt = monotonicNow();
			const styledDocument = await ensureDocxDefaultStyles(sourceBuffer);
			if (!this.isCurrentSaveOrigin(origin)) {
				return;
			}
			this.markOpenLoadPhase('ensure-default-styles-complete', {
				addedDefaultStyles: styledDocument.addedDefaultStyles,
				durationMs: Math.round((monotonicNow() - styleStartedAt) * 10) / 10,
			});

			this.buffer = styledDocument.buffer;

			const signatureStartedAt = monotonicNow();
			this.lastKnownFileSignature = await this.readFileSignature(file);
			if (!this.isCurrentSaveOrigin(origin)) {
				return;
			}
			this.markOpenLoadPhase('read-file-signature-complete', {
				durationMs: Math.round((monotonicNow() - signatureStartedAt) * 10) / 10,
				signature: this.lastKnownFileSignature,
			});
			if (styledDocument.addedDefaultStyles) {
				debugLog('file', `Added default DOCX styles for ${file.path}`, {
					sourceBytes: sourceBuffer.byteLength,
					bufferBytes: this.buffer.byteLength,
				});
			}
			infoLog('file', `Loaded ${file.path}`, {
				bytes: this.buffer.byteLength,
				signature: this.lastKnownFileSignature,
			});
			this.markOpenLoadPhase('file-load-complete');
		} catch (readError) {
			if (!this.isCurrentSaveOrigin(origin)) {
				return;
			}
			const message = readError instanceof Error ? readError.message : 'Unknown read error';
			this.error = `Could not load ${file.name}: ${message}`;
			this.finishOpenLoadTrace('file-load-failed', {
				file: file.path,
				message,
			});
			errorLog('file', this.error, readError);
			showI18nNotice(this.getPluginI18n(), this.error);
		} finally {
			if (this.isCurrentSaveOrigin(origin)) {
				this.isLoading = false;
				this.render();
			}
		}

		if (this.isCurrentSaveOrigin(origin)) {
			void this.updateReviewSidebarReservation();
		}
	}

	private handleEditorLoadPhase = (
		phase: string,
		data?: Record<string, unknown>,
		origin?: DocxSaveOrigin,
	) => {
		if (origin && !this.isCurrentSaveOrigin(origin)) {
			debugLog('load', 'Ignoring stale DOCX editor load phase', {
				phase,
				file: origin.filePath,
				documentSession: origin.documentSession,
			});
			return;
		}
		this.markOpenLoadPhase(phase, data);
		if (phase === 'editor-view-ready') {
			this.finishOpenLoadTrace('editor-view-ready', {
				file: this.file?.path,
				...data,
			});
			this.reconcileEditorChromeAfterViewReady();
		}
		if (phase === 'editor-view-ready' && origin && this.agentReloadGuard.complete(origin)) {
			debugLog('agent', 'Open DOCX editor is ready for the reloaded agent buffer', {
				file: origin.filePath,
				documentSession: origin.documentSession,
			});
		}
		if (phase === 'editor-fonts-loaded' && this.openLoadTrace) {
			this.finishOpenLoadTrace('editor-ready', {
				file: this.file?.path,
				...data,
			});
		}
	};

	async onUnloadFile(_file: TFile) {
		debugLog('file', `Unloading ${_file.path}`);
		if (!await this.promptToSaveIfDirty()) {
			warnLog('file', `Canceled unloading ${_file.path} because unsaved changes were kept`);
			return;
		}
		this.finishOpenLoadTrace('file-unloaded-before-ready', { file: _file.path });
		this.beginDocumentSession();
		this.agentReloadGuard.clear(new Error('DOCX file unloaded before the agent reload completed.'));
		this.buffer = null;
		this.error = null;
		this.isDirty = false;
		this.lastKnownFileSignature = null;
		this.backupCreatedForOpenFile = false;
		this.reserveReviewSidebar = false;
		this.render();
	}

	async onRename(file: TFile) {
		const documentSession = this.beginDocumentSession();
		this.agentReloadGuard.clear(new Error('DOCX file was renamed while the agent reload was pending.'));
		await super.onRename(file);
		this.lastKnownFileSignature = await this.readFileSignature(file);
		if (this.documentSession !== documentSession || this.file?.path !== file.path) {
			return;
		}
		infoLog('file', `File renamed or moved to ${file.path}`, {
			signature: this.lastKnownFileSignature,
		});
		this.render();
	}

	async saveCurrentDocument() {
		debugLog('save', 'Save requested', { file: this.file?.path, isLoading: this.isLoading });
		if (!this.file) {
			this.showNotice('docx:notice.noFileOpen');
			return false;
		}

		if (this.isLoading) {
			this.showNotice('docx:notice.stillLoading', { fileName: this.file.name });
			return false;
		}

		const origin = this.createSaveOrigin(this.file);
		const agentReloadBuffer = this.agentReloadGuard.getPendingBuffer(origin);
		let saved = false;
		if (agentReloadBuffer) {
			await this.saveFile(agentReloadBuffer, origin);
			saved = true;
		} else {
			saved = await this.getReactHandle()?.save() ?? false;
		}
		if (!this.isCurrentSaveOrigin(origin)) {
			return false;
		}
		if (saved) {
			this.isDirty = false;
			infoLog('save', `Save completed for ${this.file.path}`);
		} else {
			warnLog('save', `Save did not complete for ${this.file.path}`);
		}

		return saved;
	}

	/**
	 * Persist dirty DOCX state before the development hot-reloader disables the
	 * plugin. Returns false when the source file could not be updated so the
	 * caller can abort reload and keep in-memory edits alive.
	 */
	async saveBeforePluginReload(): Promise<boolean> {
		const handle = this.getReactHandle();
		const comments = handle?.getComments?.() ?? null;
		const commentCount = Array.isArray(comments) ? comments.length : null;
		debugLog('save', 'DOCX save-before-reload check', {
			file: this.file?.path ?? null,
			isDirty: this.isDirty,
			commentCount,
			hasHandle: Boolean(handle),
		});
		if (!this.isDirty && commentCount === 0) {
			return true;
		}
		if (!this.isDirty && commentCount == null && !handle) {
			return true;
		}

		debugLog('save', 'Saving DOCX before plugin reload', {
			file: this.file?.path ?? null,
			isDirty: this.isDirty,
			commentCount,
		});
		return this.saveCurrentDocument();
	}

	getLoadedDocumentPath(): string | null {
		return this.file?.path ?? null;
	}

	canAgentEdit(): boolean {
		return Boolean(this.file) && !this.isLoading && !this.error;
	}

	async exportBufferForAgent(): Promise<ArrayBuffer | null> {
		if (!this.canAgentEdit()) {
			return null;
		}
		const origin = this.createSaveOrigin();

		const agentReloadBuffer = this.agentReloadGuard.getPendingBuffer(origin);
		if (agentReloadBuffer) {
			return agentReloadBuffer;
		}

		const exported = await this.getReactHandle()?.exportBuffer({ preserveAutosave: true }) ?? null;
		if (!this.isCurrentSaveOrigin(origin)) {
			return null;
		}
		if (exported) {
			return exported;
		}

		return this.buffer ? this.buffer.slice(0) : null;
	}

	async exportRenderedPdfForAgent(): Promise<ArrayBuffer | null> {
		if (!this.canAgentEdit()) return null;
		return this.getReactHandle()?.exportRenderedPdf() ?? null;
	}

	async reloadFromAgentBuffer(buffer: ArrayBuffer): Promise<void> {
		const file = this.file;
		if (!file) throw new Error('No docx file is open.');
		const previousEditor = this.getReactHandle();
		const origin: DocxSaveOrigin = {
			documentSession: this.beginDocumentSession(),
			filePath: file.path,
		};
		this.agentReloadGuard.begin(origin, buffer);

		debugLog('agent', 'Reloading open DOCX view from agent buffer', {
			file: origin.filePath,
			bytes: buffer.byteLength,
			documentSession: origin.documentSession,
		});

		try {
			await previousEditor?.prepareForExternalReload();
			this.assertCurrentSaveOrigin(origin);

			const styledDocument = await ensureDocxDefaultStyles(buffer);
			this.assertCurrentSaveOrigin(origin);
			if (!this.agentReloadGuard.stage(origin, styledDocument.buffer)) {
				throw new Error('DOCX agent reload was superseded before preprocessing completed.');
			}

			this.buffer = styledDocument.buffer;
			this.isDirty = false;
			this.error = null;
			this.lastKnownFileSignature = await this.readFileSignature(file);
			this.assertCurrentSaveOrigin(origin);

			const needsReadyBarrier = Boolean(previousEditor && this.isLeafActive());
			if (!needsReadyBarrier) {
				this.agentReloadGuard.complete(origin);
				this.render();
				return;
			}

			const ready = this.agentReloadGuard.waitForReady(origin, AGENT_RELOAD_READY_TIMEOUT_MS);
			this.render();
			await ready;
			this.assertCurrentSaveOrigin(origin);
		} catch (error) {
			this.failAgentReload(origin, error);
			throw error;
		}
	}

	canUndoAgentEdit(): boolean {
		return Boolean(this.file && aiUndoStore.canUndo(this.file.path));
	}

	canRedoAgentEdit(): boolean {
		return Boolean(this.file && aiUndoStore.canRedo(this.file.path));
	}

	async undoAgentEdit(): Promise<boolean> {
		if (!this.file) {
			return false;
		}

		const entry = aiUndoStore.popUndo(this.file.path);
		if (!entry || entry.before.kind !== 'docx') {
			return false;
		}

		const current = await this.exportBufferForAgent();
		if (current) {
			aiUndoStore.pushRedo(this.file.path, {
				label: entry.label,
				before: { kind: 'docx', buffer: current.slice(0) },
			});
		}

		await this.reloadFromAgentBuffer(entry.before.buffer);
		return true;
	}

	async redoAgentEdit(): Promise<boolean> {
		if (!this.file) {
			return false;
		}

		const entry = aiUndoStore.popRedo(this.file.path);
		if (!entry || entry.before.kind !== 'docx') {
			return false;
		}

		const current = await this.exportBufferForAgent();
		if (current) {
			aiUndoStore.record(this.file.path, {
				label: entry.label,
				before: { kind: 'docx', buffer: current.slice(0) },
			});
		}

		await this.reloadFromAgentBuffer(entry.before.buffer);
		return true;
	}

	async saveCurrentDocumentAs() {
		const file = this.file;
		debugLog('copy', 'Save as requested', { file: file?.path });
		if (!file) {
			this.showNotice('docx:notice.openToSaveCopy');
			return false;
		}

		if (this.isLoading) {
			this.showNotice('docx:notice.stillLoading', { fileName: file.name });
			return false;
		}

		const initialPath = this.getAvailableCopyPath(file);
		const chosenPath = await new Promise<DocxPathChoice>((resolve) => {
			new DocxPathModal(
				this.app,
				'Save as',
				'Create a new DOCX in this vault. If a file with that name already exists, you can replace it or keep both.',
				initialPath,
				'Save as',
				resolve,
			).open();
		});

		if (!chosenPath) {
			return false;
		}

		return this.createCurrentDocumentCopy(chosenPath, 'Saved as');
	}

	async duplicateCurrentDocument() {
		const file = this.file;
		debugLog('copy', 'Duplicate requested', { file: file?.path });
		if (!file) {
			this.showNotice('docx:notice.openToDuplicate');
			return false;
		}

		if (this.isLoading) {
			this.showNotice('docx:notice.stillLoading', { fileName: file.name });
			return false;
		}

		return this.createCurrentDocumentCopy(this.getAvailableCopyPath(file), 'Duplicated to');
	}

	async exportCurrentDocumentAs(initialFormat: DocxExportFormatId = DEFAULT_EXPORT_FORMAT) {
		const file = this.file;
		debugLog('copy', 'Export as requested', { file: file?.path, initialFormat });
		if (!file) {
			this.showNotice('docx:notice.openToExport');
			return false;
		}

		if (this.isLoading) {
			this.showNotice('docx:notice.stillLoading', { fileName: file.name });
			return false;
		}

		const initialPath = this.getAvailableExportPath(file, initialFormat);
		const initialName = initialPath.split('/').pop() ?? initialPath;
		const choice = await new Promise<DocxExportChoice>((resolve) => {
			new DocxExportModal(
				this.app,
				initialName,
				initialFormat,
				resolve,
			).open();
		});

		if (!choice) {
			return false;
		}

		const exportPath = this.getSiblingExportPath(file, choice.name, choice.format);
		if (!exportPath) {
			this.showNotice('docx:notice.enterFileName');
			return false;
		}

		return this.createCurrentDocumentExport(exportPath, choice.format);
	}

	async findHiddenText() {
		const file = this.file;
		debugLog('security', 'Find Hidden Text requested', { file: file?.path });
		if (!file) {
			this.showNotice('docx:notice.openToScanHiddenText');
			return false;
		}

		if (this.isLoading) {
			this.showNotice('docx:notice.stillLoading', { fileName: file.name });
			return false;
		}

		try {
			const liveBuffer = await this.getReactHandle()?.exportBuffer({ preserveAutosave: true });
			const scanBuffer = liveBuffer ?? this.buffer;
			if (!scanBuffer) {
				this.showNotice('docx:notice.noDataToScan');
				return false;
			}

			const result = await findHiddenDocxText(scanBuffer);
			infoLog('security', `Hidden text scan finished for ${file.path}`, {
				findings: result.findings.length,
				partsScanned: result.partsScanned,
			});
			new HiddenTextScanModal(this.app, file.name, result.findings, result.partsScanned).open();
			if (result.findings.length > 0) {
				this.showNotice('docx:notice.hiddenTextFound', { count: result.findings.length });
			}
			return true;
		} catch (error) {
			errorLog('security', `Could not scan ${file.path} for hidden text`, error);
			const message = error instanceof Error ? error.message : 'Unknown error';
			this.showNotice('docx:notice.hiddenTextScanFailed', { message });
			return false;
		}
	}

	openFindDialog() {
		if (!this.file || this.isLoading) {
			this.showNotice('docx:notice.openLoadedToSearch');
			return;
		}

		const handle = this.getReactHandle();
		if (handle) {
			handle.openFind();
			return;
		}

		window.setTimeout(() => this.getReactHandle()?.openFind(), 50);
	}

	openFindReplaceDialog() {
		if (!this.file || this.isLoading) {
			this.showNotice('docx:notice.openLoadedToSearch');
			return;
		}

		this.getReactHandle()?.openFindReplace();
	}

	openImagePicker() {
		const editor = this.getReactHandle();
		if (!editor) {
			this.showNotice('docx:notice.editorStillLoading');
			return;
		}

		editor.openImagePicker();
	}

	openCustomTableDialog() {
		const editor = this.getReactHandle();
		if (!editor) {
			this.showNotice('docx:notice.editorStillLoading');
			return;
		}

		editor.openCustomTableDialog();
	}

	openFontPicker() {
		const editor = this.getReactHandle();
		if (!editor) {
			this.showNotice('docx:notice.editorStillLoading');
			return;
		}

		editor.openFontPicker();
	}

	setEditorMode(mode: 'editing' | 'suggesting' | 'viewing') {
		const editor = this.getReactHandle();
		if (!editor) {
			this.showNotice('docx:notice.editorStillLoading');
			return;
		}

		editor.setMode(mode);
	}

	setEditorZoom(zoom: number) {
		const editor = this.getReactHandle();
		if (!editor) {
			this.showNotice('docx:notice.editorStillLoading');
			return;
		}

		editor.setZoom(zoom);
	}

	pasteFromClipboard(preserveFormatting: boolean) {
		const editor = this.getReactHandle();
		if (!editor) {
			this.showNotice('docx:notice.editorStillLoading');
			return;
		}

		void editor.pasteFromClipboard({ preserveFormatting })
			.then((pasted) => {
				if (!pasted) {
					this.showNotice('docx:notice.nothingPasted');
				}
			})
			.catch((error) => {
				errorLog('clipboard', 'DOCX paste failed', error);
			});
	}

	private getLiveEditorOptionSearchItems(): EditorOptionSearchControlItem[] {
		const root = this.hostEl?.querySelector<HTMLElement>(DOCX_EDITOR_ROOT_SELECTOR);
		if (!root) {
			return [];
		}

		const items: EditorOptionSearchControlItem[] = [];
		const seen = new Set<string>();
		const controls = root.querySelectorAll<HTMLElement>('button, [role="button"], [role="menuitem"], [role="option"]');
		controls.forEach((element, index) => {
			if (shouldSkipEditorOptionControl(element) || !isVisibleEditorOptionControl(element)) {
				return;
			}

			const label = getEditorOptionControlLabel(element);
			if (!isSearchableEditorOptionLabel(label)) {
				return;
			}

			const key = normalizeMenuText(label);
			if (seen.has(key)) {
				return;
			}

			seen.add(key);
			items.push({
				kind: 'control',
				id: `control-live:${index}:${key}`,
				label,
				keywords: getEditorOptionControlKeywords(element, label),
				element,
			});
		});

		return items;
	}

	private getEditorOptionSearchItems(): EditorOptionSearchItem[] {
		const items: EditorOptionSearchItem[] = [...EDITOR_OPTION_SEARCH_BASE_ITEMS];
		const seen = new Set(items.map(item => normalizeMenuText(item.label)));

		for (const item of this.getLiveEditorOptionSearchItems()) {
			const key = normalizeMenuText(item.label);
			if (!seen.has(key)) {
				seen.add(key);
				items.push(item);
			}
		}

		return items;
	}

	private activateEditorControl(element: HTMLElement) {
		if (element.getAttribute('aria-disabled') === 'true' || ('disabled' in element && element.disabled === true)) {
			this.showNotice('docx:menu.optionUnavailable', { option: getEditorOptionControlLabel(element) });
			return false;
		}

		element.focus({ preventScroll: true });
		element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
		element.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
		element.click();
		return true;
	}

	private findEditorControlByLabels(labels: readonly string[]) {
		const root = this.hostEl?.querySelector<HTMLElement>(DOCX_EDITOR_ROOT_SELECTOR);
		if (!root) {
			return null;
		}

			const normalizedLabels = labels.map(label => normalizeMenuText(cleanEditorOptionLabel(label))).filter(Boolean);
			const controls = root.querySelectorAll<HTMLElement>('button, [role="button"], [role="menuitem"], [role="option"]');

			for (const element of Array.from(controls)) {
			if (shouldSkipEditorOptionControl(element) || !isVisibleEditorOptionControl(element)) {
				continue;
			}

			const candidateLabels = [
				getEditorOptionControlLabel(element),
				element.dataset.testid ?? '',
				element.textContent ?? '',
				element.dataset.nativePowerPointDocEditorNativeTitle ?? '',
				element.dataset.nativePowerPointDocEditorTooltipTitle ?? '',
			].map(label => normalizeMenuText(cleanEditorOptionLabel(label))).filter(Boolean);

			if (candidateLabels.some(candidate => normalizedLabels.some(label => (
				candidate === label
				|| candidate.startsWith(label)
				|| label.startsWith(candidate)
			)))) {
				return element;
			}
		}

		return null;
	}

	private clickEditorControlByLabels(labels: readonly string[]) {
		const control = this.findEditorControlByLabels(labels);
		if (!control) {
			this.showNotice('docx:menu.notAvailable', { option: labels[0] ?? 'That option' });
			return false;
		}

		return this.activateEditorControl(control);
	}

	private findEditorMenuButton(menuLabel: string) {
		const root = this.hostEl?.querySelector<HTMLElement>(DOCX_EDITOR_ROOT_SELECTOR);
		if (!root) {
			return null;
		}

		const normalizedMenuLabel = normalizeMenuText(menuLabel);
		for (const menuItem of Array.from(root.querySelectorAll<HTMLElement>(DOCX_EDITOR_MENU_ROOT_SELECTOR))) {
			const button = menuItem.querySelector<HTMLButtonElement>(DOCX_EDITOR_MENU_BUTTON_SELECTOR);
			const label = normalizeMenuText(button?.textContent ?? '');
			if (button && label === normalizedMenuLabel) {
				return button;
			}
		}

		return null;
	}

	private clickEditorMenuOption(menuLabel: string, optionLabels: readonly string[]) {
		const menuButton = this.findEditorMenuButton(menuLabel);
		if (!menuButton) {
			this.showNotice('docx:menu.menuUnavailable', { menu: menuLabel });
			return;
		}

		this.activateEditorControl(menuButton);
		window.setTimeout(() => {
			if (!this.clickEditorControlByLabels(optionLabels)) {
				activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
			}
		});
	}

	private runEditorOptionSearchItem(item: EditorOptionSearchItem) {
		if (item.kind === 'control') {
			this.activateEditorControl(item.element);
			return;
		}

		if (item.kind === 'control-query') {
			this.clickEditorControlByLabels(item.labels);
			return;
		}

		switch (item.actionId) {
			case 'save':
				void this.saveCurrentDocument();
				break;
			case 'save-as':
				void this.saveCurrentDocumentAs();
				break;
			case 'duplicate':
				void this.duplicateCurrentDocument();
				break;
			case 'paste':
				this.pasteFromClipboard(true);
				break;
			case 'paste-without-formatting':
				this.pasteFromClipboard(false);
				break;
			case 'export-pdf':
				void this.exportCurrentDocumentAs('pdf');
				break;
			case 'export-docx':
				void this.exportCurrentDocumentAs('docx');
				break;
			case 'export-html':
				void this.exportCurrentDocumentAs('html');
				break;
			case 'export-txt':
				void this.exportCurrentDocumentAs('txt');
				break;
			case 'export-md':
				void this.exportCurrentDocumentAs('md');
				break;
			case 'export-rtf':
				void this.exportCurrentDocumentAs('rtf');
				break;
			case 'find':
				this.openFindDialog();
				break;
			case 'find-replace':
				this.openFindReplaceDialog();
				break;
			case 'insert-image':
				this.openImagePicker();
				break;
			case 'custom-table':
				this.openCustomTableDialog();
				break;
			case 'import-font':
				this.openFontPicker();
				break;
			case 'find-hidden-text':
				void this.findHiddenText();
				break;
			case 'page-setup': {
				const labels = this.resolveEditorMenuLabels();
				this.clickEditorMenuOption(labels.file, labels.pageSetup);
				break;
			}
			case 'page-break': {
				const labels = this.resolveEditorMenuLabels();
				this.clickEditorMenuOption(labels.insert, labels.pageBreak);
				break;
			}
			case 'table-of-contents': {
				const labels = this.resolveEditorMenuLabels();
				this.clickEditorMenuOption(labels.insert, labels.tableOfContents);
				break;
			}
			case 'left-to-right': {
				const labels = this.resolveEditorMenuLabels();
				this.clickEditorMenuOption(labels.format, labels.leftToRight);
				break;
			}
			case 'right-to-left': {
				const labels = this.resolveEditorMenuLabels();
				this.clickEditorMenuOption(labels.format, labels.rightToLeft);
				break;
			}
			case 'mode-editing':
				this.setEditorMode('editing');
				break;
			case 'mode-suggesting':
				this.setEditorMode('suggesting');
				break;
			case 'mode-viewing':
				this.setEditorMode('viewing');
				break;
			case 'zoom-75':
				this.setEditorZoom(0.75);
				break;
			case 'zoom-100':
				this.setEditorZoom(1);
				break;
			case 'zoom-125':
				this.setEditorZoom(1.25);
				break;
		}
	}

	refreshSettings() {
		this.applyThemeClass();
		this.render();
	}

	private applyThemeClass() {
		if (!this.hostEl) {
			return;
		}

		this.hostEl.removeClasses([
			'native-powerpoint-doc-editor-theme-system',
			'native-powerpoint-doc-editor-theme-light',
			'native-powerpoint-doc-editor-theme-dark',
			'native-powerpoint-doc-editor-theme-resolved-light',
			'native-powerpoint-doc-editor-theme-resolved-dark',
		]);
		const editorTheme = normalizeEditorThemePreference(this.getEditorTheme());
		const resolvedTheme = this.getResolvedEditorTheme();
		this.hostEl.addClass(`native-powerpoint-doc-editor-theme-${editorTheme}`);
		this.hostEl.addClass(`native-powerpoint-doc-editor-theme-resolved-${resolvedTheme}`);
		debugLog('settings', 'DOCX host theme classes applied', {
			file: this.file?.path,
			editorTheme,
			resolvedTheme,
		});
	}

	private async updateReviewSidebarReservation() {
		const buffer = this.buffer;
		const file = this.file;
		if (!buffer || !file) {
			return;
		}

		try {
			debugLog('review', `Inspecting review markup for ${file.path}`);
			const reviewStartedAt = monotonicNow();
			const { hasReviewMarkup } = await loadDocxEditorChunk();
			const hasMarkup = await hasReviewMarkup(buffer);
			this.markOpenLoadPhase('review-markup-inspection-complete', {
				hasMarkup,
				durationMs: Math.round((monotonicNow() - reviewStartedAt) * 10) / 10,
			});
			if (buffer !== this.buffer || file !== this.file) {
				debugLog('review', `Discarded stale review markup result for ${file.path}`);
				return;
			}

			this.reserveReviewSidebar = hasMarkup;
			infoLog('review', `Review markup inspection finished for ${file.path}`, { hasMarkup });
			this.render();
		} catch (error) {
			this.reserveReviewSidebar = false;
			errorLog('review', 'Could not inspect DOCX review markup.', error);
		}
	}

	private async saveFile(buffer: ArrayBuffer, requestedOrigin?: DocxSaveOrigin) {
		const file = this.file;
		if (!file) {
			throw new Error('No docx file is open.');
		}
		const origin = requestedOrigin ?? this.createSaveOrigin(file);
		this.assertCurrentSaveOrigin(origin);

		const agentReloadVersion = this.agentReloadGuard.getVersion();
		const resolveOutputBuffer = (): { version: number; buffer: ArrayBuffer } => {
			this.assertCurrentSaveOrigin(origin);
			const pending = this.agentReloadGuard.getPendingBuffer(origin);
			if (pending) {
				return { version: this.agentReloadGuard.getVersion(), buffer: pending };
			}
			return this.agentReloadGuard.getLatestBufferAfter(agentReloadVersion, origin)
				?? { version: this.agentReloadGuard.getVersion(), buffer };
		};

		let output = resolveOutputBuffer();
		infoLog('save', `Writing ${file.path}`, { bytes: output.buffer.byteLength });
		const changedOnDisk = await this.hasFileChangedOnDisk(file);
		this.assertCurrentSaveOrigin(origin);
		if (changedOnDisk) {
			warnLog('save', `Detected external change before saving ${file.path}`);
			const choice = await this.promptForExternalChange(file.name);
			this.assertCurrentSaveOrigin(origin);
			if (choice !== 'overwrite') {
				warnLog('save', `Save canceled after external change warning for ${file.path}`);
				throw new Error('Save canceled because the file changed on disk.');
			}
			warnLog('save', `External change warning overwritten for ${file.path}`);
		}

		if (this.getCreateBackupsBeforeSave()) {
			await this.createBackupBeforeOverwrite(file);
			this.assertCurrentSaveOrigin(origin);
		}

		output = resolveOutputBuffer();
		await this.app.vault.modifyBinary(file, output.buffer);
		this.assertCurrentSaveOrigin(origin);
		while (true) {
			const newerAgentBuffer = this.agentReloadGuard.getLatestBufferAfter(output.version, origin);
			if (!newerAgentBuffer) {
				break;
			}
			this.assertCurrentSaveOrigin(origin);
			output = newerAgentBuffer;
			await this.app.vault.modifyBinary(file, output.buffer);
			this.assertCurrentSaveOrigin(origin);
		}

		this.buffer = output.buffer;
		const signature = await this.readFileSignature(file);
		this.assertCurrentSaveOrigin(origin);
		this.lastKnownFileSignature = signature;
		this.isDirty = false;
		infoLog('save', `Wrote ${file.path}`, {
			signature: this.lastKnownFileSignature,
		});
	}

	private async createCurrentDocumentCopy(path: string, successPrefix: string, options: { openFile?: boolean } = {}) {
		const normalizedPath = this.normalizeDocxPath(path, this.file?.parent?.path);
		debugLog('copy', 'Creating DOCX copy', {
			requestedPath: path,
			normalizedPath,
			openFile: options.openFile !== false,
		});
		if (!normalizedPath) {
			this.showNotice('docx:notice.enterDocxPath');
			return false;
		}

		const outputPath = await this.resolveOutputPathConflict(normalizedPath);
		if (!outputPath) {
			return false;
		}

		const editor = this.getReactHandle();
		if (!editor) {
			this.showNotice('docx:notice.editorNotReady');
			return false;
		}

		const buffer = await editor.exportBuffer();
		if (!buffer) {
			return false;
		}

		let newFile: TFile;
		try {
			await this.ensureParentFolders(outputPath.path);
			if (outputPath.existingFile) {
				await this.app.vault.modifyBinary(outputPath.existingFile, buffer);
				newFile = outputPath.existingFile;
			} else {
				newFile = await this.app.vault.createBinary(outputPath.path, buffer);
			}
			infoLog('copy', `${outputPath.replace ? 'Replaced' : 'Created'} ${newFile.path}`, { bytes: buffer.byteLength });
		} catch (copyError) {
			const message = copyError instanceof Error ? copyError.message : 'Unknown copy error';
			errorLog('copy', `Could not create ${outputPath.path}`, copyError);
			this.showNotice('errors:createFailed', { path: outputPath.path, message });
			return false;
		}

		if (options.openFile !== false) {
			const wasDirty = this.isDirty;
			try {
				this.isDirty = false;
				await this.leaf.openFile(newFile);
			} catch (openError) {
				this.isDirty = wasDirty;
				const message = openError instanceof Error ? openError.message : 'Unknown open error';
				errorLog('copy', `Created ${newFile.path}, but could not open it`, openError);
				this.showNotice('docx:notice.createdButCouldNotOpen', { path: newFile.path, message });
				return true;
			}
		}

		const successKey = outputPath.replace
			? 'docx:notice.replaced'
			: successPrefix === 'Saved as'
				? 'docx:notice.savedAs'
				: 'docx:notice.duplicatedTo';
		this.showNotice(successKey, { path: newFile.path });
		return true;
	}

	private async createCurrentDocumentExport(path: string, formatId: DocxExportFormatId) {
		const normalizedPath = this.normalizeExportPath(path, formatId, this.file?.parent?.path);
		debugLog('copy', 'Creating document export', {
			requestedPath: path,
			normalizedPath,
			format: formatId,
		});
		if (!normalizedPath) {
			this.showNotice('docx:notice.enterFilePath');
			return false;
		}

		const outputPath = await this.resolveOutputPathConflict(normalizedPath);
		if (!outputPath) {
			return false;
		}

		const editor = this.getReactHandle();
		if (!editor) {
			this.showNotice('docx:notice.editorNotReady');
			return false;
		}

		try {
			await this.ensureParentFolders(outputPath.path);
			this.showNotice('docx:notice.exportingTo', { path: outputPath.path });
			let exportContent: ArrayBuffer | ArrayBufferView | string | null = null;
			if (formatId === 'pdf') {
				exportContent = await editor.exportRenderedPdf();
				if (!exportContent) {
					warnLog('copy', 'Rendered PDF export did not finish; no PDF file was written', {
						path: outputPath.path,
					});
					this.showNotice('errors:exportPdfFailed', { path: outputPath.path });
					return false;
				}
			}
			if (!exportContent) {
				const buffer = await editor.exportBuffer();
				if (!buffer) {
					this.showNotice('errors:exportNoDocument', { path: outputPath.path });
					return false;
				}
				exportContent = await this.createExportContent(buffer, formatId, this.file?.basename ?? 'Document');
			}
			const binaryContent = getBinaryExportContent(exportContent);
			const textContent = typeof exportContent === 'string' ? exportContent : null;
			if (!binaryContent && textContent === null) {
				throw new Error('The editor returned an unsupported export payload.');
			}
			let newFile: TFile;
			if (outputPath.existingFile) {
				if (binaryContent) {
					await this.app.vault.modifyBinary(outputPath.existingFile, binaryContent);
				} else if (textContent !== null) {
					await this.app.vault.modify(outputPath.existingFile, textContent);
				}
				newFile = outputPath.existingFile;
			} else {
				if (binaryContent) {
					newFile = await this.app.vault.createBinary(outputPath.path, binaryContent);
				} else if (textContent !== null) {
					newFile = await this.app.vault.create(outputPath.path, textContent);
				} else {
					throw new Error('The editor returned an unsupported export payload.');
				}
			}
			infoLog('copy', `${outputPath.replace ? 'Replaced' : 'Exported'} ${newFile.path}`, { format: formatId });
			this.showNotice(outputPath.replace ? 'docx:notice.replaced' : 'docx:notice.exportedTo', { path: newFile.path });
			return true;
		} catch (exportError) {
			const message = exportError instanceof Error ? exportError.message : 'Unknown export error';
			errorLog('copy', `Could not export ${outputPath.path}`, exportError);
			this.showNotice('errors:exportFailed', { path: outputPath.path, message });
			return false;
		}
	}

	private async createExportContent(buffer: ArrayBuffer, formatId: DocxExportFormatId, title: string): Promise<ArrayBuffer | string> {
		if (formatId === 'docx') {
			return buffer;
		}

		if (formatId === 'md') {
			const markdown = await extractDocxMarkdown(buffer);
			return `${markdown || title}\n`;
		}

		const text = await extractDocxText(buffer);
		const exportText = text || title;

		switch (formatId) {
			case 'html':
				return createPlainTextHtml(exportText, title);
			case 'txt':
				return `${exportText}\n`;
			case 'rtf':
				return createPlainTextRtf(exportText);
			default:
				throw new Error(`Unsupported export format: ${formatId}`);
		}
	}

	private async resolveOutputPathConflict(path: string): Promise<{ path: string; existingFile: TFile | null; replace: boolean } | null> {
		const existingFile = this.app.vault.getAbstractFileByPath(path);
		if (!existingFile) {
			return { path, existingFile: null, replace: false };
		}

		if (!(existingFile instanceof TFile)) {
			this.showNotice('docx:notice.pathExistsNotFile', { path });
			return null;
		}

		const choice = await new Promise<ExistingFileChoice>((resolve) => {
			new ExistingFileModal(this.app, path, resolve).open();
		});

		const canonicalChoice: ArtifactConflictChoice = choice === 'replace' ? 'replace' : 'keep-both';
		const resolved = resolveArtifactConflict(path, existingFile, canonicalChoice, (candidate) =>
			Boolean(this.app.vault.getAbstractFileByPath(candidate)),
		);
		if (resolved && !resolved.replace) {
			infoLog('copy', `Keeping existing file and writing numbered copy`, {
				originalPath: path,
				copyPath: resolved.path,
			});
		}
		return resolved;
	}

	private getAvailableNumberedPath(path: string) {
		return resolveNumberedArtifactPath(path, (candidate) => Boolean(this.app.vault.getAbstractFileByPath(candidate)));
	}

	private async renameFile(name: string, expectedPath?: string | null) {
		const file = this.file;
		if (!file) {
			throw new Error('No docx file is open.');
		}
		if (expectedPath && file.path !== expectedPath) {
			debugLog('file', `Discarded stale DOCX title rename for ${expectedPath}`, {
				currentPath: file.path,
				requestedName: name,
			});
			return;
		}

		const normalizedName = this.normalizeDocxFileName(name);
		if (!normalizedName) {
			throw new Error('Document name cannot be empty.');
		}

		if (normalizedName === file.name) {
			return;
		}

		const previousPath = file.path;
		const result = await renameFileToSiblingName(this.app, file, normalizedName);
		infoLog('file', `Renamed ${previousPath} to ${result.path}`);
		this.lastKnownFileSignature = await this.readFileSignature(file);
		this.showNotice('docx:notice.renamedTo', { name: normalizedName });
	}

	private async readFileSignature(file: TFile): Promise<DocxFileSignature> {
		try {
			const stat = await this.app.vault.adapter.stat(file.path);
			if (stat?.type === 'file') {
				return {
					path: file.path,
					mtime: stat.mtime,
					size: stat.size,
				};
			}
		} catch (error) {
			debugLog('file', `Falling back to cached file stat for ${file.path}`, error);
			// Fall back to Obsidian's cached stat below.
		}

		return {
			path: file.path,
			mtime: file.stat.mtime,
			size: file.stat.size,
		};
	}

	private signaturesMatch(a: DocxFileSignature, b: DocxFileSignature) {
		return a.path === b.path && a.mtime === b.mtime && a.size === b.size;
	}

	private async hasFileChangedOnDisk(file: TFile) {
		if (!this.lastKnownFileSignature) {
			this.lastKnownFileSignature = await this.readFileSignature(file);
			debugLog('save', `Initialized save conflict signature for ${file.path}`, this.lastKnownFileSignature);
			return false;
		}

		const currentSignature = await this.readFileSignature(file);
		const changed = !this.signaturesMatch(this.lastKnownFileSignature, currentSignature);
		debugLog('save', `Compared disk signature for ${file.path}`, {
			changed,
			lastKnown: this.lastKnownFileSignature,
			current: currentSignature,
		});
		return changed;
	}

	private async promptForExternalChange(fileName: string) {
		return new Promise<DocxConflictChoice>((resolve) => {
			new ExternalDocxChangeModal(this.app, fileName, resolve).open();
		});
	}

	private async createBackupBeforeOverwrite(file: TFile) {
		if (this.backupCreatedForOpenFile) {
			debugLog('backup', `Backup already created for this open session: ${file.path}`);
			return;
		}

		const sourceBuffer = await this.app.vault.readBinary(file);
		const backupPath = this.getAvailableBackupPath(file);
		await this.ensureParentFolders(backupPath);
		await this.app.vault.createBinary(backupPath, sourceBuffer);
		this.backupCreatedForOpenFile = true;
		infoLog('backup', `Created backup for ${file.path}`, {
			backupPath,
			bytes: sourceBuffer.byteLength,
		});
	}

	private getAvailableBackupPath(file: TFile) {
		const backupFolder = normalizePath(`${getVaultFolderPrefix(file.parent?.path)}.native-powerpoint-doc-editor-backups`);
		const baseName = file.basename || file.name.replace(/\.docx$/i, '');
		const timestamp = new Date().toISOString().replace(/\.\d{3}Z$/, '').replace('T', ' ').replace(/:/g, '-');

		for (let index = 0; index < 1000; index += 1) {
			const suffix = index === 0 ? '' : ` ${index + 1}`;
			const candidatePath = normalizePath(`${backupFolder}/${baseName} backup ${timestamp}${suffix}.docx`);
			if (!this.app.vault.getAbstractFileByPath(candidatePath)) {
				return candidatePath;
			}
		}

		return normalizePath(`${backupFolder}/${baseName} backup ${Date.now()}.docx`);
	}

	private normalizeDocxFileName(name: string) {
		const trimmedName = name.trim().replace(/[\\/]/g, '-');
		if (!trimmedName || trimmedName === '.docx') {
			return null;
		}

		return trimmedName.toLowerCase().endsWith('.docx') ? trimmedName : `${trimmedName}.docx`;
	}

	private normalizeExportFileName(name: string, formatId: DocxExportFormatId) {
		const normalizedName = withExportExtension(name, formatId);
		return normalizedName || null;
	}

	private normalizeDocxPath(path: string, fallbackFolderPath?: string) {
		const trimmedPath = path.trim().replace(/\\/g, '/');
		if (!trimmedPath || trimmedPath === '.docx' || trimmedPath.endsWith('/')) {
			return null;
		}

		const pathWithExtension = trimmedPath.toLowerCase().endsWith('.docx') ? trimmedPath : `${trimmedPath}.docx`;
		const hasFolder = pathWithExtension.includes('/');
		const parentPath = fallbackFolderPath && fallbackFolderPath !== '/' ? fallbackFolderPath : '';
		const fullPath = parentPath && !hasFolder ? `${parentPath}/${pathWithExtension}` : pathWithExtension;
		const normalizedDocxPath = normalizePath(fullPath);
		const fileName = normalizedDocxPath.split('/').pop();

		return fileName && fileName !== '.docx' ? normalizedDocxPath : null;
	}

	private normalizeExportPath(path: string, formatId: DocxExportFormatId, fallbackFolderPath?: string) {
		const trimmedPath = path.trim().replace(/\\/g, '/');
		if (!trimmedPath || trimmedPath.endsWith('/')) {
			return null;
		}

		const pathParts = trimmedPath.split('/');
		const rawFileName = pathParts.pop() ?? '';
		const normalizedFileName = this.normalizeExportFileName(rawFileName, formatId);
		if (!normalizedFileName) {
			return null;
		}

		const pathWithExtension = [...pathParts, normalizedFileName].filter(Boolean).join('/');
		const hasFolder = pathWithExtension.includes('/');
		const parentPath = fallbackFolderPath && fallbackFolderPath !== '/' ? fallbackFolderPath : '';
		return normalizePath(parentPath && !hasFolder ? `${parentPath}/${pathWithExtension}` : pathWithExtension);
	}

	private getAvailableCopyPath(file: TFile) {
		const folderPrefix = getVaultFolderPrefix(file.parent?.path);
		const baseName = file.basename || file.name.replace(/\.docx$/i, '');

		for (let index = 1; index < 1000; index += 1) {
			const suffix = index === 1 ? 'copy' : `copy ${index}`;
			const candidatePath = normalizePath(`${folderPrefix}${baseName} ${suffix}.docx`);
			if (!this.app.vault.getAbstractFileByPath(candidatePath)) {
				return candidatePath;
			}
		}

		return normalizePath(`${folderPrefix}${baseName} copy ${Date.now()}.docx`);
	}

	private getAvailableExportPath(file: TFile, formatId: DocxExportFormatId) {
		const folderPrefix = getVaultFolderPrefix(file.parent?.path);
		const baseName = file.basename || file.name.replace(/\.docx$/i, '');
		const extension = getExportFormat(formatId).extension;
		const preferredPath = normalizePath(`${folderPrefix}${baseName}.${extension}`);

		if (preferredPath.toLowerCase() === normalizePath(file.path).toLowerCase()) {
			return this.getAvailableNumberedPath(preferredPath);
		}

		return preferredPath;
	}

	private getSiblingDocxPath(file: TFile, name: string) {
		const normalizedName = this.normalizeDocxFileName(name);
		if (!normalizedName) {
			return null;
		}

		const folderPath = file.parent?.path;
		const folderPrefix = folderPath && folderPath !== '/' ? `${folderPath}/` : '';
		return normalizePath(`${folderPrefix}${normalizedName}`);
	}

	private getSiblingExportPath(file: TFile, name: string, formatId: DocxExportFormatId) {
		const normalizedName = this.normalizeExportFileName(name, formatId);
		if (!normalizedName) {
			return null;
		}

		return buildSiblingPath(file, normalizedName);
	}

	private async ensureParentFolders(path: string) {
		const folderPath = path.split('/').slice(0, -1).join('/');
		if (!folderPath) {
			return;
		}

		const segments = folderPath.split('/').filter(Boolean);
		let currentPath = '';

		for (const segment of segments) {
			currentPath = currentPath ? `${currentPath}/${segment}` : segment;
			const existingFile = this.app.vault.getAbstractFileByPath(currentPath);

			if (existingFile instanceof TFile) {
				throw new Error(`${currentPath} is a file.`);
			}

			if (!existingFile) {
				await this.app.vault.createFolder(currentPath);
			}
		}
	}

	private async promptToSaveIfDirty(): Promise<boolean> {
		// Close often races an in-flight autosave (`state === 'saving'` still counts
		// as dirty). Flush first so a completed write does not show the modal.
		const handle = this.getReactHandle();
		if (handle?.flushPendingSave) {
			try {
				await handle.flushPendingSave();
			} catch (error) {
				warnLog('save', 'DOCX flush before close failed', {
					file: this.file?.path ?? null,
					error: error instanceof Error ? error.message : String(error),
				});
			}
		}

		if (!this.isDirty || !this.file) {
			return true;
		}

		warnLog('save', `Prompting for unsaved changes in ${this.file.path}`);
		const choice = await new Promise<UnsavedDocxChoice>((resolve) => {
			new UnsavedDocxModal(this.app, this.file?.name ?? 'Document', resolve).open();
		});

		if (choice === 'save') {
			infoLog('save', `Saving dirty document before closing ${this.file.path}`);
			const saved = await this.saveCurrentDocument();
			if (!saved) {
				warnLog('save', `Keeping ${this.file.path} open because save did not complete`);
				this.showNotice('docx:notice.saveIncomplete');
				return false;
			}
			return true;
		}
		if (choice === 'discard') {
			warnLog('save', `Discarding unsaved changes in ${this.file.path}`);
			return true;
		}

		warnLog('save', `Canceled close for unsaved DOCX ${this.file.path}`);
		return false;
	}

	private prepareViewHost() {
		if (!this.hostEl) {
			return;
		}

		this.hostEl.setCssProps({
			'--native-powerpoint-doc-editor-fixed-left-offset': '0px',
			'--native-powerpoint-doc-editor-fixed-top-offset': '0px',
		});
		this.hostEl.setAttribute('data-native-powerpoint-doc-editor-toolbar-tooltips', 'custom');
	}

	private syncToolbarTooltipMetadata() {
		if (!this.hostEl) {
			return;
		}

		neutralizeToolbarButtonTooltipSources(this.hostEl);
	}

	private registerHostMetrics() {
		const updateHostMetrics = () => {
			if (!this.hostEl) {
				return;
			}

			const fixedProbe = this.hostEl.createDiv({ cls: 'native-powerpoint-doc-editor-fixed-probe' });
			const fixedRect = fixedProbe.getBoundingClientRect();
			fixedProbe.remove();

			this.hostEl.setCssProps({
				'--native-powerpoint-doc-editor-fixed-left-offset': `${Math.round(fixedRect.left)}px`,
				'--native-powerpoint-doc-editor-fixed-top-offset': `${Math.round(fixedRect.top)}px`,
			});
		};

		updateHostMetrics();
		this.registerDomEvent(window, 'resize', updateHostMetrics);
		this.registerDomEvent(window, 'scroll', updateHostMetrics, true);
		this.hostResizeObserver = new ResizeObserver(updateHostMetrics);
		this.hostResizeObserver.observe(this.contentEl);
		this.register(() => {
			this.hostResizeObserver?.disconnect();
			this.hostResizeObserver = null;
		});
	}

	private removeNativeButtonTitles() {
		if (!this.hostEl) {
			return;
		}

		this.hostEl?.querySelectorAll(`${DOCX_EDITOR_ROOT_SELECTOR} button[title]`).forEach((button) => {
			if (isHTMLElement(button)) {
				const title = button.getAttribute('title');
				if (title) {
					button.dataset.nativePowerPointDocEditorNativeTitle = title;
				}
			}
			button.removeAttribute('title');
		});
	}

	private removeEditorHelpMenu() {
		if (!this.hostEl) {
			return;
		}

		const helpLabel = this.resolveEditorMenuLabels().help;

		this.hostEl?.querySelectorAll(DOCX_EDITOR_MENU_ROOT_SELECTOR).forEach((menuItem) => {
			const button = menuItem.querySelector(DOCX_EDITOR_MENU_BUTTON_SELECTOR);
			const label = normalizeMenuText(button?.textContent ?? '');
			if (label === helpLabel) {
				menuItem.remove();
			}
		});
	}

	private dedupeEditorChromeMenuItems() {
		if (!this.hostEl) {
			return;
		}

		for (const item of Object.values(EDITOR_CHROME_MENU_ITEMS)) {
			dedupeEditorChromeMenuItem(this.hostEl, item.selector);
		}
	}

	private createDetachedEditorChromeElement<K extends keyof HTMLElementTagNameMap>(
		parent: HTMLElement,
		tagName: K,
		menu: string,
	): HTMLElementTagNameMap[K] {
		const element = createDetachedDocxEditorChromeElement(parent, tagName);
		debugLog('view', 'DOCX chrome fallback element created', {
			file: this.file?.path,
			menu,
			tagName,
			parentTagName: parent.tagName,
			parentConnected: parent.isConnected,
			ownerDocumentMatchesHost: parent.ownerDocument === this.hostEl?.ownerDocument,
		});
		return element;
	}

	private addEditorEditMenuButton() {
		if (!this.hostEl) {
			return;
		}

		const labels = this.resolveEditorMenuLabels();
			const menubar = getActiveEditorMenubar(this.hostEl);
			if (!menubar) {
				return;
			}

			const existingEditItem = dedupeEditorChromeMenuItem(
				this.hostEl,
				EDITOR_CHROME_MENU_ITEMS.edit.selector,
			);
			if (existingEditItem) {
				markEditorChromeMenuItem(existingEditItem, 'edit');
				return;
			}

			const menuItems = Array.from(menubar.children).filter((child): child is HTMLElement => isHTMLElement(child));
			const findTopLevelMenu = (label: string) => menuItems.find((item) => {
				const button = item.querySelector(DOCX_EDITOR_MENU_BUTTON_SELECTOR);
				return normalizeMenuText(button?.textContent ?? '') === label;
			});
			const fileWrapper = findTopLevelMenu(labels.file);
			const formatWrapper = findTopLevelMenu(labels.format);
			const existingEditWrapper = findTopLevelMenu(labels.edit);
			const sourceWrapper = existingEditWrapper
				?? menuItems.find((child) => (
					!child.matches('[data-native-powerpoint-doc-editor-edit-menu-item], [data-native-powerpoint-doc-editor-search-menu-item], [data-native-powerpoint-doc-editor-settings-menu-item]')
					&& Boolean(child.querySelector(DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR))
				));
			const wrapper = existingEditWrapper
				?? (sourceWrapper ? sourceWrapper.cloneNode(true) as HTMLElement : this.createDetachedEditorChromeElement(menubar, 'div', 'edit'));
			markEditorChromeMenuItem(wrapper, 'edit');
			wrapper.addClass('native-powerpoint-doc-editor-edit-menu-item');
			wrapper.setCssProps({ position: 'relative' });
			wrapper.removeAttribute('id');

			let button = wrapper.querySelector<HTMLButtonElement>(DOCX_EDITOR_MENU_BUTTON_SELECTOR);
			Array.from(wrapper.children).forEach((child) => {
				if (child !== button) {
					child.remove();
				}
			});
			if (!button) {
				button = this.createDetachedEditorChromeElement(wrapper, 'button', 'edit');
				wrapper.appendChild(button);
			}

			button.type = 'button';
			button.textContent = this.docxT('docx:chrome.edit');
			markEditorChromeMenuButton(button);
			markEditorChromeNoToolbarTooltip(button);
			button.addClasses(['native-powerpoint-doc-editor-search-menu-button', 'native-powerpoint-doc-editor-edit-menu-button']);
			button.removeAttribute('aria-haspopup');
			button.removeAttribute('data-state');
			button.removeAttribute('id');
			button.setAttribute('aria-label', this.docxT('docx:chrome.edit'));
			button.setAttribute('aria-expanded', 'false');
			button.setAttribute('role', 'menuitem');
			button.addEventListener('mousedown', (evt) => {
				evt.preventDefault();
			});
			button.addEventListener('click', (evt) => {
				evt.preventDefault();
				evt.stopImmediatePropagation();
				evt.stopPropagation();
				this.openEditorEditMenu(wrapper);
				button?.setAttribute('aria-expanded', this.editorEditPopoverEl ? 'true' : 'false');
			});

		if (fileWrapper && fileWrapper.parentElement === menubar) {
			fileWrapper.after(wrapper);
		} else if (formatWrapper && formatWrapper.parentElement === menubar) {
			menubar.insertBefore(wrapper, formatWrapper);
		} else {
			menubar.prepend(wrapper);
		}
	}

	private closeEditorEditMenu() {
		this.editorEditCleanup?.();
		this.editorEditCleanup = null;
		this.editorEditPopoverEl?.remove();
		this.editorEditPopoverEl = null;
		this.hostEl?.querySelector('[data-native-powerpoint-doc-editor-edit-menu-item] > button')?.setAttribute('aria-expanded', 'false');
	}

	private openEditorEditMenu(anchorEl: HTMLElement) {
		if (this.editorEditPopoverEl && anchorEl.contains(this.editorEditPopoverEl)) {
			this.closeEditorEditMenu();
			return;
		}

		this.closeEditorOptionSearchMenu();
		this.closeEditorEditMenu();

		const popoverEl = createPopoverShell(anchorEl, {
			className: 'native-powerpoint-doc-editor-edit-menu native-powerpoint-doc-editor-option-search-menu',
			role: 'menu',
		});
		this.editorEditPopoverEl = popoverEl;

		const addAction = (label: string, preserveFormatting: boolean) => {
			createMenuItem(popoverEl, {
				className: 'native-powerpoint-doc-editor-option-search-result native-powerpoint-doc-editor-file-menu-button',
				text: label,
				role: 'menuitem',
				preventMouseDown: true,
				preventDefaultOnClick: true,
				stopClickPropagation: true,
				onClick: () => {
				this.closeEditorEditMenu();
				this.pasteFromClipboard(preserveFormatting);
				},
			});
		};

		addAction('Paste', true);
		addAction('Paste without formatting', false);

		this.editorEditCleanup = bindPopoverDismissHandlers({
			popover: popoverEl,
			anchor: anchorEl,
			onDismiss: () => this.closeEditorEditMenu(),
			pointerEvent: 'mousedown',
		});
	}

	private addEditorSearchMenuButton() {
		if (!this.hostEl) {
			return;
		}

		const menubar = getActiveEditorMenubar(this.hostEl);
			if (!menubar) {
				return;
			}

			const existingSearchItem = dedupeEditorChromeMenuItem(
				this.hostEl,
				EDITOR_CHROME_MENU_ITEMS.search.selector,
			);
			if (existingSearchItem) {
				markEditorChromeMenuItem(existingSearchItem, 'search');
				ensureEditorChromeMenuItemPosition(
					menubar,
					existingSearchItem,
					EDITOR_CHROME_MENU_ITEMS.settings.selector,
				);
				return;
			}

				const sourceWrapper = Array.from(menubar.children).find((child): child is HTMLElement => (
					isHTMLElement(child)
					&& !child.matches('[data-native-powerpoint-doc-editor-edit-menu-item], [data-native-powerpoint-doc-editor-search-menu-item], [data-native-powerpoint-doc-editor-settings-menu-item]')
					&& Boolean(child.querySelector(DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR))
				));
				const wrapper = sourceWrapper
					? sourceWrapper.cloneNode(true) as HTMLElement
					: this.createDetachedEditorChromeElement(menubar, 'div', 'search');
				markEditorChromeMenuItem(wrapper, 'search');
				wrapper.addClass('native-powerpoint-doc-editor-search-menu-item');
				wrapper.setCssProps({ position: 'relative' });
				wrapper.removeAttribute('id');

				let button = wrapper.querySelector<HTMLButtonElement>(DOCX_EDITOR_MENU_BUTTON_SELECTOR);
				Array.from(wrapper.children).forEach((child) => {
					if (child !== button) {
						child.remove();
					}
				});
				if (!button) {
					button = this.createDetachedEditorChromeElement(wrapper, 'button', 'search');
					wrapper.appendChild(button);
				}
				button.type = 'button';
				button.textContent = this.docxT('docx:chrome.search');
				markEditorChromeMenuButton(button);
			markEditorChromeNoToolbarTooltip(button);
				button.addClass('native-powerpoint-doc-editor-search-menu-button');
				button.removeAttribute('aria-expanded');
				button.removeAttribute('aria-haspopup');
				button.removeAttribute('data-state');
				button.removeAttribute('id');
				button.setAttribute('aria-label', this.docxT('docx:chrome.search'));
			button.setAttribute('role', 'menuitem');
			button.addEventListener('mousedown', (evt) => {
				evt.preventDefault();
			});
			button.addEventListener('click', (evt) => {
					evt.preventDefault();
					evt.stopImmediatePropagation();
					evt.stopPropagation();
					this.openEditorOptionSearchMenu(wrapper);
				});

			const settingsWrapper = menubar.querySelector('[data-native-powerpoint-doc-editor-settings-menu-item]');
			if (settingsWrapper) {
				menubar.insertBefore(wrapper, settingsWrapper);
			} else {
				menubar.appendChild(wrapper);
			}
	}

		private closeEditorOptionSearchMenu() {
			this.optionSearchCleanup?.();
			this.optionSearchCleanup = null;
			this.optionSearchPopoverEl?.remove();
			this.optionSearchPopoverEl = null;
		}

		private openEditorOptionSearchMenu(anchorEl: HTMLElement) {
			if (this.optionSearchPopoverEl && anchorEl.contains(this.optionSearchPopoverEl)) {
				this.closeEditorOptionSearchMenu();
				return;
			}

			this.closeEditorOptionSearchMenu();
			this.closeEditorEditMenu();

			let activeIndex = 0;
			const popoverEl = createPopoverShell(anchorEl, { className: 'native-powerpoint-doc-editor-option-search-menu' });
			const inputEl = popoverEl.createEl('input', {
				cls: 'native-powerpoint-doc-editor-option-search-input',
				type: 'search',
			});
			inputEl.placeholder = this.docxT('docx:chrome.searchOptionsPlaceholder');
			inputEl.setAttribute('aria-label', this.docxT('docx:chrome.searchOptions'));

			const resultsEl = popoverEl.createDiv({ cls: 'native-powerpoint-doc-editor-option-search-results' });
			this.optionSearchPopoverEl = popoverEl;

			const getMatches = () => {
				const items = this.getEditorOptionSearchItems();
				const query = normalizeMenuText(inputEl.value);
				if (!query) {
					return items;
				}

				return items.filter((item) => {
					const haystack = normalizeMenuText([item.label, ...item.keywords].join(' '));
					return haystack.includes(query);
				});
			};

			const chooseItem = (item: EditorOptionSearchItem) => {
				this.closeEditorOptionSearchMenu();
				this.runEditorOptionSearchItem(item);
			};

			const renderResults = () => {
				const matches = getMatches();
				activeIndex = Math.max(0, Math.min(activeIndex, matches.length - 1));
				resultsEl.empty();

				if (matches.length === 0) {
					resultsEl.createDiv({ cls: 'native-powerpoint-doc-editor-option-search-empty', text: 'No options found' });
					return;
				}

				matches.forEach((item, index) => {
					createMenuItem(resultsEl, {
						className: 'native-powerpoint-doc-editor-option-search-result native-powerpoint-doc-editor-file-menu-button',
						text: item.label,
						role: 'option',
						selected: index === activeIndex,
						preventDefaultOnClick: true,
						stopClickPropagation: true,
						onMouseEnter: () => {
							activeIndex = index;
							renderResults();
						},
						onClick: () => chooseItem(item),
					});
				});
			};

			const handleInput = () => {
				activeIndex = 0;
				renderResults();
			};
			const handleKeyDown = (evt: KeyboardEvent) => {
				const matches = getMatches();
				if (evt.key === 'Escape') {
					evt.preventDefault();
					this.closeEditorOptionSearchMenu();
					return;
				}
				if (evt.key === 'ArrowDown') {
					evt.preventDefault();
					activeIndex = matches.length > 0 ? (activeIndex + 1) % matches.length : 0;
					renderResults();
					return;
				}
				if (evt.key === 'ArrowUp') {
					evt.preventDefault();
					activeIndex = matches.length > 0 ? (activeIndex - 1 + matches.length) % matches.length : 0;
					renderResults();
					return;
				}
				if (evt.key === 'Enter') {
					evt.preventDefault();
					const item = matches[activeIndex];
					if (item) {
						chooseItem(item);
					}
				}
			};
			const dismissCleanup = bindPopoverDismissHandlers({
				popover: popoverEl,
				anchor: anchorEl,
				onDismiss: () => this.closeEditorOptionSearchMenu(),
				pointerEvent: 'mousedown',
				closeOnEscape: false,
			});

			inputEl.addEventListener('input', handleInput);
			inputEl.addEventListener('keydown', handleKeyDown);
			this.optionSearchCleanup = () => {
				inputEl.removeEventListener('input', handleInput);
				inputEl.removeEventListener('keydown', handleKeyDown);
				dismissCleanup();
			};

			renderResults();
			window.setTimeout(() => inputEl.focus());
		}

	private addEditorSettingsMenuButton() {
		if (!this.hostEl) {
			return;
		}

		const menubar = getActiveEditorMenubar(this.hostEl);
				if (!menubar) {
					return;
				}

				const existingSettingsItem = dedupeEditorChromeMenuItem(
					this.hostEl,
					EDITOR_CHROME_MENU_ITEMS.settings.selector,
				);
				if (existingSettingsItem) {
					markEditorChromeMenuItem(existingSettingsItem, 'settings');
					ensureEditorChromeMenuItemPosition(menubar, existingSettingsItem);
					return;
				}

				const sourceWrapper = Array.from(menubar.children).find((child): child is HTMLElement => (
					isHTMLElement(child)
					&& !child.matches('[data-native-powerpoint-doc-editor-edit-menu-item], [data-native-powerpoint-doc-editor-search-menu-item], [data-native-powerpoint-doc-editor-settings-menu-item]')
					&& Boolean(child.querySelector(DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR))
				));
				const wrapper = sourceWrapper
					? sourceWrapper.cloneNode(true) as HTMLElement
					: this.createDetachedEditorChromeElement(menubar, 'div', 'settings');
				markEditorChromeMenuItem(wrapper, 'settings');
				wrapper.addClass('native-powerpoint-doc-editor-settings-menu-item');
				wrapper.setCssProps({ position: 'relative' });
				wrapper.removeAttribute('id');

				let button = wrapper.querySelector<HTMLButtonElement>(DOCX_EDITOR_MENU_BUTTON_SELECTOR);
				Array.from(wrapper.children).forEach((child) => {
					if (child !== button) {
						child.remove();
					}
				});
				if (!button) {
					button = this.createDetachedEditorChromeElement(wrapper, 'button', 'settings');
					wrapper.appendChild(button);
				}

				button.type = 'button';
				button.textContent = this.docxT('docx:chrome.settings');
				markEditorChromeMenuButton(button);
			markEditorChromeNoToolbarTooltip(button);
				button.addClasses(['native-powerpoint-doc-editor-search-menu-button', 'native-powerpoint-doc-editor-settings-menu-button']);
				button.removeAttribute('aria-haspopup');
				button.removeAttribute('data-state');
				button.removeAttribute('id');
				button.setAttribute('aria-label', this.docxT('docx:chrome.settings'));
				button.setAttribute('role', 'menuitem');
				button.addEventListener('mousedown', (evt) => {
					evt.preventDefault();
				});
				button.addEventListener('click', (evt) => {
					evt.preventDefault();
					evt.stopImmediatePropagation();
					evt.stopPropagation();
					this.closeEditorOptionSearchMenu();
					this.closeEditorEditMenu();
					this.openPluginSettings();
				});

			const searchWrapper = menubar.querySelector('[data-native-powerpoint-doc-editor-search-menu-item]');
			if (searchWrapper) {
				searchWrapper.after(wrapper);
			} else {
				menubar.appendChild(wrapper);
			}
	}

	private openPluginSettings(): void {
		const setting = (
			this.app as unknown as {
				setting?: { open?: () => void; openTabById?: (id: string) => void };
			}
		).setting;
		if (!setting?.open || !setting.openTabById) {
			this.showNotice('powerpoint:notice.settingsUnavailable');
			return;
		}
		setting.open();
		setting.openTabById('native-powerpoint-doc-editor');
	}

		private addEditorFileExportAsMenuItem() {
		if (!this.hostEl) {
			return;
		}

		const retitleMenuButton = (
			button: HTMLButtonElement,
			label: string,
			sourceLabels: string[],
			options: { showChevron?: boolean } = {},
		) => {
			const labelElement = Array.from(button.children).find((child): child is HTMLElement => (
				isHTMLElement(child)
				&& sourceLabels.some(sourceLabel => textStartsWithMenuLabel(normalizeMenuText(child.textContent ?? ''), sourceLabel))
			));

			if (labelElement) {
				labelElement.textContent = label;
				Array.from(button.children).forEach((child) => {
					if (
						isHTMLElement(child)
						&& child !== labelElement
						&& /^(?:ctrl|cmd|⌘)/.test(normalizeMenuText(child.textContent ?? ''))
					) {
						child.remove();
					}
				});
			} else {
				button.textContent = label;
			}

			button.querySelectorAll(
				'[data-native-powerpoint-doc-editor-export-chevron], [data-native-power-point-doc-editor-export-chevron]',
			).forEach(chevron => chevron.remove());
				if (options.showChevron) {
					const chevron = this.createDetachedEditorChromeElement(button, 'span', 'export-as');
					setDocxEditorDataAttr(chevron, 'export-chevron');
					chevron.textContent = '›';
					chevron.addClass('native-powerpoint-doc-editor-export-chevron');
					button.appendChild(chevron);
				}
			};

		const labels = this.resolveEditorMenuLabels();
		this.hostEl.querySelectorAll<HTMLElement>(DOCX_EDITOR_MENU_ROOT_SELECTOR).forEach((menuItem) => {
				const menuButton = menuItem.querySelector(DOCX_EDITOR_MENU_BUTTON_SELECTOR);
				const menuLabel = normalizeMenuText(menuButton?.textContent ?? '');
				if (menuLabel !== labels.file) {
					return;
				}

				const dropdown = Array.from(menuItem.children).find((child): child is HTMLElement => (
					isHTMLElement(child)
					&& child !== menuButton
					&& child.matches(DOCX_EDITOR_MENU_DROPDOWN_SELECTOR)
				));
				if (!dropdown) {
					return;
				}

				const itemWrappers = Array.from(dropdown.children).filter((child): child is HTMLElement => isHTMLElement(child));
				const saveWrapper = itemWrappers.find((itemWrapper) => {
					const button = itemWrapper.querySelector(DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR);
					const text = normalizeMenuText(button?.textContent ?? '');
					return labels.save.some((label) => textStartsWithMenuLabel(text, label));
				});
				const sourceWrapper = saveWrapper ?? itemWrappers.find((itemWrapper) => itemWrapper.querySelector(DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR));

				let duplicateWrapper = takeUniqueMenuItem(
					dropdown,
					'duplicate-menu-item',
					'native-powerpoint-doc-editor-duplicate-menu-item',
				);
				if (!duplicateWrapper) {
					duplicateWrapper = sourceWrapper
						? sourceWrapper.cloneNode(true) as HTMLElement
						: this.createDetachedEditorChromeElement(dropdown, 'div', 'duplicate');
					setDocxEditorDataAttr(duplicateWrapper, 'duplicate-menu-item');
					duplicateWrapper.addClasses(['native-powerpoint-doc-editor-file-menu-item', 'native-powerpoint-doc-editor-duplicate-menu-item']);

					const duplicateButton = duplicateWrapper.querySelector('button') ?? duplicateWrapper.createEl('button');
					retitleMenuButton(duplicateButton, 'Duplicate current DOCX', labels.save);
					duplicateButton.removeAttribute('disabled');
					duplicateButton.removeAttribute('aria-disabled');
					configureMenuItemButton(duplicateButton, {
						className: 'native-powerpoint-doc-editor-file-menu-button',
						preventMouseDown: true,
						preventDefaultOnClick: true,
						stopClickPropagation: true,
						onClick: (evt) => {
							evt.stopImmediatePropagation();
							activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
							void this.duplicateCurrentDocument();
						},
					});

					if (saveWrapper) {
						saveWrapper.after(duplicateWrapper);
					} else {
						dropdown.prepend(duplicateWrapper);
					}
				}

				let exportWrapper = takeUniqueMenuItem(
					dropdown,
					'export-as-menu-item',
					'native-powerpoint-doc-editor-export-menu-item',
				);
				if (!exportWrapper) {
					exportWrapper = sourceWrapper
						? sourceWrapper.cloneNode(true) as HTMLElement
						: this.createDetachedEditorChromeElement(dropdown, 'div', 'export-as');
					setDocxEditorDataAttr(exportWrapper, 'export-as-menu-item');

					let exportButton = exportWrapper.querySelector('button');
					if (!exportButton) {
						exportButton = this.createDetachedEditorChromeElement(exportWrapper, 'button', 'export-as');
						exportWrapper.appendChild(exportButton);
					}
					exportWrapper.addClasses(['native-powerpoint-doc-editor-file-menu-item', 'native-powerpoint-doc-editor-export-menu-item']);
					retitleMenuButton(exportButton, 'Export as...', labels.save, { showChevron: true });
					exportButton.removeAttribute('disabled');
					exportButton.removeAttribute('aria-disabled');
					configureMenuItemButton(exportButton, {
						className: 'native-powerpoint-doc-editor-file-menu-button',
						preventMouseDown: true,
						preventDefaultOnClick: true,
						stopClickPropagation: true,
						onClick: (evt) => {
							evt.stopImmediatePropagation();
							activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
							void this.exportCurrentDocumentAs();
						},
					});

					const exportSubmenu = exportWrapper.createDiv({ cls: 'native-powerpoint-doc-editor-export-submenu' });
					for (const format of DOCX_EXPORT_FORMATS) {
						const optionWrapper = sourceWrapper
							? sourceWrapper.cloneNode(true) as HTMLElement
							: this.createDetachedEditorChromeElement(exportSubmenu, 'div', 'export-format');
						optionWrapper.removeAttribute('data-native-powerpoint-doc-editor-export-as-menu-item');
						optionWrapper.removeAttribute('data-native-power-point-doc-editor-export-as-menu-item');
						const optionButton = optionWrapper.querySelector('button') ?? optionWrapper.createEl('button');
						retitleMenuButton(optionButton, format.label, labels.save);
						optionButton.removeAttribute('disabled');
						optionButton.removeAttribute('aria-disabled');
						configureMenuItemButton(optionButton, {
							className: 'native-powerpoint-doc-editor-file-menu-button',
							preventMouseDown: true,
							preventDefaultOnClick: true,
							stopClickPropagation: true,
							onClick: (evt) => {
								evt.stopImmediatePropagation();
								activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
								void this.exportCurrentDocumentAs(format.id);
							},
						});
						exportSubmenu.appendChild(optionWrapper);
					}

					if (duplicateWrapper.parentElement === dropdown) {
						duplicateWrapper.after(exportWrapper);
					} else if (saveWrapper) {
						saveWrapper.after(exportWrapper);
					} else {
						dropdown.prepend(exportWrapper);
					}
				}

				let hiddenTextWrapper = takeUniqueMenuItem(
					dropdown,
					'find-hidden-text-menu-item',
					'native-powerpoint-doc-editor-find-hidden-text-menu-item',
				);
				if (!hiddenTextWrapper) {
					hiddenTextWrapper = sourceWrapper
						? sourceWrapper.cloneNode(true) as HTMLElement
						: this.createDetachedEditorChromeElement(dropdown, 'div', 'find-hidden-text');
					setDocxEditorDataAttr(hiddenTextWrapper, 'find-hidden-text-menu-item');
					hiddenTextWrapper.addClasses(['native-powerpoint-doc-editor-file-menu-item', 'native-powerpoint-doc-editor-find-hidden-text-menu-item']);
					const hiddenTextButton = hiddenTextWrapper.querySelector('button') ?? hiddenTextWrapper.createEl('button');
					retitleMenuButton(hiddenTextButton, 'Find hidden text...', labels.save);
					hiddenTextButton.removeAttribute('disabled');
					hiddenTextButton.removeAttribute('aria-disabled');
					configureMenuItemButton(hiddenTextButton, {
						className: 'native-powerpoint-doc-editor-file-menu-button',
						preventMouseDown: true,
						preventDefaultOnClick: true,
						stopClickPropagation: true,
						onClick: (evt) => {
							evt.stopImmediatePropagation();
							activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
							void this.findHiddenText();
						},
					});

					if (exportWrapper.parentElement === dropdown) {
						exportWrapper.after(hiddenTextWrapper);
					} else if (saveWrapper) {
						saveWrapper.after(hiddenTextWrapper);
					} else {
						dropdown.prepend(hiddenTextWrapper);
					}
				}
			});
	}

	private addEditorInsertMenuItems() {
		if (!this.hostEl) {
			return;
		}

		const labels = this.resolveEditorMenuLabels();
		this.hostEl.querySelectorAll<HTMLElement>(DOCX_EDITOR_MENU_ROOT_SELECTOR).forEach((menuItem) => {
					const menuButton = menuItem.querySelector(DOCX_EDITOR_MENU_BUTTON_SELECTOR);
					const menuLabel = normalizeMenuText(menuButton?.textContent ?? '');
					if (menuLabel !== labels.insert) {
						return;
					}

					const dropdown = Array.from(menuItem.children).find((child): child is HTMLElement => (
						isHTMLElement(child)
						&& child !== menuButton
						&& child.matches(DOCX_EDITOR_MENU_DROPDOWN_SELECTOR)
					));
					if (!dropdown) {
						return;
					}

					takeUniqueMenuItem(
						dropdown,
						'insert-image-menu-item',
						'native-powerpoint-doc-editor-insert-image-menu-item',
					);
					if (dropdown.querySelector(getDocxEditorDataAttrSelector('insert-image-menu-item'))) {
						return;
					}

					const itemWrappers = Array.from(dropdown.children).filter((child): child is HTMLElement => isHTMLElement(child));
					const sourceWrapper = itemWrappers.find((itemWrapper) => itemWrapper.querySelector(DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR));
					const insertImageWrapper = sourceWrapper
						? sourceWrapper.cloneNode(true) as HTMLElement
						: this.createDetachedEditorChromeElement(dropdown, 'div', 'insert-image');
					setDocxEditorDataAttr(insertImageWrapper, 'insert-image-menu-item');
					insertImageWrapper.addClasses(['native-powerpoint-doc-editor-file-menu-item', 'native-powerpoint-doc-editor-insert-image-menu-item']);

					const insertImageButton = insertImageWrapper.querySelector('button') ?? insertImageWrapper.createEl('button');
					insertImageButton.textContent = this.docxT('docx:chrome.insertImage');
					insertImageButton.removeAttribute('disabled');
					insertImageButton.removeAttribute('aria-disabled');
					configureMenuItemButton(insertImageButton, {
						className: 'native-powerpoint-doc-editor-file-menu-button',
						preventMouseDown: true,
						preventDefaultOnClick: true,
						stopClickPropagation: true,
						onClick: (evt) => {
							evt.stopImmediatePropagation();
							activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
							this.openImagePicker();
						},
					});

					const imageLikeWrapper = itemWrappers.find((itemWrapper) => {
						const button = itemWrapper.querySelector(DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR);
						const text = normalizeMenuText(button?.textContent ?? '');
						return text.includes('image') || text.includes('picture');
					});
					if (imageLikeWrapper) {
						imageLikeWrapper.replaceWith(insertImageWrapper);
					} else {
						dropdown.appendChild(insertImageWrapper);
				}
			});
	}

	private normalizeNativeEditorMenuActionItems() {
		if (!this.hostEl) {
			return;
		}

		const labels = this.resolveEditorMenuLabels();
				const menuSpecs = [
					{ menuLabel: labels.file, optionLabels: labels.pageSetup },
					{ menuLabel: labels.format, optionLabels: labels.rightToLeft },
					{ menuLabel: labels.insert, optionLabels: labels.pageBreak },
					{ menuLabel: labels.insert, optionLabels: labels.tableOfContents },
				];

				this.hostEl.querySelectorAll<HTMLElement>(DOCX_EDITOR_MENU_ROOT_SELECTOR).forEach((menuItem) => {
					const menuButton = menuItem.querySelector(DOCX_EDITOR_MENU_BUTTON_SELECTOR);
					const menuLabel = normalizeMenuText(menuButton?.textContent ?? '');
					const specs = menuSpecs.filter(spec => spec.menuLabel === menuLabel);
					if (specs.length === 0) {
						return;
					}

					const dropdown = Array.from(menuItem.children).find((child): child is HTMLElement => (
						isHTMLElement(child)
						&& child !== menuButton
						&& child.matches(DOCX_EDITOR_MENU_DROPDOWN_SELECTOR)
					));
					if (!dropdown) {
						return;
					}

					for (const button of Array.from(dropdown.querySelectorAll<HTMLButtonElement>(DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR))) {
						const buttonLabel = normalizeMenuText(cleanEditorOptionLabel(button.textContent ?? ''));
						const matchesTarget = specs.some(spec => spec.optionLabels.some(label => (
							textStartsWithMenuLabel(buttonLabel, label)
							|| textStartsWithMenuLabel(normalizeMenuText(label), buttonLabel)
						)));
						if (!matchesTarget) {
							continue;
						}

						const wrapper = Array.from(dropdown.children).find((child): child is HTMLElement => (
							isHTMLElement(child) && child.contains(button)
						));
						wrapper?.addClasses(['native-powerpoint-doc-editor-file-menu-item', 'native-powerpoint-doc-editor-native-menu-action-item']);
						button.addClasses(['native-powerpoint-doc-editor-file-menu-button', 'native-powerpoint-doc-editor-native-menu-action-button']);
				}
			});
	}

	private registerEditorSaveInterceptor() {
		this.registerDomEvent(activeDocument, 'click', (evt) => {
				if (
					!this.hostEl
					|| this.app.workspace.getActiveViewOfType(DocxView) !== this
					|| (isElement(evt.target) && !!evt.target.closest('.modal'))
					|| !shouldHandleEditorSaveClick(evt.target, this.resolveEditorMenuLabels().save)
				) {
				return;
			}

			evt.preventDefault();
			evt.stopImmediatePropagation();
			void this.saveCurrentDocument();
		}, true);
	}

	private registerEditorListAwareCopyInterceptor() {
		this.registerDomEvent(activeDocument, 'copy', (evt) => {
			if (!this.hostEl) {
				return;
			}

				const targetInsideHost = isNode(evt.target) && this.hostEl.contains(evt.target);
				const activeInsideHost = isNode(activeDocument.activeElement) && this.hostEl.contains(activeDocument.activeElement);
			const selection = activeDocument.getSelection();
			const selectionInsideHost = Boolean(
				selection
				&& !selection.isCollapsed
				&& (
					(selection.anchorNode && this.hostEl.contains(selection.anchorNode))
					|| (selection.focusNode && this.hostEl.contains(selection.focusNode))
				),
			);
			if (!targetInsideHost && !activeInsideHost && !selectionInsideHost) {
				return;
			}

			const handle = this.getReactHandle();
			if (handle) {
				window.setTimeout(() => {
					void handle.rewriteClipboardTextWithListMarkers();
				}, 0);
			}
		}, true);
	}

	private registerSaveShortcut() {
		this.registerDomEvent(activeDocument, 'keydown', (evt) => {
			if (
					!this.hostEl
					|| evt.key.toLowerCase() !== 's'
					|| (!evt.metaKey && !evt.ctrlKey)
					|| !isNode(activeDocument.activeElement)
					|| !this.hostEl.contains(activeDocument.activeElement)
				) {
				return;
			}

			evt.preventDefault();
			evt.stopImmediatePropagation();
			void this.saveCurrentDocument();
		}, true);
	}

	private registerAgentUndoShortcut() {
		const handleAgentUndo = (evt: KeyboardEvent) => {
			if (!this.hostEl || !this.file || (!evt.metaKey && !evt.ctrlKey) || evt.key.toLowerCase() !== 'z') {
				return;
			}
			if (!isNode(activeDocument.activeElement) || !this.hostEl.contains(activeDocument.activeElement)) {
				return;
			}

			if (evt.shiftKey) {
				if (!this.canRedoAgentEdit()) {
					return;
				}
				evt.preventDefault();
				evt.stopImmediatePropagation();
				void this.redoAgentEdit();
				return;
			}

			if (!this.canUndoAgentEdit()) {
				return;
			}

			evt.preventDefault();
			evt.stopImmediatePropagation();
			void this.undoAgentEdit();
		};

		this.registerDomEvent(window, 'keydown', handleAgentUndo, true);
		this.registerDomEvent(activeDocument, 'keydown', handleAgentUndo, true);
	}

	private registerFindShortcut() {
		const handleFindShortcut = (evt: KeyboardEvent) => {
			if (!this.hostEl || !isPrimaryFindShortcut(evt) || !this.isActiveDocxView()) {
				return;
			}

				const target = isElement(evt.target) ? evt.target : null;
			if (target?.closest('.modal') && !target.closest('.native-powerpoint-doc-editor-find-dialog')) {
				return;
			}

			evt.preventDefault();
			evt.stopImmediatePropagation();
			this.openFindDialog();
		};

		this.registerDomEvent(window, 'keydown', handleFindShortcut, true);
		this.registerDomEvent(activeDocument, 'keydown', handleFindShortcut, true);
	}

	private isActiveDocxView() {
		if (this.app.workspace.getActiveViewOfType(DocxView) === this) {
			return true;
		}

			if (this.contentEl.closest('.workspace-leaf.mod-active')) {
				return true;
			}

			const activeElement = activeDocument.activeElement;
			return Boolean(isNode(activeElement) && this.hostEl?.contains(activeElement));
		}

		private registerEditorDropdownScrollGuard() {
			const keepEditorListboxOpen = (evt: Event) => {
				if (!this.hostEl || this.app.workspace.getActiveViewOfType(DocxView) !== this || !isElement(evt.target)) {
					return;
				}

			const listbox = evt.target.closest('[role="listbox"]');
			if (listbox && this.hostEl.contains(listbox)) {
				evt.stopImmediatePropagation();
				evt.stopPropagation();
			}
		};

		this.registerDomEvent(window, 'scroll', keepEditorListboxOpen, true);
	}

	private getReactHandle(): DocxReactViewHandle | null {
		return this.reactMount?.getHandle() ?? null;
	}

	private getReactProps(): DocxReactViewProps {
		if (!this.editorAdapter) {
			throw new Error('DOCX editor adapter is not initialized');
		}
		const origin = this.file ? this.createSaveOrigin(this.file) : null;

		return {
			file: this.file,
			buffer: this.buffer,
			documentKey: origin ? `${origin.filePath}:${origin.documentSession}` : 'native-powerpoint-doc-editor-empty',
			editorAdapter: this.editorAdapter,
			error: this.error,
			isLoading: this.isLoading,
			authorName: this.getAuthorName(),
			resolvedEditorTheme: this.getResolvedEditorTheme(),
			i18n: this.getEditorLocale(),
			pluginI18n: this.getPluginI18n(),
			showNotice: (key, values) => this.showNotice(key, values),
			showRuler: this.getShowRuler(),
			autosave: this.getAutosave(),
			defaultZoom: this.getDefaultZoom(),
			reserveReviewSidebar: this.reserveReviewSidebar,
			hostDocument: this.hostEl?.ownerDocument,
			onDirtyChange: (isDirty) => {
				if (origin && this.isCurrentSaveOrigin(origin)) {
					this.isDirty = isDirty;
				}
			},
			onSave: (buffer) => this.saveFile(buffer, origin ?? this.createSaveOrigin()),
			onDocumentNameChange: (name, expectedPath) => this.renameFile(name, expectedPath),
			onWordCountChange: this.onWordCountChange,
			onLoadPhase: (phase, data) => this.handleEditorLoadPhase(phase, data, origin ?? undefined),
		};
	}

	private async ensureReactMount() {
		if (this.reactMount || this.reactMountLoading || !this.hostEl) {
			debugLog('editor', 'Skipping React mount request', {
				hasMount: Boolean(this.reactMount),
				isLoading: this.reactMountLoading,
				hasHost: Boolean(this.hostEl),
			});
			this.markOpenLoadPhase('react-mount-skipped', {
				hasMount: Boolean(this.reactMount),
				isLoading: this.reactMountLoading,
				hasHost: Boolean(this.hostEl),
			});
			return;
		}

		this.reactMountLoading = true;
		try {
			infoLog('editor', 'Loading DOCX editor UI');
			this.markOpenLoadPhase('react-mount-start');
			const chunkStartedAt = monotonicNow();
			const { createDocxReactMount } = await loadDocxEditorChunk();
			this.markOpenLoadPhase('react-chunk-loaded', {
				durationMs: Math.round((monotonicNow() - chunkStartedAt) * 10) / 10,
			});
			if (!this.hostEl) {
				debugLog('editor', 'Aborted DOCX editor mount because host was removed');
				return;
			}

			this.hostEl.empty();
			const mountStartedAt = monotonicNow();
			const mountOrigin = this.file ? this.createSaveOrigin(this.file) : null;
			this.reactMount = createDocxReactMount(this.hostEl, (renderError) => {
				if (mountOrigin && !this.isCurrentSaveOrigin(mountOrigin)) {
					return;
				}
				this.finishOpenLoadTrace('react-render-failed', {
					file: mountOrigin?.filePath ?? this.file?.path,
					message: renderError.message,
					name: renderError.name,
				});
			});
			this.reactMount.render(this.getReactProps());
			this.markOpenLoadPhase('react-mount-rendered', {
				durationMs: Math.round((monotonicNow() - mountStartedAt) * 10) / 10,
			});
			this.registerEditorChromeCustomization();
			infoLog('editor', 'Mounted DOCX editor UI', { file: this.file?.path });
		} catch (loadError) {
			const message = loadError instanceof Error ? loadError.message : 'Unknown load error';
			this.error = `Could not load DOCX editor: ${message}`;
			this.finishOpenLoadTrace('react-mount-failed', {
				file: this.file?.path,
				message,
			});
			errorLog('editor', this.error, loadError);
			showI18nNotice(this.getPluginI18n(), this.error);
			if (this.hostEl) {
				this.hostEl.empty();
				this.hostEl.createDiv({ cls: 'native-powerpoint-doc-editor-editor-load-error', text: this.error });
			}
		} finally {
			this.reactMountLoading = false;
		}
	}

	private render() {
		if (!this.hostEl) {
			return;
		}

		if (!this.isLeafActive() && !this.reactMount) {
			if (!this.app.workspace.layoutReady) {
				this.markOpenLoadPhase('render-startup-loading-shell');
				this.hostEl.empty();
				this.hostEl.createDiv({
					cls: 'native-powerpoint-doc-editor-editor-loading',
					text: 'Loading DOCX editor...',
				});
				return;
			}
			this.markOpenLoadPhase('render-inactive-placeholder');
			this.renderInactivePlaceholder();
			return;
		}

		if (!this.reactMount) {
			this.markOpenLoadPhase('render-loading-shell');
			this.hostEl.empty();
			this.hostEl.createDiv({ cls: 'native-powerpoint-doc-editor-editor-loading', text: 'Loading DOCX editor...' });
			this.scheduleEditorMount();
			return;
		}

		this.reactMount.render(this.getReactProps());
		this.registerEditorChromeCustomization();
	}
}
