import type { TFile } from 'obsidian';
import { forwardRef, useCallback, useEffect, useImperativeHandle, useMemo, useRef, useState, type ChangeEvent, type ComponentProps } from 'react';
import { AllSelection, Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import type { Mark, Node as ProseMirrorNode, Slice } from 'prosemirror-model';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import proseMirrorViewStyles from 'prosemirror-view/style/prosemirror.css';
import type { I18nService } from './i18n/I18nService';
import {
	DocxEditor,
	clearParagraphMeasureCache,
	insertTable,
	setFontSize,
	setFontFamily,
	setLineSpacing,
	loadFontFromBuffer,
	isSuggestionModeActive,
} from './docx/runtime/bridge.mjs';
import type { DocxEditorRef, EditorMode, FontOption, RenderedDomContext, Translations } from './docx/runtime/contract';
import { docxEditorRuntimeStyles } from './docx/runtime/styles';
import { parsePrimaryFontFamily } from './powerpoint/textUtils';
import { isClipboardEvent, isElement, isHTMLElement, isHTMLButtonElement, isInputEvent, isNode, isPointerEvent } from './domGuards';
import { summarizeDocxComment, summarizeDocxComments } from './docxCommentLogging';
import { debugLog, errorLog, warnLog } from './logger';
import { Platform } from './obsidianRuntime';
import { configureToolbarIconButton, createMenuItem, createMenuSection, hardenInjectedMenuOption } from './menuControls';
import { DOCX_SAVE_STATUS_TO_STATE, getSaveStatusFlags, type DocxSaveStatus } from './save/saveStatus';
import { formatFindResultStatus, wrapMatchIndex, type FindReplaceMode } from './find/findReplaceShell';
import { countDocumentWords, type DocumentWordCount } from './documentWordCount';
import { attachDocxImeTransformNeutralizer } from './docxImeTransformNeutralizer';
import { exportRenderedPagesToPdf } from './renderedPdfExport';
import { didListLayoutChange, didParagraphLayoutChange } from './docxParagraphLayoutRelayout';
import {
	countDocTextblocks,
	insertRichClipboardSlice,
	insertPlainTextAsParagraphs,
	insertPlainTypedText as dispatchPlainTypedText,
	summarizeRichClipboardSlice,
	summarizeTransactionMeta,
	summarizeTransactionSteps,
} from './docxPlainTextInsert';
import { preserveDocxTableCellFontSizes } from './docxTableCellFontSizePreserver';
import type { DocxEditorAdapterController, DocxFindMatch } from './docx/adapter/DocxEditorAdapter';
import { DocxSession, type DocxSaveSource } from './docx/session/DocxSession';
import {
	calculateClampedFloatingLayerPosition,
	getFormattingDropdownScrollTransition,
	isFullscreenDialogLayer,
} from './docxFloatingLayerLayout';
import type { EditorThemeResolution } from './settings';
import {
	attachDocxToolbarTooltipManager,
	stripFormattingDropdownButtonTitles,
	suppressVendorToolbarTooltips,
} from './docxToolbarTooltip';
import {
	DOCX_EDITOR_FORMATTING_BAR_SELECTOR,
	DOCX_EDITOR_MENUBAR_SELECTOR,
	DOCX_EDITOR_PAGES_SELECTOR,
	DOCX_EDITOR_TITLE_BAR_SELECTOR,
	DOCX_FONT_SIZE_DECREASE_SELECTOR,
	DOCX_FONT_SIZE_DISPLAY_SELECTOR,
	DOCX_FONT_SIZE_INCREASE_SELECTOR,
	DOCX_FONT_SIZE_INPUT_SELECTOR,
	DOCX_RENDERED_LIST_MARKER_SELECTOR,
	DOCX_RENDERED_PAGE_SELECTOR,
	DOCX_RENDERED_PARAGRAPH_SELECTOR,
} from './docxEditorChromeMarkers';

let stylesInjected = false;
let editorInstanceCounter = 0;
const docxEditorStyles = [
	proseMirrorViewStyles,
	docxEditorRuntimeStyles,
].join('\n');

interface DocxSectionProperties {
	pageHeight?: number;
	marginTop?: number;
	marginBottom?: number;
}

interface DocxDocumentWithSectionProperties {
	package?: {
		[DOCX_PACKAGE_DOCUMENT_KEY]?: {
			finalSectionProperties?: DocxSectionProperties;
			sections?: Array<{
				properties?: DocxSectionProperties;
			}>;
		};
	};
}

interface DocxPaginationSourceDiagnostics {
	paragraphs: number;
	tables: number;
	tabRuns: number;
	tabHeavyParagraphs: number;
	maxTabsInParagraph: number;
	longSpaceRuns: number;
	explicitPageBreaks: number;
}

const DEFAULT_PAGE_HEIGHT_TWIPS = 15840;
const DEFAULT_MARGIN_TWIPS = 1440;
const DOCX_PACKAGE_DOCUMENT_KEY = 'document';
const MIN_TOUCH_ZOOM = 0.25;
const MAX_TOUCH_ZOOM = 4;
const TOUCH_ZOOM_SENSITIVITY = 0.55;
const TOUCH_ZOOM_MIN_DELTA = 0.006;
const IMPORT_FONT_MENU_LABEL = 'Import font...';
const CUSTOM_TABLE_MENU_LABEL = 'Custom...';
const FONT_FAMILY_TRIGGER_ATTRIBUTE = 'data-native-powerpoint-doc-editor-font-family-trigger';
const FONT_FAMILY_TRIGGER_DATA_KEY = 'nativePowerPointDocEditorFontFamilyTrigger';
const FONT_FILE_ACCEPT = '.ttf,.otf,.woff,.woff2,font/ttf,font/otf,font/woff,font/woff2,application/font-woff,application/font-woff2';
const FONT_FILE_EXTENSION_PATTERN = /\.(?:ttf|otf|woff2?)$/i;
const MIN_CUSTOM_TABLE_SIZE = 1;
const MAX_CUSTOM_TABLE_SIZE = 50;
const DEFAULT_FONT_SIZE_POINTS = 11;
const MIN_FONT_SIZE_POINTS = 1;
const MAX_FONT_SIZE_POINTS = 400;
const FONT_SIZE_HOLD_INITIAL_DELAY_MS = 420;
const FONT_SIZE_HOLD_INITIAL_INTERVAL_MS = 180;
const FONT_SIZE_HOLD_INTERVAL_DECAY = 0.82;
const FONT_SIZE_HOLD_MIN_INTERVAL_MS = 36;
const LINE_SPACING_TWIPS = new Set([240, 276, 360, 480]);
const SELECTED_LIST_MARKER_CLASS = 'native-powerpoint-doc-editor-list-marker-selected';
const LIST_PARAGRAPH_SELECTOR = `${DOCX_RENDERED_PARAGRAPH_SELECTOR}[data-pm-start]`;
const LIST_MARKER_SELECTOR = DOCX_RENDERED_LIST_MARKER_SELECTOR;
const DEFAULT_EDITOR_FONT_FAMILIES: FontOption[] = [
	{ name: 'Arial', fontFamily: 'Arial, Helvetica, sans-serif', category: 'sans-serif' },
	{ name: 'Calibri', fontFamily: '"Calibri", Arial, sans-serif', category: 'sans-serif' },
	{ name: 'Helvetica', fontFamily: 'Helvetica, Arial, sans-serif', category: 'sans-serif' },
	{ name: 'Verdana', fontFamily: 'Verdana, Geneva, sans-serif', category: 'sans-serif' },
	{ name: 'Open Sans', fontFamily: '"Open Sans", sans-serif', category: 'sans-serif' },
	{ name: 'Roboto', fontFamily: 'Roboto, sans-serif', category: 'sans-serif' },
	{ name: 'Times New Roman', fontFamily: '"Times New Roman", Times, serif', category: 'serif' },
	{ name: 'Georgia', fontFamily: 'Georgia, serif', category: 'serif' },
	{ name: 'Cambria', fontFamily: 'Cambria, Georgia, serif', category: 'serif' },
	{ name: 'Garamond', fontFamily: 'Garamond, serif', category: 'serif' },
	{ name: 'Courier New', fontFamily: '"Courier New", Courier, monospace', category: 'monospace' },
	{ name: 'Consolas', fontFamily: 'Consolas, monospace', category: 'monospace' },
];

function isPrimaryShortcut(evt: KeyboardEvent, key: string): boolean {
	const normalizedKey = evt.key.toLowerCase();
	const isMacShortcut = evt.metaKey && !evt.ctrlKey;
	const isNonMacShortcut = evt.ctrlKey && !evt.metaKey && !Platform.isMacOS;
	const hasPrimaryModifier = isMacShortcut || isNonMacShortcut;
	return normalizedKey === key && hasPrimaryModifier && !evt.altKey && !evt.shiftKey;
}

type SaveStatus = DocxSaveStatus;
type FontSizeStepDirection = -1 | 1;

interface FontSizeHoldState {
	control: HTMLElement | null;
	currentSize: number;
	direction: FontSizeStepDirection;
	repeatCount: number;
	repeatTimer: number | null;
	startTimer: number | null;
}

interface ExportDocumentBufferOptions {
	preserveAutosave?: boolean;
}

type FindMatch = DocxFindMatch;

interface DocxSaveContext {
	file: TFile;
	persist(buffer: ArrayBuffer): Promise<void>;
}

interface RefreshFindOptions {
	select?: boolean;
}

export interface PasteClipboardOptions {
	preserveFormatting: boolean;
}

interface FindHighlightState {
	matches: FindMatch[];
	currentIndex: number;
}

interface ObsidianDesktopWindow extends Window {
	require?: (moduleId: string) => unknown;
}

interface PinchZoomState {
	source: 'touch' | 'gesture' | 'pointer';
	startDistance: number;
	lastDistance: number;
	startZoom: number;
	lastZoom: number;
}

interface PointerPoint {
	x: number;
	y: number;
}

type WebKitGestureEvent = Event & {
	clientX?: number;
	clientY?: number;
	scale?: number;
};

interface FindReplaceLabels {
	find: string;
	findAndReplace: string;
	findText: string;
	replaceWith: string;
	replace: string;
	replaceAll: string;
	matchCase: string;
	wholeWords: string;
	showReplace: string;
	close: string;
	previous: string;
	next: string;
	noMatches: string;
	resultCount: (current: number, total: number) => string;
}

function createFindReplaceLabels(i18n: I18nService): FindReplaceLabels {
	return {
		find: i18n.t('docx:find.title'),
		findAndReplace: i18n.t('docx:find.titleReplace'),
		findText: i18n.t('docx:find.placeholder'),
		replaceWith: i18n.t('docx:find.replacePlaceholder'),
		replace: i18n.t('docx:find.replaceButton'),
		replaceAll: i18n.t('docx:find.replaceAllButton'),
		matchCase: i18n.t('docx:find.matchCase'),
		wholeWords: i18n.t('docx:find.wholeWords'),
		showReplace: i18n.t('docx:find.toggleReplace'),
		close: i18n.t('common:actions.close'),
		previous: i18n.t('docx:find.findPrevious'),
		next: i18n.t('docx:find.findNext'),
		noMatches: i18n.t('docx:find.noResults'),
		resultCount: (current, total) => i18n.t('docx:find.resultCount', { current, total }),
	};
}

const findHighlightPluginKey = new PluginKey<FindHighlightState>('native-powerpoint-doc-editor-find-highlight');
const paragraphLayoutRelayoutPluginKey = new PluginKey('native-powerpoint-doc-editor-paragraph-layout-relayout');
const preserveTypedSpacePluginKey = new PluginKey('native-powerpoint-doc-editor-preserve-typed-space');
const contentShrinkDiagnosticsPluginKey = new PluginKey('native-powerpoint-doc-editor-content-shrink-diagnostics');
const LIST_PARAGRAPH_DEFAULT_LEFT_INDENT = 720;
const LIST_PARAGRAPH_DEFAULT_HANGING_INDENT = 360;
const LIST_LAYOUT_NORMALIZED_META = 'native-powerpoint-doc-editor-list-layout-normalized';

function isFindHighlightState(value: unknown): value is FindHighlightState {
	if (!value || typeof value !== 'object') {
		return false;
	}

	const state = value as Partial<FindHighlightState>;
	return Array.isArray(state.matches) && typeof state.currentIndex === 'number';
}

function getDocxPaginationSourceDiagnostics(doc: ProseMirrorNode | null | undefined): DocxPaginationSourceDiagnostics {
	const diagnostics: DocxPaginationSourceDiagnostics = {
		paragraphs: 0,
		tables: 0,
		tabRuns: 0,
		tabHeavyParagraphs: 0,
		maxTabsInParagraph: 0,
		longSpaceRuns: 0,
		explicitPageBreaks: 0,
	};
	if (!doc) {
		return diagnostics;
	}

	doc.descendants((node) => {
		if (node.type.name === 'table') {
			diagnostics.tables += 1;
		}
		if (node.type.name === 'pageBreak'
			|| (node.type.name === 'hardBreak' && (node.attrs.breakType === 'page' || node.attrs.type === 'page'))) {
			diagnostics.explicitPageBreaks += 1;
		}
		if (node.type.name !== 'paragraph') {
			return true;
		}

		diagnostics.paragraphs += 1;
		let paragraphTabRuns = 0;
		node.descendants((child) => {
			if (child.type.name === 'tab') {
				paragraphTabRuns += 1;
			}
			if (child.isText && child.text) {
				diagnostics.longSpaceRuns += (child.text.match(/ {8,}/g) ?? []).length;
			}
			return true;
		});
		diagnostics.tabRuns += paragraphTabRuns;
		diagnostics.maxTabsInParagraph = Math.max(diagnostics.maxTabsInParagraph, paragraphTabRuns);
		if (paragraphTabRuns >= 2) {
			diagnostics.tabHeavyParagraphs += 1;
		}
		return true;
	});

	return diagnostics;
}

function getListParagraphIndentAttrs(attrs: Record<string, unknown>) {
	const numPr = attrs.numPr;
	if (!numPr || typeof numPr !== 'object') {
		return null;
	}

	const { numId, ilvl } = numPr as { numId?: unknown; ilvl?: unknown };
	if (typeof numId !== 'number' || numId === 0) {
		return null;
	}

	const level = typeof ilvl === 'number' && Number.isFinite(ilvl) ? Math.max(0, ilvl) : 0;
	const hasLeftIndent = typeof attrs.indentLeft === 'number' && Number.isFinite(attrs.indentLeft);
	const hasHangingIndent = attrs.hangingIndent === true
		&& typeof attrs.indentFirstLine === 'number'
		&& Number.isFinite(attrs.indentFirstLine)
		&& attrs.indentFirstLine < 0;

	if (hasLeftIndent && hasHangingIndent) {
		return null;
	}

	return {
		...attrs,
		indentLeft: hasLeftIndent ? attrs.indentLeft : (level + 1) * LIST_PARAGRAPH_DEFAULT_LEFT_INDENT,
		indentFirstLine: hasHangingIndent ? attrs.indentFirstLine : -LIST_PARAGRAPH_DEFAULT_HANGING_INDENT,
		hangingIndent: true,
	};
}

function createParagraphLayoutRelayoutPlugin(scheduleRelayout: () => void) {
	return new Plugin({
		key: paragraphLayoutRelayoutPluginKey,
		appendTransaction: (transactions, previousState, nextState) => {
			const shouldNormalize = transactions.some((transaction) => transaction.docChanged)
				&& transactions.every((transaction) => transaction.getMeta(LIST_LAYOUT_NORMALIZED_META) !== true)
				&& didListLayoutChange(previousState.doc, nextState.doc);

			if (!shouldNormalize) {
				return null;
			}

			let transaction = nextState.tr;
			let normalizedCount = 0;
			nextState.doc.descendants((node, position) => {
				if (node.type.name !== 'paragraph') {
					return true;
				}

				const normalizedAttrs = getListParagraphIndentAttrs(node.attrs);
				if (normalizedAttrs) {
					transaction = transaction.setNodeMarkup(position, undefined, normalizedAttrs, node.marks);
					normalizedCount += 1;
				}

				return false;
			});

			if (!transaction.docChanged) {
				return null;
			}

			transaction.setMeta(LIST_LAYOUT_NORMALIZED_META, true);
			debugLog('editor', 'Normalized DOCX list paragraph indentation', { paragraphs: normalizedCount });
			return transaction;
		},
		view: () => ({
			update: (view, previousState) => {
				if (previousState.doc === view.state.doc || previousState.doc.eq(view.state.doc)) {
					return;
				}

				if (didParagraphLayoutChange(previousState.doc, view.state.doc)) {
					debugLog('editor', 'DOCX paragraph layout changed; scheduling relayout');
					scheduleRelayout();
				}
			},
		}),
	});
}

type DocxTextInputRoute = 'beforeinput' | 'keydown' | 'text-input';

let lastDocxSpaceInputLogAt = 0;

function logDocxSpaceInputRoute(
	route: DocxTextInputRoute,
	details: {
		from: number;
		to: number;
		inputType?: string;
		key?: string;
		code?: string;
		isComposing: boolean;
		defaultPreventedBefore?: boolean;
		defaultPreventedAfter?: boolean;
	},
) {
	const now = Date.now();
	if (now - lastDocxSpaceInputLogAt < 250) {
		return;
	}
	lastDocxSpaceInputLogAt = now;
	debugLog('text-input', 'DOCX space input routed', {
		route,
		isSpace: true,
		selectionStart: details.from,
		selectionEnd: details.to,
		selectedCharacters: details.to - details.from,
		inputType: details.inputType ?? null,
		key: details.key ?? null,
		code: details.code ?? null,
		isComposing: details.isComposing,
		defaultPreventedBefore: details.defaultPreventedBefore ?? null,
		defaultPreventedAfter: details.defaultPreventedAfter ?? null,
	});
}

function insertPlainTypedText(view: EditorView, text: string, from = view.state.selection.from, to = view.state.selection.to) {
	if (/[\r\n]/.test(text)) {
		const range = insertPlainTextAsParagraphs(view, text, from, to);
		debugLog('clipboard', 'DOCX multi-line plain paste inserted as paragraphs', {
			textLength: text.length,
			lineCount: text.split(/\r\n|\n|\r/).length,
			replacedCrossBlock: range.collapsedCrossBlock,
			replaceFrom: range.from,
			replaceTo: range.to,
		});
		return true;
	}

	const range = dispatchPlainTypedText(view, text, from, to);
	if (range.collapsedCrossBlock) {
		debugLog('text-input', 'DOCX plain insert collapsed cross-block selection', {
			requestedFrom: from,
			requestedTo: to,
			insertAt: range.from,
			textLength: text.length,
			hasNewline: false,
		});
	}
	return true;
}

function isSupportedFontFile(file: File) {
	return FONT_FILE_EXTENSION_PATTERN.test(file.name) || file.type.startsWith('font/');
}

function getImportedFontBaseName(file: File) {
	return file.name
		.replace(FONT_FILE_EXTENSION_PATTERN, '')
		.replace(/[-_]+/g, ' ')
		.replace(/\s+/g, ' ')
		.trim() || 'Imported Font';
}

function getUniqueImportedFontName(file: File, fonts: FontOption[]) {
	const baseName = getImportedFontBaseName(file);
	const usedNames = new Set(fonts.map((font) => font.name.toLowerCase()));
	let fontName = baseName;
	let suffix = 2;

	while (usedNames.has(fontName.toLowerCase())) {
		fontName = `${baseName} ${suffix}`;
		suffix += 1;
	}

	return fontName;
}

function cssFontFamilyName(fontName: string) {
	return `"${fontName.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

function clampFontSizePoints(value: number) {
	if (!Number.isFinite(value)) {
		return DEFAULT_FONT_SIZE_POINTS;
	}

	return Math.min(MAX_FONT_SIZE_POINTS, Math.max(MIN_FONT_SIZE_POINTS, Math.round(value * 2) / 2));
}

function fontSizePointsToHalfPoints(value: number) {
	return Math.round(clampFontSizePoints(value) * 2);
}

function parseFontSizePoints(value: string | null | undefined) {
	const match = value?.replace(',', '.').match(/\d+(?:\.\d+)?/);
	if (!match) {
		return DEFAULT_FONT_SIZE_POINTS;
	}

	return clampFontSizePoints(Number(match[0]));
}

function getFontSizeControl(button: HTMLElement) {
	return button.parentElement;
}

function readFontSizeControlPoints(control: HTMLElement | null) {
	const input = control?.querySelector<HTMLInputElement>(DOCX_FONT_SIZE_INPUT_SELECTOR);
	if (input) {
		return parseFontSizePoints(input.value);
	}

	const display = control?.querySelector<HTMLElement>(DOCX_FONT_SIZE_DISPLAY_SELECTOR);
	return parseFontSizePoints(display?.textContent);
}

function updateFontSizeControlDisplay(control: HTMLElement | null, value: number) {
	const nextText = String(clampFontSizePoints(value));
	const input = control?.querySelector<HTMLInputElement>(DOCX_FONT_SIZE_INPUT_SELECTOR);
	if (input) {
		input.value = nextText;
	}

	const display = control?.querySelector<HTMLElement>(DOCX_FONT_SIZE_DISPLAY_SELECTOR);
	if (display) {
		display.textContent = nextText;
	}
}

function getFontSizeStepTarget(target: EventTarget | null) {
	if (!isElement(target)) {
		return null;
	}

	const button = target.closest<HTMLButtonElement>(`${DOCX_FONT_SIZE_DECREASE_SELECTOR}, ${DOCX_FONT_SIZE_INCREASE_SELECTOR}`);
	if (!button) {
		return null;
	}

	const direction: FontSizeStepDirection = button.matches(DOCX_FONT_SIZE_DECREASE_SELECTOR) ? -1 : 1;
	return { button, direction };
}

function parseLineSpacingTwips(value: string | null | undefined) {
	if (!value) {
		return null;
	}

	const twips = Number.parseInt(value, 10);
	return LINE_SPACING_TWIPS.has(twips) ? twips : null;
}

function getLineSpacingOptionTarget(target: EventTarget | null) {
	if (!isElement(target)) {
		return null;
	}

	const option = target.closest<HTMLElement>('[role="option"]');
	if (!option) {
		return null;
	}

	const twips = parseLineSpacingTwips(option.getAttribute('data-value') ?? option.getAttribute('value'));
	if (twips === null) {
		return null;
	}

	return { option, twips };
}

function applyLineSpacingToEditor(view: EditorView, twipsValue: number) {
	const { state } = view;
	const { selection } = state;

	if (!selection.empty) {
		return setLineSpacing(twipsValue)(state, view.dispatch);
	}

	const caret = selection.from;
	const selectAllState = state.apply(state.tr.setSelection(new AllSelection(state.doc)));
	let applied = false;

	setLineSpacing(twipsValue)(selectAllState, (spacingTr) => {
		applied = spacingTr.docChanged;
		if (!applied) {
			return;
		}

		const nextPos = Math.min(caret, Math.max(1, spacingTr.doc.content.size - 1));
		view.dispatch(
			spacingTr
				.setSelection(TextSelection.create(spacingTr.doc, nextPos))
				.scrollIntoView(),
		);
	});

	return applied;
}

const COMMENTS_SIDEBAR_TOGGLE_ATTR = 'data-native-powerpoint-doc-editor-comments-sidebar-toggle';

function getTopLevelCommentCount(editor: DocxEditorRef | null | undefined) {
	return (editor?.getComments() ?? []).filter((comment) => comment.parentId == null).length;
}

function findCommentsSidebarToggleButton(editorRoot: HTMLElement, toggleLabel: string) {
	const marked = editorRoot.querySelector<HTMLButtonElement>(`[${COMMENTS_SIDEBAR_TOGGLE_ATTR}]`);
	if (marked) {
		return marked;
	}

	const titleBar = editorRoot.querySelector(DOCX_EDITOR_TITLE_BAR_SELECTOR);
	if (!titleBar) {
		return null;
	}

	const escapedLabel = typeof CSS !== 'undefined' && 'escape' in CSS
		? CSS.escape(toggleLabel)
		: toggleLabel.replace(/"/g, '\\"');
	const button = titleBar.querySelector<HTMLButtonElement>(`button[aria-label="${escapedLabel}"]`);
	if (!button) {
		return null;
	}

	button.setAttribute(COMMENTS_SIDEBAR_TOGGLE_ATTR, 'true');
	return button;
}

function setCommentsSidebarToggleEnabled(button: HTMLButtonElement, enabled: boolean) {
	button.toggleAttribute('disabled', !enabled);
	button.setAttribute('aria-disabled', enabled ? 'false' : 'true');
	button.classList.toggle('native-powerpoint-doc-editor-comments-sidebar-toggle-disabled', !enabled);
}

function markLightMenuSurface(surface: HTMLElement, className: string) {
	surface.classList.add('native-powerpoint-doc-editor-light-menu-surface', className);

	const shell = surface.closest<HTMLElement>('[style*="position: fixed"], [style*="position: absolute"]');
	if (shell) {
		shell.classList.add('native-powerpoint-doc-editor-light-menu-shell');
	}
}

function isFullscreenFixedDialogLayer(layer: HTMLElement): boolean {
	// HyperlinkDialog puts role="dialog" on the fixed overlay itself; other
	// dialogs use a fixed wrapper with a dialog child (Find/Replace, Page Setup).
	return isFullscreenDialogLayer(
		layer.getAttribute('role'),
		layer.getAttribute('aria-modal'),
		Boolean(layer.querySelector(':scope > [role="dialog"]')),
	);
}

function clampFormattingDropdownToViewport(layer: HTMLElement): void {
	const activeWindow = layer.ownerDocument.defaultView;
	const inlineLeft = Number.parseFloat(layer.style.left);
	const inlineTop = Number.parseFloat(layer.style.top);
	if (!activeWindow || !Number.isFinite(inlineLeft) || !Number.isFinite(inlineTop)) {
		return;
	}

	const rect = layer.getBoundingClientRect();
	const next = calculateClampedFloatingLayerPosition(
		rect,
		{ width: activeWindow.innerWidth, height: activeWindow.innerHeight },
		{ left: inlineLeft, top: inlineTop },
	);
	if (Math.abs(next.left - inlineLeft) > 0.5) {
			layer.setCssProps({ left: `${next.left}px` });
	}
	if (Math.abs(next.top - inlineTop) > 0.5) {
			layer.setCssProps({ top: `${next.top}px` });
	}
}

function scheduleFormattingDropdownClamp(layer: HTMLElement): void {
	const clampIfConnected = () => {
		if (layer.isConnected) {
			clampFormattingDropdownToViewport(layer);
		}
	};
	window.requestAnimationFrame(() => {
		window.requestAnimationFrame(clampIfConnected);
	});
	window.setTimeout(clampIfConnected, 100);
}

function syncFormattingBarDropdownState(formattingBar: HTMLElement, open: boolean): void {
	const className = 'native-powerpoint-doc-editor-formatting-dropdown-open';
	const scrollLeftProperty = '--native-powerpoint-doc-editor-formatting-scroll-left';
	const parsedSavedScrollLeft = Number.parseFloat(
		formattingBar.dataset.nativePowerPointDocEditorDropdownScrollLeft ?? '',
	);
	const savedScrollLeft = Number.isFinite(parsedSavedScrollLeft) ? parsedSavedScrollLeft : null;
	const transition = getFormattingDropdownScrollTransition(
		open,
		formattingBar.scrollLeft,
		savedScrollLeft,
	);
	if (open) {
		if (formattingBar.classList.contains(className)) {
			return;
		}

		formattingBar.dataset.nativePowerPointDocEditorDropdownScrollLeft = `${transition.savedScrollLeft ?? 0}`;
		formattingBar.style.setProperty(scrollLeftProperty, `${transition.visualOffset}px`);
		formattingBar.classList.add(className);
		return;
	}

	formattingBar.classList.remove(className);
	formattingBar.style.removeProperty(scrollLeftProperty);
	delete formattingBar.dataset.nativePowerPointDocEditorDropdownScrollLeft;
	if (transition.restoreScrollLeft !== null) {
		formattingBar.scrollLeft = transition.restoreScrollLeft;
	}
}

function normalizeEditorFloatingLayers(editorRoot: HTMLElement) {
	const host = editorRoot.closest('.native-powerpoint-doc-editor-host') ?? editorRoot;
	host.querySelectorAll<HTMLElement>('div[style*="position: fixed"]').forEach((layer) => {
		const isDialogLayer = isFullscreenFixedDialogLayer(layer);
		layer.classList.toggle('native-powerpoint-doc-editor-fixed-dialog-layer', isDialogLayer);
		// The editor positions these layers with inline top/left/transform, which a
		// stylesheet rule cannot override. Pin dialog layers to the viewport inline.
		if (isDialogLayer) {
			layer.setCssProps({ inset: '0', transform: 'none' });
			layer.dataset.nativePowerPointDocEditorFixedLayerPinned = 'true';
		} else if (layer.dataset.nativePowerPointDocEditorFixedLayerPinned === 'true') {
			layer.style.removeProperty('inset');
			layer.style.removeProperty('transform');
			delete layer.dataset.nativePowerPointDocEditorFixedLayerPinned;
		}

		if (!isDialogLayer && layer.closest(DOCX_EDITOR_FORMATTING_BAR_SELECTOR)) {
			scheduleFormattingDropdownClamp(layer);
		}
	});

	const formattingBar = editorRoot.querySelector<HTMLElement>(DOCX_EDITOR_FORMATTING_BAR_SELECTOR);
	const hasFormattingBarDropdown = Boolean(
		formattingBar?.querySelector(':scope div[style*="position: fixed"]'),
	);
	if (formattingBar) {
		syncFormattingBarDropdownState(formattingBar, hasFormattingBarDropdown);
		if (hasFormattingBarDropdown) {
			stripFormattingDropdownButtonTitles(formattingBar);
		}
	}

	suppressVendorToolbarTooltips(editorRoot);

	const hasFloatingMenu = Boolean(
		editorRoot.querySelector(`${DOCX_EDITOR_MENUBAR_SELECTOR} [style*="position: fixed"], ${DOCX_EDITOR_MENUBAR_SELECTOR} [style*="position: absolute"]`),
	);
	editorRoot.classList.toggle('native-powerpoint-doc-editor-has-floating-menu', hasFloatingMenu);
}

function isFontDropdownListbox(listbox: HTMLElement) {
	const optionLabels = Array.from(
		listbox.querySelectorAll<HTMLElement>('[role="option"]'),
	)
		.filter((option) => (
			!option.hasAttribute('data-native-powerpoint-doc-editor-import-font-option')
			&& !option.hasAttribute('data-native-power-point-doc-editor-import-font-option')
		))
		.map((option) => option.textContent?.replace(/\s+/g, ' ').trim().toLowerCase() ?? '')
		.filter(Boolean);

	return optionLabels.includes('arial')
		&& optionLabels.includes('times new roman')
		&& optionLabels.includes('courier new');
}

function resolveFontOptionByName(fontName: string, fonts: FontOption[]) {
	const normalized = fontName.trim().toLowerCase();
	return fonts.find((font) => font.name.trim().toLowerCase() === normalized);
}

function normalizeFontFamilyName(value: string | null | undefined): string {
	return (value ?? '').replace(/\s+/g, ' ').trim();
}

function resolveFontFamilyDisplayName(fontFamily: string, fonts: FontOption[]): string {
	const normalized = normalizeFontFamilyName(fontFamily).toLowerCase();
	if (!normalized) {
		return '';
	}

	const option = fonts.find((font) => {
		const optionName = normalizeFontFamilyName(font.name).toLowerCase();
		const optionFamily = normalizeFontFamilyName(font.fontFamily).toLowerCase();
		const optionPrimary = normalizeFontFamilyName(parsePrimaryFontFamily(font.fontFamily)).toLowerCase();
		return optionName === normalized
			|| optionFamily === normalized
			|| optionPrimary === normalized;
	});

	const primaryName = normalizeFontFamilyName(parsePrimaryFontFamily(fontFamily));
	return option?.name ?? (primaryName || fontFamily);
}

function isFontFamilySelectTrigger(candidate: HTMLElement, fonts: FontOption[]): boolean {
	if (candidate.dataset[FONT_FAMILY_TRIGGER_DATA_KEY] === 'true') {
		return true;
	}

	if (candidate.getAttribute('aria-label') === 'Select font family') {
		return true;
	}

	const label = normalizeFontFamilyName(candidate.textContent).toLowerCase();
	return Boolean(label && fonts.some((font) => font.name.trim().toLowerCase() === label));
}

function getFontFamilySelectTrigger(editorRoot: HTMLElement, fonts: FontOption[]): HTMLElement | null {
	const tagged = editorRoot.querySelector<HTMLElement>(`[${FONT_FAMILY_TRIGGER_ATTRIBUTE}]`);
	if (tagged) {
		return tagged;
	}

	const formattingBar = editorRoot.querySelector<HTMLElement>(DOCX_EDITOR_FORMATTING_BAR_SELECTOR);
	if (!formattingBar) {
		return null;
	}

	const candidates = formattingBar.querySelectorAll<HTMLElement>(
		'button[role="combobox"], button[aria-haspopup="listbox"], button[aria-label]',
	);
	return Array.from(candidates).find((candidate) => isFontFamilySelectTrigger(candidate, fonts)) ?? null;
}

function getFontFamilyTriggerValueElement(trigger: HTMLElement): HTMLElement | null {
	return Array.from(trigger.children).find((child): child is HTMLElement => {
		return isHTMLElement(child)
			&& child.tagName.toLowerCase() !== 'svg'
			&& normalizeFontFamilyName(child.textContent).length > 0;
	}) ?? null;
}

function getFontFamilyTriggerValueTextNode(trigger: HTMLElement): Text | null {
	return Array.from(trigger.childNodes).find((child): child is Text => {
		return child.nodeType === Node.TEXT_NODE
			&& normalizeFontFamilyName(child.textContent).length > 0;
	}) ?? null;
}

function syncFontFamilySelectDisplay(
	editorRoot: HTMLElement | null,
	fontFamily: string,
	fonts: FontOption[],
): boolean {
	if (!editorRoot) {
		return false;
	}

	const displayName = resolveFontFamilyDisplayName(fontFamily, fonts);
	const trigger = getFontFamilySelectTrigger(editorRoot, fonts);
	if (!displayName || !trigger) {
		return false;
	}

	trigger.dataset[FONT_FAMILY_TRIGGER_DATA_KEY] = 'true';

	const valueElement = getFontFamilyTriggerValueElement(trigger);
	if (valueElement) {
		valueElement.textContent = displayName;
	} else {
		const valueText = getFontFamilyTriggerValueTextNode(trigger);
		if (valueText) {
			valueText.textContent = displayName;
		} else if (normalizeFontFamilyName(trigger.textContent) !== displayName) {
			trigger.prepend(activeDocument.createTextNode(displayName));
		}
	}

	return true;
}

function scheduleFontFamilySelectDisplaySync(
	editorRoot: HTMLElement | null,
	fontFamily: string,
	fonts: FontOption[],
	shouldSync: () => boolean = () => true,
): void {
	const sync = () => {
		if (!shouldSync()) {
			return false;
		}
		return syncFontFamilySelectDisplay(editorRoot, fontFamily, fonts);
	};
	sync();
	window.requestAnimationFrame(sync);
	window.setTimeout(sync, 0);
	window.setTimeout(sync, 120);
	window.setTimeout(sync, 320);
}

interface TextSelectionRange {
	from: number;
	to: number;
}

function clampTextSelectionRange(doc: ProseMirrorNode, range: TextSelectionRange): TextSelectionRange {
	const docSize = doc.content.size;
	const from = Math.max(0, Math.min(range.from, docSize));
	const to = Math.max(from, Math.min(range.to, docSize));
	return { from, to };
}

function rememberTextSelectionFromView(view: EditorView): TextSelectionRange | null {
	const { empty, from, to } = view.state.selection;
	if (empty) {
		return null;
	}

	return clampTextSelectionRange(view.state.doc, { from, to });
}

function getFontFamilyFromMark(mark: Mark | undefined): string | null {
	if (!mark) {
		return null;
	}

	const ascii: unknown = mark.attrs.ascii;
	if (typeof ascii === 'string') {
		return ascii;
	}

	const hAnsi: unknown = mark.attrs.hAnsi;
	return typeof hAnsi === 'string' ? hAnsi : null;
}

function getFontFamilyNameFromEditorSelection(view: EditorView): string | null {
	const stored = view.state.storedMarks?.find((mark) => mark.type.name === 'fontFamily');
	if (stored) {
		return getFontFamilyFromMark(stored);
	}

	let activeMarks: readonly Mark[] = view.state.selection.$from.marks();
	if (!view.state.selection.empty) {
		view.state.doc.nodesBetween(view.state.selection.from, view.state.selection.to, (node) => {
			if (!node.isText) {
				return;
			}

			activeMarks = node.marks;
			return false;
		});
	}

	const mark = activeMarks.find((candidate) => candidate.type.name === 'fontFamily');
	if (mark) {
		return getFontFamilyFromMark(mark);
	}

	const paragraph = view.state.selection.$from.parent;
	if (paragraph.type.name === 'paragraph' && paragraph.content.size === 0) {
		const defaults = paragraph.attrs.defaultTextFormatting as {
			fontFamily?: { ascii?: unknown; hAnsi?: unknown };
		} | null | undefined;
		const fontFamily = defaults?.fontFamily?.ascii ?? defaults?.fontFamily?.hAnsi;
		if (typeof fontFamily === 'string' && fontFamily.trim()) {
			return fontFamily;
		}
	}

	try {
		const dom = view.domAtPos(view.state.selection.from);
		const element = isElement(dom.node) ? dom.node : dom.node.parentElement;
		const cssFont = element?.ownerDocument.defaultView?.getComputedStyle(element).fontFamily;
		return parsePrimaryFontFamily(cssFont ?? '');
	} catch {
		return null;
	}
}

function getActiveTextSelectionRange(
	view: EditorView,
	preserved: TextSelectionRange | null,
): TextSelectionRange | null {
	if (!view.state.selection.empty) {
		return clampTextSelectionRange(view.state.doc, {
			from: view.state.selection.from,
			to: view.state.selection.to,
		});
	}

	return preserved;
}

function summarizeSelectedEmptyParagraphFonts(
	doc: ProseMirrorNode,
	range: TextSelectionRange | null,
	fontFamily: string,
): { emptyParagraphs: number; matchingFontFamily: number } {
	if (!range || range.to <= range.from) {
		return { emptyParagraphs: 0, matchingFontFamily: 0 };
	}

	let emptyParagraphs = 0;
	let matchingFontFamily = 0;
	const target = normalizeFontFamilyName(fontFamily).toLowerCase();
	doc.nodesBetween(range.from, range.to, (node) => {
		if (node.type.name !== 'paragraph' || node.content.size !== 0) {
			return;
		}
		emptyParagraphs += 1;
		const defaults = node.attrs.defaultTextFormatting as {
			fontFamily?: { ascii?: unknown; hAnsi?: unknown };
		} | null | undefined;
		const family = defaults?.fontFamily?.ascii ?? defaults?.fontFamily?.hAnsi;
		if (typeof family === 'string' && normalizeFontFamilyName(family).toLowerCase() === target) {
			matchingFontFamily += 1;
		}
	});

	return { emptyParagraphs, matchingFontFamily };
}

function isFontPickerMenuItem(target: Element): HTMLElement | null {
	const item = target.closest<HTMLElement>('[role="option"]');
	if (!item || item.closest('[data-native-powerpoint-doc-editor-import-font-option], [data-native-power-point-doc-editor-import-font-option]')) {
		return null;
	}

	if (item.closest('[data-native-powerpoint-doc-editor-font-menu-decorated], [data-native-power-point-doc-editor-font-menu-decorated]')) {
		return item;
	}

	const listbox = item.closest<HTMLElement>('[role="listbox"]');
	if (listbox && isFontDropdownListbox(listbox)) {
		return item;
	}

	return null;
}

function applyFontFamilyToEditorView(
	view: EditorView,
	fontFamily: string,
	preservedRange: TextSelectionRange | null,
): { applied: boolean; range: TextSelectionRange | null } {
	const activeRange = getActiveTextSelectionRange(view, preservedRange);

	if (activeRange && (
		activeRange.from !== view.state.selection.from
		|| activeRange.to !== view.state.selection.to
	)) {
		const selection = activeRange.from === 0 && activeRange.to === view.state.doc.content.size
			? new AllSelection(view.state.doc)
			: TextSelection.create(view.state.doc, activeRange.from, activeRange.to);
		view.dispatch(view.state.tr.setSelection(selection));
	}

	// Use the core command for both a range and an empty paragraph. It updates
	// text marks plus defaultTextFormatting for every selected empty paragraph;
	// raw addMark() cannot represent those empty lines.
	const applied = setFontFamily(fontFamily)(view.state, view.dispatch);
	const $from = view.state.selection.$from;
	const emptyPara = $from.parent.isTextblock && $from.parent.textContent.length === 0;
	if (emptyPara) {
		debugLog('editor', 'DOCX empty-paragraph font family applied', {
			fontFamily,
			applied,
			hasDefaultTextFormatting: Boolean($from.parent.attrs.defaultTextFormatting),
		});
	}
	return { applied, range: activeRange };
}

function tagFontFamilySelectTrigger(container: HTMLElement) {
	const contentId = container.id;
	if (!contentId) {
		return;
	}

	const trigger = activeDocument.querySelector<HTMLElement>(`[aria-controls="${CSS.escape(contentId)}"]`);
	if (trigger) {
		trigger.dataset[FONT_FAMILY_TRIGGER_DATA_KEY] = 'true';
	}
}

function scheduleFontFamilySelectTriggerTag(container: HTMLElement) {
	let attempts = 0;
	const tryTag = () => {
		tagFontFamilySelectTrigger(container);
		if (!container.id && attempts < 5) {
			attempts += 1;
			window.requestAnimationFrame(tryTag);
		}
	};

	tryTag();
}

function appendImportFontOption(listbox: HTMLElement, onImportFont: () => void) {
	// Radix puts role="listbox" on Select.Content. Append the Import Font footer
	// *inside* that panel (after the scroll viewport), not on the theme portal
	// wrapper — otherwise the row sits outside the bordered dropdown chrome.
	// Use setAttribute: `dataset.nativePowerPoint…` writes `data-native-power-point-…`
	// and breaks `data-native-powerpoint-…` selectors / CSS.
	const decoratedAttr = 'data-native-powerpoint-doc-editor-font-menu-decorated';
	const importAttr = 'data-native-powerpoint-doc-editor-import-font-option';
	if (listbox.getAttribute(decoratedAttr) === 'true') {
		return;
	}

	if (!isFontDropdownListbox(listbox)) {
		return;
	}

	listbox.setAttribute(decoratedAttr, 'true');
	listbox.removeAttribute('data-native-power-point-doc-editor-font-menu-decorated');

	if (
		listbox.querySelector(`[${importAttr}], [data-native-power-point-doc-editor-import-font-option]`)
	) {
		return;
	}

	const footer = listbox.createDiv();
	footer.className = 'native-powerpoint-doc-editor-font-menu-footer';
	footer.setAttribute('data-native-powerpoint-doc-editor-font-menu-footer', 'true');

	const separator = createMenuSection(footer, {
		className: 'native-powerpoint-doc-editor-import-font-separator',
	});
	separator.setAttribute(importAttr, 'true');

	const button = createMenuItem(footer, {
		className: 'native-powerpoint-doc-editor-import-font-option',
		text: IMPORT_FONT_MENU_LABEL,
	});
	button.setAttribute(importAttr, 'true');
	hardenInjectedMenuOption(button, { onSelect: onImportFont });

	scheduleFontFamilySelectTriggerTag(listbox);
}

function clampCustomTableSize(value: number) {
	if (!Number.isFinite(value)) {
		return MIN_CUSTOM_TABLE_SIZE;
	}

	return Math.min(MAX_CUSTOM_TABLE_SIZE, Math.max(MIN_CUSTOM_TABLE_SIZE, Math.round(value)));
}

function isTableSizeGrid(grid: HTMLElement) {
	const gridCells = grid.querySelectorAll('[role="gridcell"]');
	if (gridCells.length < 9) {
		return false;
	}

	const label = grid.getAttribute('aria-label')?.toLowerCase() ?? '';
	return label.includes('table') || gridCells.length === 36;
}

function appendCustomTableOption(grid: HTMLElement, onCustomTable: () => void) {
	const container = grid.parentElement;
	if (!container || container.dataset.nativePowerPointDocEditorTableSizeMenuDecorated === 'true') {
		return;
	}

	if (!isTableSizeGrid(grid)) {
		return;
	}

	container.dataset.nativePowerPointDocEditorTableSizeMenuDecorated = 'true';

	if (container.querySelector('[data-native-powerpoint-doc-editor-custom-table-option]')) {
		return;
	}

	markLightMenuSurface(container, 'native-powerpoint-doc-editor-table-size-menu');

	const separator = createMenuSection(container, {
		className: 'native-powerpoint-doc-editor-custom-table-separator',
		role: 'separator',
	});
	separator.dataset.nativePowerPointDocEditorCustomTableOption = 'true';

	const button = createMenuItem(container, {
		className: 'native-powerpoint-doc-editor-custom-table-option',
		text: CUSTOM_TABLE_MENU_LABEL,
	});
	button.dataset.nativePowerPointDocEditorCustomTableOption = 'true';
	hardenInjectedMenuOption(button, { onSelect: onCustomTable });
}

function getPlainTextFromKeyboardEvent(event: KeyboardEvent) {
	if (event.metaKey || event.ctrlKey || event.altKey || event.isComposing) {
		return null;
	}

	if (event.key === ' ' || event.key === 'Spacebar') {
		return ' ';
	}

	return event.key.length === 1 ? event.key : null;
}

function getPlainTextFromInputEvent(event: InputEvent) {
	if (event.isComposing || event.inputType !== 'insertText' || !event.data || /[\r\n]/.test(event.data)) {
		return null;
	}

	return event.data;
}

const preserveTypedSpacePlugin = new Plugin({
	key: preserveTypedSpacePluginKey,
	props: {
		handleDOMEvents: {
			beforeinput(view, event) {
				if (!isInputEvent(event)) {
					return false;
				}

				const text = getPlainTextFromInputEvent(event);
				if (!text) {
					return false;
				}

				const defaultPreventedBefore = event.defaultPrevented;
				const { from, to } = view.state.selection;
				event.preventDefault();
				const handled = insertPlainTypedText(view, text, from, to);
				if (text === ' ') {
					logDocxSpaceInputRoute('beforeinput', {
						from,
						to,
						inputType: event.inputType,
						isComposing: event.isComposing,
						defaultPreventedBefore,
						defaultPreventedAfter: event.defaultPrevented,
					});
				}
				return handled;
			},
			paste(view, event) {
				if (!isClipboardEvent(event)) {
					return false;
				}

				const html = event.clipboardData?.getData('text/html') ?? '';
				const hasHtml = Boolean(html.trim());
				const plainText = getPlainTextFromClipboardEvent(event);
				const { from, to } = view.state.selection;
				const $from = view.state.doc.resolve(from);
				const $to = view.state.doc.resolve(to);
				const crossBlock = from !== to && !$from.sameParent($to);
				const multiLine = /[\r\n]/.test(plainText);

				// The editor's own copy command serializes a ProseMirror slice in
				// `text/html`, which is the only clipboard representation that
				// retains marks such as bold, italics, and font formatting. Let
				// ProseMirror parse and replace this slice for every selection
				// shape; routing multi-line or cross-block HTML through the plain
				// text helper silently strips those marks.
				if (hasHtml) {
					debugLog('clipboard', 'DOCX paste deferred to HTML clipboardParser', {
						selectionStart: from,
						selectionEnd: to,
						sameParent: $from.sameParent($to),
						crossBlock,
						plainTextLength: plainText?.length ?? 0,
						hasNewline: multiLine,
						htmlLength: html.length,
					});
					return false;
				}

				if (!plainText) {
					return false;
				}

				const marker = getCurrentParagraphListMarker(view);
				const strippedText = stripMatchingListMarkerPrefixFromPastedText(plainText, marker);
				const pasteText = strippedText ?? plainText;

				debugLog('clipboard', 'DOCX plain paste routed through safe insert', {
					branch: strippedText !== null
						? 'list-marker-strip'
						: (multiLine ? 'multi-line-paragraphs' : (crossBlock ? 'cross-block' : 'single-line')),
					selectionStart: from,
					selectionEnd: to,
					sameParent: $from.sameParent($to),
					crossBlock,
					textLength: pasteText.length,
					hasNewline: multiLine,
				});
				event.preventDefault();
				return insertPlainTypedText(view, pasteText, from, to);
			},
		},
		handlePaste(view, event, slice: Slice) {
			// Let tracked-changes own paste while suggestions are active. Its
			// transaction adds insertion/deletion marks that a direct replaceRange
			// would bypass.
			if (isSuggestionModeActive(view.state)) {
				return false;
			}

			const html = event.clipboardData?.getData('text/html') ?? '';
			if (!html.trim()) {
				return false;
			}

			const { from, to } = view.state.selection;
			const $from = view.state.doc.resolve(from);
			const $to = view.state.doc.resolve(to);
			const crossBlock = from !== to && !$from.sameParent($to);
			const multiLine = /[\r\n]/.test(event.clipboardData?.getData('text/plain') ?? '');
			// A collapsed cursor (including a multi-line paste) must use
			// ProseMirror's native `replaceSelection` behavior. Its open slice
			// merges at the cursor in source order; forcing it through the
			// cross-block replacement path can place an opened blank paragraph
			// before the section heading. Only an actual cross-block selection
			// needs the closed, structured replacement below.
			if (!crossBlock) {
				return false;
			}

			const sliceSummary = summarizeRichClipboardSlice(slice);
			const range = insertRichClipboardSlice(view, slice, from, to);
			debugLog('clipboard', 'DOCX rich paste replaced through structured slice', {
				selectionStart: from,
				selectionEnd: to,
				crossBlock,
				multiLine,
				replaceFrom: range.from,
				replaceTo: range.to,
				sliceSize: slice.size,
				openStart: slice.openStart,
				openEnd: slice.openEnd,
				...sliceSummary,
			});
			return true;
		},
		handleKeyDown(view, event) {
			if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'z') {
				const { from, to } = view.state.selection;
				debugLog('history', 'DOCX keyboard history shortcut received', {
					action: event.shiftKey ? 'redo' : 'undo',
					selectionStart: from,
					selectionEnd: to,
					defaultPrevented: event.defaultPrevented,
				});
				// Let the core HistoryExtension consume the shortcut.
				return false;
			}

			if (event.key === 'Backspace' || event.key === 'Delete') {
				const { selection } = view.state;
				const paragraph = selection.$from.parent;
				const paragraphAttrs = paragraph.attrs as { numPr?: { ilvl?: number } | null };
				const listProperties = paragraph.type.name === 'paragraph' ? paragraphAttrs.numPr : null;
				debugLog('text-input', 'DOCX deletion key received', {
					key: event.key,
					code: event.code,
					selectionStart: selection.from,
					selectionEnd: selection.to,
					parentOffset: selection.$from.parentOffset,
					isList: Boolean(listProperties),
					listLevel: listProperties?.ilvl ?? null,
					suggestionMode: isSuggestionModeActive(view.state),
				});
				return false;
			}

			const text = getPlainTextFromKeyboardEvent(event);
			if (!text) {
				return false;
			}

			const defaultPreventedBefore = event.defaultPrevented;
			const { from, to } = view.state.selection;
			event.preventDefault();
			const handled = insertPlainTypedText(view, text, from, to);
			if (text === ' ') {
				logDocxSpaceInputRoute('keydown', {
					from,
					to,
					key: event.key,
					code: event.code,
					isComposing: event.isComposing,
					defaultPreventedBefore,
					defaultPreventedAfter: event.defaultPrevented,
				});
			}
			return handled;
		},
		handleTextInput(view, from, to, text) {
			if (!text || /[\r\n]/.test(text)) {
				return false;
			}

			const handled = insertPlainTypedText(view, text, from, to);
			if (text === ' ') {
				logDocxSpaceInputRoute('text-input', {
					from,
					to,
					isComposing: view.composing,
				});
			}
			return handled;
		},
	},
});

const contentShrinkDiagnosticsPlugin = new Plugin({
	key: contentShrinkDiagnosticsPluginKey,
	appendTransaction(transactions, oldState, newState) {
		if (!transactions.some((transaction) => transaction.docChanged)) {
			return null;
		}

		const paragraphsBefore = countDocTextblocks(oldState.doc);
		const paragraphsAfter = countDocTextblocks(newState.doc);
		const charactersBefore = oldState.doc.textContent.length;
		const charactersAfter = newState.doc.textContent.length;
		if (paragraphsAfter >= paragraphsBefore && charactersAfter >= charactersBefore) {
			return null;
		}

		const { $from, $to, from, to } = oldState.selection;
		debugLog('editor', 'DOCX content shrunk', {
			paragraphsBefore,
			paragraphsAfter,
			charactersBefore,
			charactersAfter,
			selectionStart: from,
			selectionEnd: to,
			sameParent: $from.sameParent($to),
			steps: transactions.flatMap((transaction) => summarizeTransactionSteps(transaction)),
			meta: transactions.flatMap((transaction) => summarizeTransactionMeta(transaction)),
			docChangedCount: transactions.filter((transaction) => transaction.docChanged).length,
		});
		return null;
	},
});

function clampZoom(zoom: number) {
	return Math.max(MIN_TOUCH_ZOOM, Math.min(MAX_TOUCH_ZOOM, Math.round(zoom * 100) / 100));
}

function scaleTouchZoom(startZoom: number, rawScale: number) {
	if (!Number.isFinite(rawScale) || rawScale <= 0) {
		return startZoom;
	}

	return clampZoom(startZoom * Math.pow(rawScale, TOUCH_ZOOM_SENSITIVITY));
}

function getTouchDistance(first: Touch, second: Touch) {
	return Math.hypot(first.clientX - second.clientX, first.clientY - second.clientY);
}

function getTouchCenter(first: Touch, second: Touch) {
	return {
		x: (first.clientX + second.clientX) / 2,
		y: (first.clientY + second.clientY) / 2,
	};
}

function getPointDistance(first: PointerPoint, second: PointerPoint) {
	return Math.hypot(first.x - second.x, first.y - second.y);
}

function getPointCenter(first: PointerPoint, second: PointerPoint) {
	return {
		x: (first.x + second.x) / 2,
		y: (first.y + second.y) / 2,
	};
}

function getScrollableEditorElement(root: HTMLElement) {
	const pages = root.querySelector<HTMLElement>(DOCX_EDITOR_PAGES_SELECTOR);
	let candidate: HTMLElement | null = pages?.parentElement ?? root;

	while (candidate && candidate !== root) {
		const style = window.getComputedStyle(candidate);
		const canScroll = /(auto|scroll)/.test(`${style.overflow} ${style.overflowX} ${style.overflowY}`);
		if (canScroll && (candidate.scrollHeight > candidate.clientHeight || candidate.scrollWidth > candidate.clientWidth)) {
			return candidate;
		}
		candidate = candidate.parentElement;
	}

	return root;
}

function centerEditorViewport(root: HTMLElement) {
	const pages = root.querySelector<HTMLElement>(DOCX_EDITOR_PAGES_SELECTOR);
	if (!pages) {
		return false;
	}

	const scrollContainer = getScrollableEditorElement(root);
	const scrollRect = scrollContainer.getBoundingClientRect();
	const pagesRect = pages.getBoundingClientRect();
	if (scrollRect.width <= 0 || pagesRect.width <= 0) {
		return false;
	}

	const pagesLeftInScrollSpace = pagesRect.left - scrollRect.left + scrollContainer.scrollLeft;
	const pagesCenter = pagesLeftInScrollSpace + (pagesRect.width / 2);
	const nextScrollLeft = Math.max(0, pagesCenter - (scrollContainer.clientWidth / 2));
	if (!Number.isFinite(nextScrollLeft)) {
		return false;
	}

	scrollContainer.scrollLeft = nextScrollLeft;
	return true;
}

function shouldEnableTouchPinchZoom() {
	if (typeof window === 'undefined' || typeof navigator === 'undefined') {
		return false;
	}

	return Platform.isMobile || Platform.isMobileApp || (navigator.maxTouchPoints >= 2 && window.matchMedia('(hover: none)').matches);
}

function getEditorModeFromButton(button: HTMLButtonElement): EditorMode | null {
	const label = button.textContent?.replace(/\s+/g, ' ').trim().toLowerCase() ?? '';

	if (label.startsWith('bearbeiten') || label.startsWith('edit')) {
		return 'editing';
	}
	if (label.startsWith('vorschlagen') || label.startsWith('suggest')) {
		return 'suggesting';
	}
	if (label.startsWith('anzeigen') || label.startsWith('view')) {
		return 'viewing';
	}

	return null;
}

function escapeRegExp(value: string) {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function getVisibleListMarker(attrs: Record<string, unknown>) {
	if (attrs.listMarkerHidden) {
		return '';
	}

	const marker = attrs.listMarker;
	return typeof marker === 'string' ? marker.trimEnd() : '';
}

function getCurrentParagraphListMarker(view: EditorView) {
	const { $from } = view.state.selection;
	for (let depth = $from.depth; depth >= 0; depth -= 1) {
		const node = $from.node(depth);
		if (node.isTextblock) {
			return getVisibleListMarker(node.attrs);
		}
	}

	return '';
}

function stripMatchingListMarkerPrefixFromPastedText(text: string, marker: string) {
	if (!marker || !text) {
		return null;
	}

	const singleLineMatch = text.match(/^([^\r\n]*)(?:\r\n|\n|\r)?$/);
	if (!singleLineMatch) {
		return null;
	}

	const line = singleLineMatch[1] ?? '';
	const markerPrefixPattern = new RegExp(`^${escapeRegExp(marker)}(?:\\t|\\s)+`);
	if (!markerPrefixPattern.test(line)) {
		return null;
	}

	return line.replace(markerPrefixPattern, '');
}

function getPlainTextFromClipboardEvent(event: ClipboardEvent) {
	return event.clipboardData?.getData('text/plain') ?? event.clipboardData?.getData('Text') ?? '';
}

function clearListMarkerSelectionHighlights(root: HTMLElement) {
	root.querySelectorAll<HTMLElement>(`.${SELECTED_LIST_MARKER_CLASS}`).forEach((marker) => {
		marker.classList.remove(SELECTED_LIST_MARKER_CLASS);
	});
}

function updateListMarkerSelectionHighlights(root: HTMLElement, view: EditorView) {
	clearListMarkerSelectionHighlights(root);

	const { selection } = view.state;
	if (selection.empty) {
		return;
	}

	const selectionStart = Math.min(selection.from, selection.to);
	const selectionEnd = Math.max(selection.from, selection.to);
	root.querySelectorAll<HTMLElement>(LIST_PARAGRAPH_SELECTOR).forEach((paragraph) => {
		const marker = paragraph.querySelector<HTMLElement>(LIST_MARKER_SELECTOR);
		if (!marker) {
			return;
		}

		const paragraphStart = Number(paragraph.dataset.pmStart);
		if (!Number.isFinite(paragraphStart)) {
			return;
		}

		if (selectionStart <= paragraphStart + 1 && selectionEnd > paragraphStart) {
			marker.classList.add(SELECTED_LIST_MARKER_CLASS);
		}
	});
}

function splitClipboardLines(text: string) {
	return text.split(/\r\n|\n|\r/);
}

function addListMarkersToPlainText(view: EditorView, text: string) {
	if (!text.trim()) {
		return null;
	}

	const selectedLines = splitClipboardLines(text);
	const textToMarker = new Map<string, string>();
	view.state.doc.descendants((node) => {
		if (!node.isTextblock) {
			return true;
		}

		const marker = getVisibleListMarker(node.attrs);
		const paragraphText = node.textContent;
		if (marker.length > 0 && paragraphText.length > 0 && !textToMarker.has(paragraphText)) {
			textToMarker.set(paragraphText, marker);
		}

		return false;
	});

	let includedMarker = false;
	const lines = selectedLines.map((line) => {
		const marker = textToMarker.get(line);
		if (!marker || line.startsWith(marker)) {
			return line;
		}

		includedMarker = true;
		return `${marker}\t${line}`;
	});

	return includedMarker ? lines.join('\n') : null;
}

interface ElectronClipboard {
	readHTML?: () => string;
	readRTF?: () => string;
	readText: () => string;
	write?: (data: { html?: string; rtf?: string; text?: string }) => void;
	writeText: (text: string) => void;
}

function getElectronClipboard() {
	try {
		const runtimeRequire = (window as ObsidianDesktopWindow).require;
		if (typeof runtimeRequire !== 'function') {
			return null;
		}

		const electron = runtimeRequire('electron') as {
			clipboard?: ElectronClipboard;
		};
		return electron.clipboard ?? null;
	} catch {
		return null;
	}
}

async function readPlainTextClipboard() {
	const electronClipboard = getElectronClipboard();
	if (electronClipboard) {
		return electronClipboard.readText();
	}

	return await navigator.clipboard.readText();
}

function readHtmlClipboard() {
	return getElectronClipboard()?.readHTML?.() ?? '';
}

async function writePlainTextClipboard(text: string) {
	const electronClipboard = getElectronClipboard();
	if (electronClipboard) {
		const html = electronClipboard.readHTML?.() ?? '';
		const rtf = electronClipboard.readRTF?.() ?? '';
		if (electronClipboard.write && (html.length > 0 || rtf.length > 0)) {
			electronClipboard.write({
				text,
				...(html.length > 0 ? { html } : {}),
				...(rtf.length > 0 ? { rtf } : {}),
			});
		} else {
			electronClipboard.writeText(text);
		}
		return;
	}

	await navigator.clipboard.writeText(text);
}

async function rewritePlainTextClipboardWithListMarkers(view: EditorView) {
	try {
		const clipboardText = await readPlainTextClipboard();
		const listAwareText = addListMarkersToPlainText(view, clipboardText);
		if (!listAwareText || listAwareText === clipboardText) {
			return false;
		}

		await writePlainTextClipboard(listAwareText);
		debugLog('clipboard', 'Rewrote clipboard text with list markers', {
			originalTextLength: clipboardText.length,
			listAwareTextLength: listAwareText.length,
		});
		return true;
	} catch (error) {
		debugLog('clipboard', 'Could not rewrite clipboard text with list markers', error);
		return false;
	}
}

function getPasteTextWithListMarkerGuard(view: EditorView, text: string) {
	const marker = getCurrentParagraphListMarker(view);
	return stripMatchingListMarkerPrefixFromPastedText(text, marker) ?? text;
}

async function pasteClipboardIntoEditor(view: EditorView, options: PasteClipboardOptions) {
	view.focus();
	const text = await readPlainTextClipboard();
	if (!options.preserveFormatting) {
		if (!text) return false;
		const docBefore = view.state.doc;
		insertPlainTypedText(view, getPasteTextWithListMarkerGuard(view, text));
		return view.state.doc !== docBefore;
	}

	const html = readHtmlClipboard();
	if (html.trim()) {
		try {
			const docBefore = view.state.doc;
			const clipboardData = new DataTransfer();
			clipboardData.setData('text/html', html);
			clipboardData.setData('text/plain', text);
			const pasteEvent = new ClipboardEvent('paste', {
				bubbles: true,
				cancelable: true,
				clipboardData,
			});
			const defaultAllowed = view.dom.dispatchEvent(pasteEvent);
			if ((pasteEvent.defaultPrevented || !defaultAllowed) && view.state.doc !== docBefore) {
				return true;
			}
			debugLog('clipboard', 'Formatted DOCX paste made no document change; trying plain text', {
				htmlLength: html.length,
				plainTextLength: text.length,
				defaultPrevented: pasteEvent.defaultPrevented,
				defaultAllowed,
			});
		} catch (error) {
			debugLog('clipboard', 'Could not dispatch formatted paste event; falling back to plain text paste', error);
		}
	}

	if (!text) return false;
	const docBefore = view.state.doc;
	insertPlainTypedText(view, getPasteTextWithListMarkerGuard(view, text));
	return view.state.doc !== docBefore;
}

export function ensureEditorStyles() {
	if (stylesInjected) {
		return;
	}

	// After vendor tokens: pin --doc-caret to page ink. Vendor `.dark` uses a
	// light caret for inverted canvas; Obsidian keeps Word-white pages.
	const hostCaretOverride = `
.docx-editor-root.docx-editor,
.docx-editor-root.docx-editor.dark,
.docx-editor-root.dark {
	--doc-caret: #000000;
}
.docx-editor-root.dark .layout-page,
.docx-editor-root.dark .layout-page img,
.docx-editor-root.dark .layout-page svg,
.docx-editor-root.dark .layout-page canvas,
.docx-editor-root.dark .layout-page video,
.docx-editor-root.dark .layout-page [data-no-color-invert] {
	filter: none;
}
`.trim();

	const useAdoptedSheets =
		typeof CSSStyleSheet !== 'undefined'
		&& 'adoptedStyleSheets' in activeDocument
		&& Array.isArray(activeDocument.adoptedStyleSheets);

	if (!useAdoptedSheets) {
		throw new Error('DOCX editor styles require adoptedStyleSheets (Obsidian Chromium).');
	}

	const styleSheet = new CSSStyleSheet();
	styleSheet.replaceSync(docxEditorStyles);
	const caretSheet = new CSSStyleSheet();
	caretSheet.replaceSync(hostCaretOverride);
	activeDocument.adoptedStyleSheets = [
		...activeDocument.adoptedStyleSheets,
		styleSheet,
		caretSheet,
	];
	stylesInjected = true;
	debugLog('editor', 'DOCX editor styles injected with white-page caret override', {
		via: 'adoptedStyleSheets',
		caretOverride: '--doc-caret: #000000; layout-page filter: none',
	});
}

const SaveButton = ({ onClick }: { onClick: () => void }) => {
	const ref = useRef<HTMLButtonElement>(null);
	useEffect(() => {
		if (ref.current) {
			configureToolbarIconButton(ref.current, {
				className: 'native-powerpoint-doc-editor-logo-save-button',
				icon: 'save',
				label: 'Save',
			});
		}
	}, []);

	return (
		<button
			ref={ref}
			type="button"
			className="clickable-icon native-powerpoint-doc-editor-logo-save-button"
			onClick={onClick}
			aria-label="Save"
			style={{
				background: 'transparent',
				border: 'none',
				boxShadow: 'none',
				padding: '4px 8px',
				cursor: 'pointer',
				display: 'flex',
				alignItems: 'center',
				justifyContent: 'center',
				color: 'inherit'
			}}
		/>
	);
};

const SAVE_STATUS_LABELS: Record<SaveStatus, string> = {
	saved: 'Saved',
	saving: 'Saving...',
	unsaved: 'Unsaved',
	failed: 'Save failed',
};

const SaveStatusIndicator = ({ status }: { status: SaveStatus }) => {
	const canonicalState = DOCX_SAVE_STATUS_TO_STATE[status];
	const { busy } = getSaveStatusFlags(canonicalState);
	return (
		<span
			className={`native-powerpoint-doc-editor-save-status native-powerpoint-doc-editor-save-status-${status}`}
			data-state={canonicalState}
			role="status"
			aria-live="polite"
			aria-busy={busy ? 'true' : 'false'}
			title={SAVE_STATUS_LABELS[status]}
		>
			<span className="native-powerpoint-doc-editor-save-status-dot" aria-hidden="true" />
			{SAVE_STATUS_LABELS[status]}
		</span>
	);
};

interface FindReplaceDialogProps {
	isOpen: boolean;
	labels: FindReplaceLabels;
	mode: FindReplaceMode;
	searchText: string;
	replaceText: string;
	matchCase: boolean;
	wholeWord: boolean;
	matchCount: number;
	currentIndex: number;
	onSearchTextChange: (value: string) => void;
	onReplaceTextChange: (value: string) => void;
	onMatchCaseChange: (value: boolean) => void;
	onWholeWordChange: (value: boolean) => void;
	onModeChange: (mode: FindReplaceMode) => void;
	onNext: () => void;
	onPrevious: () => void;
	onReplace: () => void;
	onReplaceAll: () => void;
	onClose: () => void;
}

const FindReplaceDialog = ({
	isOpen,
	labels,
	mode,
	searchText,
	replaceText,
	matchCase,
	wholeWord,
	matchCount,
	currentIndex,
	onSearchTextChange,
	onReplaceTextChange,
	onMatchCaseChange,
	onWholeWordChange,
	onModeChange,
	onNext,
	onPrevious,
	onReplace,
	onReplaceAll,
	onClose,
}: FindReplaceDialogProps) => {
	if (!isOpen) {
		return null;
	}

	const resultText = formatFindResultStatus(searchText, currentIndex, matchCount, {
		noSearch: '',
		noMatches: labels.noMatches,
		resultCount: labels.resultCount,
	});

	return (
		<div
			className="native-powerpoint-doc-editor-find-dialog"
			style={{
				position: 'fixed',
				right: '24px',
				top: '92px',
				zIndex: 100050,
				width: '360px',
				background: 'white',
				border: '1px solid var(--background-modifier-border, #d1d5db)',
				borderRadius: '8px',
				boxShadow: '0 10px 30px rgba(0, 0, 0, 0.18)',
				padding: '12px',
				color: 'var(--text-normal, #202124)',
			}}
		>
			<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '10px' }}>
				<strong>{mode === 'replace' ? labels.findAndReplace : labels.find}</strong>
				<button type="button" aria-label={labels.close} onClick={onClose} style={{ border: 'none', background: 'transparent', cursor: 'pointer' }}>×</button>
			</div>
			<div style={{ display: 'flex', gap: '6px', marginBottom: '8px' }}>
				<input
					value={searchText}
					onChange={(evt) => onSearchTextChange(evt.currentTarget.value)}
					placeholder={labels.findText} aria-label={labels.findText}
					autoFocus
					style={{ flex: 1, height: '30px' }}
					onKeyDown={(evt) => {
						if (evt.key === 'Enter') {
							evt.preventDefault();
							if (evt.shiftKey) {
								onPrevious();
							} else {
								onNext();
							}
						}
					}}
				/>
				<button type="button" aria-label={labels.previous} title={labels.previous} onClick={onPrevious} disabled={matchCount === 0}>↑</button>
				<button type="button" aria-label={labels.next} title={labels.next} onClick={onNext} disabled={matchCount === 0}>↓</button>
			</div>
			{mode === 'replace' && (
				<div style={{ display: 'flex', gap: '6px', marginBottom: '8px' }}>
					<input
						value={replaceText}
						onChange={(evt) => onReplaceTextChange(evt.currentTarget.value)}
						placeholder={labels.replaceWith} aria-label={labels.replaceWith}
						style={{ flex: 1, height: '30px' }}
					/>
					<button type="button" onClick={onReplace} disabled={matchCount === 0}>{labels.replace}</button>
					<button type="button" onClick={onReplaceAll} disabled={matchCount === 0}>{labels.replaceAll}</button>
				</div>
			)}
			<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px' }}>
				<div style={{ display: 'flex', gap: '10px', fontSize: '12px' }}>
					<label><input type="checkbox" aria-label={labels.matchCase} checked={matchCase} onChange={(evt) => onMatchCaseChange(evt.currentTarget.checked)} /> {labels.matchCase}</label>
					<label><input type="checkbox" aria-label={labels.wholeWords} checked={wholeWord} onChange={(evt) => onWholeWordChange(evt.currentTarget.checked)} /> {labels.wholeWords}</label>
				</div>
				<div style={{ fontSize: '12px', color: 'var(--text-muted, #6b7280)', whiteSpace: 'nowrap' }}>{resultText}</div>
			</div>
			{mode === 'find' && (
				<button type="button" onClick={() => onModeChange('replace')} style={{ marginTop: '10px' }}>{labels.showReplace}</button>
			)}
		</div>
	);
};

interface CustomTableDialogProps {
	isOpen: boolean;
	onClose: () => void;
	onInsert: (rows: number, columns: number) => void;
}

const CustomTableDialog = ({ isOpen, onClose, onInsert }: CustomTableDialogProps) => {
	const [rows, setRows] = useState(3);
	const [columns, setColumns] = useState(3);

	useEffect(() => {
		if (isOpen) {
			setRows(3);
			setColumns(3);
		}
	}, [isOpen]);

	const submit = useCallback(() => {
		onInsert(clampCustomTableSize(rows), clampCustomTableSize(columns));
	}, [columns, onInsert, rows]);

	if (!isOpen) {
		return null;
	}

	return (
		<div
			className="native-powerpoint-doc-editor-custom-table-backdrop"
			onMouseDown={(evt) => {
				if (evt.target === evt.currentTarget) {
					onClose();
				}
			}}
			onKeyDown={(evt) => {
				if (evt.key === 'Escape') {
					evt.preventDefault();
					onClose();
				}
				if (evt.key === 'Enter') {
					evt.preventDefault();
					submit();
				}
			}}
		>
			<div
				className="native-powerpoint-doc-editor-custom-table-dialog"
				role="dialog"
				aria-modal="true"
				aria-label="Custom table"
				onMouseDown={(evt) => evt.stopPropagation()}
			>
				<div className="native-powerpoint-doc-editor-custom-table-title">Custom table</div>
				<label className="native-powerpoint-doc-editor-custom-table-field">
					<span>Rows</span>
					<input
						type="number"
						min={MIN_CUSTOM_TABLE_SIZE}
						max={MAX_CUSTOM_TABLE_SIZE}
						value={rows}
						aria-label="Rows"
						autoFocus
						onChange={(evt) => setRows(clampCustomTableSize(Number(evt.currentTarget.value)))}
					/>
				</label>
				<label className="native-powerpoint-doc-editor-custom-table-field">
					<span>Columns</span>
					<input
						type="number"
						min={MIN_CUSTOM_TABLE_SIZE}
						max={MAX_CUSTOM_TABLE_SIZE}
						value={columns}
						aria-label="Columns"
						onChange={(evt) => setColumns(clampCustomTableSize(Number(evt.currentTarget.value)))}
					/>
				</label>
				<div className="native-powerpoint-doc-editor-custom-table-actions">
					<button type="button" onClick={onClose}>Cancel</button>
					<button type="button" className="native-powerpoint-doc-editor-custom-table-primary" onClick={submit}>Insert</button>
				</div>
			</div>
		</div>
	);
};

export interface DocxReactViewProps {
	file: TFile | null;
	buffer: ArrayBuffer | null;
	documentKey: string;
	editorAdapter: DocxEditorAdapterController;
	error: string | null;
	isLoading: boolean;
	authorName: string;
	resolvedEditorTheme: EditorThemeResolution;
	i18n: Translations | undefined;
	pluginI18n: I18nService | null;
	showNotice: (key: string, values?: Record<string, string | number | boolean>) => void;
	showRuler: boolean;
	autosave: boolean;
	defaultZoom: number;
	reserveReviewSidebar: boolean;
	onDirtyChange: (isDirty: boolean) => void;
	onSave: (buffer: ArrayBuffer) => Promise<void>;
	onDocumentNameChange: (name: string, expectedPath?: string | null) => Promise<void>;
	onWordCountChange: (wordCount: DocumentWordCount) => void;
	onLoadPhase?: (phase: string, data?: Record<string, unknown>) => void;
}

export interface DocxReactViewHandle {
	save: () => Promise<boolean>;
	/** Persist pending/in-flight edits before close prompts; does not cancel autosave. */
	flushPendingSave: () => Promise<boolean>;
	prepareForExternalReload: () => Promise<void>;
	exportBuffer: (options?: ExportDocumentBufferOptions) => Promise<ArrayBuffer | null>;
	exportRenderedPdf: () => Promise<ArrayBuffer | null>;
	getComments: () => Array<{ id: number; parentId?: number | null; author?: string; text?: string }>;
	pasteFromClipboard: (options: PasteClipboardOptions) => Promise<boolean>;
	rewriteClipboardTextWithListMarkers: () => Promise<boolean>;
	openFind: () => void;
	openFindReplace: () => void;
	openImagePicker: () => void;
	openCustomTableDialog: () => void;
	openFontPicker: () => void;
	setMode: (mode: EditorMode) => void;
	setZoom: (zoom: number) => void;
}

export const DocxReactView = forwardRef<DocxReactViewHandle, DocxReactViewProps>(function DocxReactView(
	{ file, buffer, documentKey, editorAdapter, error, isLoading, authorName, resolvedEditorTheme, i18n, pluginI18n, showNotice, showRuler, autosave, defaultZoom, reserveReviewSidebar, onDirtyChange, onSave, onDocumentNameChange, onWordCountChange, onLoadPhase },
	ref,
) {
	const editorRef = useRef<DocxEditorRef>(null);
	const sourceBufferRef = useRef<ArrayBuffer | null | undefined>(buffer);
	const renderedDomContextRef = useRef<RenderedDomContext | null>(null);
	const imageInputRef = useRef<HTMLInputElement>(null);
	const fontInputRef = useRef<HTMLInputElement>(null);
	const editorClassNameRef = useRef(`native-powerpoint-doc-editor-editor-${++editorInstanceCounter}`);
	useEffect(() => {
		debugLog('settings', 'DOCX React colorMode', {
			file: file?.path,
			resolvedEditorTheme,
		});
	}, [file?.path, resolvedEditorTheme]);
	const rulerSyncFrameRef = useRef<number | null>(null);
	const rulerSyncTimeoutRef = useRef<number | null>(null);
	const initialCenterFrameRef = useRef<number | null>(null);
	const initialCenterTimeoutsRef = useRef<number[]>([]);
	const centeredDocumentKeyRef = useRef<string | null>(null);
	const pinchZoomStateRef = useRef<PinchZoomState | null>(null);
	const pinchZoomScrollFrameRef = useRef<number | null>(null);
	const activeTouchPointersRef = useRef<Map<number, PointerPoint>>(new Map());
	const fontSizeHoldRef = useRef<FontSizeHoldState | null>(null);
	const listMarkerSelectionFrameRef = useRef<number | null>(null);
	const listLayoutRelayoutFrameRef = useRef<number | null>(null);
	const listLayoutRelayoutSecondFrameRef = useRef<number | null>(null);
	const commentsSidebarToggleFrameRef = useRef<number | null>(null);
	const paginationLogTimeoutRef = useRef<number | null>(null);
	const lastPaginationLogSignatureRef = useRef<string | null>(null);
	const dirtyTrackingEnabledRef = useRef(false);
	const renameTimeoutRef = useRef<number | null>(null);
	const sessionRef = useRef<DocxSession<DocxSaveContext, ArrayBuffer | null, ArrayBuffer, ArrayBuffer> | null>(null);
	const activeDocumentKeyRef = useRef(documentKey);
	const externalReloadBlockedRef = useRef(false);
	if (activeDocumentKeyRef.current !== documentKey) {
		activeDocumentKeyRef.current = documentKey;
		externalReloadBlockedRef.current = false;
	}
	const saveHostRef = useRef({ file, onSave, showNotice, autosave, documentKey });
	saveHostRef.current = { file, onSave, showNotice, autosave, documentKey };
	const [documentName, setDocumentName] = useState(file?.name ?? '');
	const [editorMode, setEditorMode] = useState<EditorMode>('editing');
	const [saveStatus, setSaveStatus] = useState<SaveStatus>('saved');
	const [findDialogMode, setFindDialogMode] = useState<FindReplaceMode | null>(null);
	const [findSearchText, setFindSearchText] = useState('');
	const [findReplaceText, setFindReplaceText] = useState('');
	const publishWordCount = useCallback((view: EditorView | null | undefined) => {
		if (!view) {
			return;
		}

		const { doc, selection } = view.state;
		onWordCountChange({
		totalWords: countDocumentWords(doc.textBetween(0, doc.content.size, ' ')),
			selectedWords: selection.empty
				? null
				: countDocumentWords(doc.textBetween(selection.from, selection.to, ' ')),
		});
	}, [onWordCountChange]);
	const [findMatchCase, setFindMatchCase] = useState(false);
	const [findWholeWord, setFindWholeWord] = useState(false);
	const [findMatches, setFindMatches] = useState<FindMatch[]>([]);
	const [currentFindIndex, setCurrentFindIndex] = useState(0);
	const [importedFonts, setImportedFonts] = useState<FontOption[]>([]);
	const [customTableDialogOpen, setCustomTableDialogOpen] = useState(false);
	const [commentsSidebarOpen, setCommentsSidebarOpen] = useState(false);
	const filePath = file?.path ?? null;
	const findReplaceLabels = useMemo(
		() => (pluginI18n ? createFindReplaceLabels(pluginI18n) : null),
		[pluginI18n],
	);
	const fontFamilies = useMemo<FontOption[]>(() => [
		...DEFAULT_EDITOR_FONT_FAMILIES,
		...importedFonts,
	], [importedFonts]);
	const fontFamiliesRef = useRef(fontFamilies);
	fontFamiliesRef.current = fontFamilies;
	const preservedTextSelectionRef = useRef<TextSelectionRange | null>(null);
	const fontFamilyDisplaySyncVersionRef = useRef(0);
	useEffect(() => {
		editorAdapter.bindEditor(() => editorRef.current);
		editorAdapter.bindMode((mode) => setEditorMode(mode));
		return () => {
			editorAdapter.bindEditor(() => null);
			editorAdapter.bindMode(() => {});
		};
	}, [editorAdapter]);
	const schedulePaginationDiagnostics = useCallback((trigger: string) => {
		if (paginationLogTimeoutRef.current !== null) {
			window.clearTimeout(paginationLogTimeoutRef.current);
		}

		paginationLogTimeoutRef.current = window.setTimeout(() => {
			paginationLogTimeoutRef.current = null;
			const editor = editorRef.current;
			const editorCore = editor?.getEditorRef();
			const renderedPages = renderedDomContextRef.current?.pagesContainer.querySelectorAll(DOCX_RENDERED_PAGE_SELECTOR).length
				?? activeDocument.querySelectorAll(`.${editorClassNameRef.current} ${DOCX_RENDERED_PAGE_SELECTOR}`).length;
			const sourceDiagnostics = getDocxPaginationSourceDiagnostics(editorCore?.getView()?.state.doc);
			const sourceDocument = editor?.getDocument();
			const documentProperties = sourceDocument?.package?.[DOCX_PACKAGE_DOCUMENT_KEY];
			const sectionProperties = {
				...documentProperties?.sections?.[0]?.properties,
				...documentProperties?.finalSectionProperties,
			};
			const details = {
				file: filePath,
				totalPages: editor?.getTotalPages() ?? 0,
				renderedPages,
				pageHeightTwips: sectionProperties.pageHeight ?? DEFAULT_PAGE_HEIGHT_TWIPS,
				topMarginTwips: sectionProperties.marginTop ?? DEFAULT_MARGIN_TWIPS,
				bottomMarginTwips: sectionProperties.marginBottom ?? DEFAULT_MARGIN_TWIPS,
				...sourceDiagnostics,
			};
			const signature = JSON.stringify(details);
			if (signature === lastPaginationLogSignatureRef.current) {
				return;
			}
			lastPaginationLogSignatureRef.current = signature;
			debugLog('pagination', 'DOCX pagination diagnostics', {
				trigger,
				...details,
			});
		}, 100);
	}, [filePath]);
	const scheduleParagraphLayoutRelayout = useCallback(() => {
		clearParagraphMeasureCache();
		if (listLayoutRelayoutFrameRef.current !== null) {
			window.cancelAnimationFrame(listLayoutRelayoutFrameRef.current);
		}
		if (listLayoutRelayoutSecondFrameRef.current !== null) {
			window.cancelAnimationFrame(listLayoutRelayoutSecondFrameRef.current);
			listLayoutRelayoutSecondFrameRef.current = null;
		}

		listLayoutRelayoutFrameRef.current = window.requestAnimationFrame(() => {
			listLayoutRelayoutFrameRef.current = null;
			editorRef.current?.getEditorRef()?.relayout();

			listLayoutRelayoutSecondFrameRef.current = window.requestAnimationFrame(() => {
				listLayoutRelayoutSecondFrameRef.current = null;
				editorRef.current?.getEditorRef()?.relayout();
			});
		});
	}, []);
	const findHighlightPlugin = useMemo(() => new Plugin<FindHighlightState>({
		key: findHighlightPluginKey,
		state: {
			init: () => ({ matches: [], currentIndex: 0 }),
			apply: (transaction, previous) => {
				const nextState: unknown = transaction.getMeta(findHighlightPluginKey);
				return isFindHighlightState(nextState) ? nextState : previous;
			},
		},
		props: {
			decorations: (state) => {
				const pluginState = findHighlightPluginKey.getState(state);
				if (!pluginState || pluginState.matches.length === 0) {
					return DecorationSet.empty;
				}

				return DecorationSet.create(
					state.doc,
					pluginState.matches.map((match, index) => Decoration.inline(
						match.from,
						match.to,
						{ class: index === pluginState.currentIndex ? 'native-powerpoint-doc-editor-find-current' : 'native-powerpoint-doc-editor-find-match' },
					)),
				);
			},
		},
	}), []);
	const paragraphLayoutRelayoutPlugin = useMemo(
		() => createParagraphLayoutRelayoutPlugin(scheduleParagraphLayoutRelayout),
		[scheduleParagraphLayoutRelayout],
	);
	const externalPlugins = useMemo(
		() => [preserveTypedSpacePlugin, contentShrinkDiagnosticsPlugin, findHighlightPlugin, paragraphLayoutRelayoutPlugin],
		[findHighlightPlugin, paragraphLayoutRelayoutPlugin],
	);
	const pluginSidebarItems = useMemo<NonNullable<ComponentProps<typeof DocxEditor>['pluginSidebarItems']>>(() => {
		if (!reserveReviewSidebar) {
			return [];
		}

		return [{
			id: 'native-powerpoint-doc-editor-review-sidebar-reservation',
			anchorPos: 1,
			estimatedHeight: 1,
			priority: Number.MAX_SAFE_INTEGER,
			render: () => null,
		}];
	}, [reserveReviewSidebar]);

	const commentsSidebarToggleLabel = useMemo(
		() => pluginI18n?.t('docx:comments.toggleSidebar') ?? '',
		[pluginI18n],
	);

	const handleCommentsSidebarOpenChange = useCallback((open: boolean) => {
		// Vendor editor only renders the new-comment input when the controlled sidebar is
		// open. Do not block programmatic opens while the document still has zero
		// saved comments; the title-bar toggle stays disabled separately.
		setCommentsSidebarOpen(open);
	}, []);

	const syncCommentsSidebarToggle = useCallback(() => {
		const commentCount = getTopLevelCommentCount(editorRef.current);

		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		const button = findCommentsSidebarToggleButton(editorRoot, commentsSidebarToggleLabel);
		if (!button) {
			return;
		}

		setCommentsSidebarToggleEnabled(button, commentCount > 0);
	}, [commentsSidebarToggleLabel]);

	const scheduleCommentsSidebarToggleSync = useCallback(() => {
		if (commentsSidebarToggleFrameRef.current !== null) {
			window.cancelAnimationFrame(commentsSidebarToggleFrameRef.current);
		}

		commentsSidebarToggleFrameRef.current = window.requestAnimationFrame(() => {
			commentsSidebarToggleFrameRef.current = null;
			syncCommentsSidebarToggle();
		});
	}, [syncCommentsSidebarToggle]);

	const markCommentsDirty = useCallback((reason: 'change' | 'add' | 'reply' | 'delete' | 'resolve' | 'unresolve') => {
		const blocked = externalReloadBlockedRef.current;
		const wrongDocument = activeDocumentKeyRef.current !== documentKey;
		const dirtyTrackingEnabled = dirtyTrackingEnabledRef.current;
		const editorReady = Boolean(editorRef.current && editorAdapter.getView());
		if (blocked || wrongDocument) {
			debugLog('comments', 'DOCX comment dirty skipped', {
				file: filePath,
				reason,
				blocked,
				wrongDocument,
				documentKey,
				activeDocumentKey: activeDocumentKeyRef.current,
			});
			scheduleCommentsSidebarToggleSync();
			return;
		}
		// Load/hydrate and teardown emit onCommentsChange (often total:0) when the
		// editor cannot serialize. Marking dirty then forces close prompts / failed
		// flush-before-close even after a successful content autosave.
		if (reason === 'change' && (!dirtyTrackingEnabled || !editorReady)) {
			debugLog('comments', 'DOCX comment dirty skipped (hydrate)', {
				file: filePath,
				reason,
				dirtyTrackingEnabled,
				editorReady,
			});
			scheduleCommentsSidebarToggleSync();
			return;
		}
		sessionRef.current?.markDirty();
		debugLog('comments', 'DOCX comment dirty marked', {
			file: filePath,
			reason,
			dirtyTrackingEnabled,
			editorReady,
			editVersion: sessionRef.current?.editVersion ?? null,
		});
		if (!editorReady) {
			debugLog('comments', 'DOCX comment dirty flush deferred (editor not ready)', {
				file: filePath,
				reason,
			});
			scheduleCommentsSidebarToggleSync();
			return;
		}
		void sessionRef.current?.save('autosave').then((ok) => {
			debugLog('comments', 'DOCX comment dirty save settled', {
				file: filePath,
				reason,
				ok,
			});
		}).catch((saveFailure: unknown) => {
			debugLog('comments', 'DOCX comment dirty save failed', {
				file: filePath,
				reason,
				error: saveFailure instanceof Error ? saveFailure.message : String(saveFailure),
			});
		});
		scheduleCommentsSidebarToggleSync();
	}, [documentKey, editorAdapter, filePath, scheduleCommentsSidebarToggleSync]);

	const handleCommentsChange = useCallback((comments: Parameters<NonNullable<ComponentProps<typeof DocxEditor>['onCommentsChange']>>[0]) => {
		debugLog('comments', 'DOCX comments changed', {
			file: filePath,
			...summarizeDocxComments(comments),
		});
		markCommentsDirty('change');
	}, [filePath, markCommentsDirty]);

	const handleCommentAdd = useCallback((comment: Parameters<NonNullable<ComponentProps<typeof DocxEditor>['onCommentAdd']>>[0]) => {
		debugLog('comments', 'DOCX comment added', {
			file: filePath,
			comment: summarizeDocxComment(comment),
		});
		markCommentsDirty('add');
	}, [filePath, markCommentsDirty]);

	const handleCommentReply = useCallback((
		reply: Parameters<NonNullable<ComponentProps<typeof DocxEditor>['onCommentReply']>>[0],
		parent: Parameters<NonNullable<ComponentProps<typeof DocxEditor>['onCommentReply']>>[1],
	) => {
		debugLog('comments', 'DOCX comment reply added', {
			file: filePath,
			reply: summarizeDocxComment(reply),
			parent: summarizeDocxComment(parent),
		});
		markCommentsDirty('reply');
	}, [filePath, markCommentsDirty]);

	const handleCommentDelete = useCallback((comment: Parameters<NonNullable<ComponentProps<typeof DocxEditor>['onCommentDelete']>>[0]) => {
		debugLog('comments', 'DOCX comment deleted', {
			file: filePath,
			comment: summarizeDocxComment(comment),
		});
		markCommentsDirty('delete');
	}, [filePath, markCommentsDirty]);

	const handleCommentResolve = useCallback((comment: Parameters<NonNullable<ComponentProps<typeof DocxEditor>['onCommentResolve']>>[0]) => {
		debugLog('comments', 'DOCX comment resolved', {
			file: filePath,
			comment: summarizeDocxComment(comment),
		});
		markCommentsDirty('resolve');
	}, [filePath, markCommentsDirty]);

	const handleCommentUnresolve = useCallback((comment: Parameters<NonNullable<ComponentProps<typeof DocxEditor>['onCommentUnresolve']>>[0]) => {
		debugLog('comments', 'DOCX comment unresolved', {
			file: filePath,
			comment: summarizeDocxComment(comment),
		});
		markCommentsDirty('unresolve');
	}, [filePath, markCommentsDirty]);

	useEffect(() => {
		ensureEditorStyles();
	}, []);

	useEffect(() => {
		setCommentsSidebarOpen(false);
	}, [documentKey]);

	useEffect(() => {
		lastPaginationLogSignatureRef.current = null;
		return () => {
			if (paginationLogTimeoutRef.current !== null) {
				window.clearTimeout(paginationLogTimeoutRef.current);
				paginationLogTimeoutRef.current = null;
			}
		};
	}, [documentKey]);

	useEffect(() => {
		if (!file || !buffer || isLoading) {
			return;
		}

		let detach: (() => void) | undefined;
		let retryTimeouts: number[] = [];

		const attachNeutralizer = (): boolean => {
			if (detach) {
				return true;
			}
			const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			if (!editorRoot) {
				return false;
			}

			detach = attachDocxImeTransformNeutralizer(editorRoot, {
				getEditorView: () => editorRef.current?.getEditorRef()?.getView() ?? null,
				getRenderedDomContext: () => renderedDomContextRef.current,
				onDiagnostic: ({ event, details }) => {
					debugLog('ime', `DOCX IME ${event}`, {
						file: filePath,
						...details,
					});
				},
			});
			return true;
		};

		if (!attachNeutralizer()) {
			retryTimeouts = [100, 500, 1500].map((delay) => window.setTimeout(attachNeutralizer, delay));
		}

		return () => {
			for (const timeout of retryTimeouts) {
				window.clearTimeout(timeout);
			}
			retryTimeouts = [];
			detach?.();
		};
	}, [buffer, documentKey, filePath, isLoading]);

	useEffect(() => {
		// Reset baseline when the open document session changes.
		sourceBufferRef.current = buffer ?? null;
	}, [documentKey, filePath]);

	useEffect(() => {
		// Async first package for this session (null → buffer) without remount.
		// Ignore later host buffer identity churn from self-save.
		if (sourceBufferRef.current == null && buffer) {
			sourceBufferRef.current = buffer;
		}
	}, [buffer]);

	const syncListMarkerSelectionHighlights = useCallback(() => {
		const root = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		const view = editorRef.current?.getEditorRef()?.getView();
		if (!root || !view) {
			return;
		}

		updateListMarkerSelectionHighlights(root, view);
	}, []);

	const scheduleListMarkerSelectionHighlightSync = useCallback(() => {
		if (listMarkerSelectionFrameRef.current !== null) {
			window.cancelAnimationFrame(listMarkerSelectionFrameRef.current);
		}

		listMarkerSelectionFrameRef.current = window.requestAnimationFrame(() => {
			listMarkerSelectionFrameRef.current = null;
			syncListMarkerSelectionHighlights();
		});
	}, [syncListMarkerSelectionHighlights]);

	useEffect(() => () => {
		if (listMarkerSelectionFrameRef.current !== null) {
			window.cancelAnimationFrame(listMarkerSelectionFrameRef.current);
			listMarkerSelectionFrameRef.current = null;
		}
		if (listLayoutRelayoutFrameRef.current !== null) {
			window.cancelAnimationFrame(listLayoutRelayoutFrameRef.current);
			listLayoutRelayoutFrameRef.current = null;
		}
		if (listLayoutRelayoutSecondFrameRef.current !== null) {
			window.cancelAnimationFrame(listLayoutRelayoutSecondFrameRef.current);
			listLayoutRelayoutSecondFrameRef.current = null;
		}
	}, []);

	useEffect(() => {
		dirtyTrackingEnabledRef.current = false;
		const timeout = window.setTimeout(() => {
			dirtyTrackingEnabledRef.current = true;
		}, 500);

		return () => {
			window.clearTimeout(timeout);
			dirtyTrackingEnabledRef.current = false;
		};
		// Gate on document session, not host buffer identity. Post-autosave
		// `this.buffer = output` used to disable dirty tracking for 500ms and
		// drop paste/typing from the save pipeline.
	}, [file, documentKey]);

	useEffect(() => {
		setDocumentName(file?.name ?? '');
		setEditorMode('editing');
		setSaveStatus('saved');
		setFindDialogMode(null);
		setFindSearchText('');
		setFindReplaceText('');
		setFindMatches([]);
		setCurrentFindIndex(0);
	}, [documentKey, filePath]);

	const setMode = useCallback((mode: EditorMode) => {
		debugLog('editor', 'DOCX editor mode changed', { file: filePath, mode });
		editorAdapter.setMode(mode);
	}, [editorAdapter, filePath]);

	const publishFindHighlights = useCallback((matches: FindMatch[], currentIndex: number) => {
		const view = editorAdapter.getView();
		if (!view) {
			return;
		}

		view.dispatch(view.state.tr.setMeta(findHighlightPluginKey, { matches, currentIndex }));
	}, [editorAdapter]);

	const selectFindMatch = useCallback((matches: FindMatch[], index: number) => {
		const match = matches[index];
		if (!match) {
			return;
		}

		editorAdapter.select(match);
	}, [editorAdapter]);

	const refreshFindMatches = useCallback((searchText: string, matchCase = findMatchCase, wholeWord = findWholeWord, preferredIndex = 0, options: RefreshFindOptions = {}) => {
		const matches = editorAdapter.find(searchText, { matchCase, wholeWord });
		const nextIndex = matches.length > 0 ? Math.max(0, Math.min(preferredIndex, matches.length - 1)) : 0;

		setFindMatches(matches);
		setCurrentFindIndex(nextIndex);
		publishFindHighlights(matches, nextIndex);
		if (options.select && matches.length > 0) {
			selectFindMatch(matches, nextIndex);
		}

		debugLog('search', 'DOCX find results refreshed', {
			file: filePath,
			queryLength: searchText.length,
			matchCase,
			wholeWord,
			matchCount: matches.length,
			selectedMatch: matches.length > 0 ? nextIndex + 1 : 0,
		});
		return matches;
	}, [editorAdapter, filePath, findMatchCase, findWholeWord, publishFindHighlights, selectFindMatch]);

	const openFindReplacePanel = useCallback((mode: FindReplaceMode) => {
		const selectedText = editorAdapter.getSelectedText();
		const nextSearchText = selectedText || findSearchText;

		debugLog('search', 'Opened DOCX find panel', {
			file: filePath,
			mode,
			seededFromSelection: Boolean(selectedText),
			queryLength: nextSearchText.length,
		});
		setFindDialogMode(mode);
		if (selectedText) {
			setFindSearchText(selectedText);
		}
		refreshFindMatches(nextSearchText);
	}, [editorAdapter, filePath, findSearchText, refreshFindMatches]);

	const openFindReplaceDialog = useCallback((mode: FindReplaceMode) => {
		openFindReplacePanel(mode);
	}, [openFindReplacePanel]);

	useEffect(() => {
		const handleFindShortcut = (evt: KeyboardEvent) => {
			const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			const key = evt.key.toLowerCase();
			const isFindShortcut = isPrimaryShortcut(evt, 'f');
			const isMacReplaceShortcut = Platform.isMacOS && key === 'f' && evt.metaKey && !evt.ctrlKey && evt.altKey && !evt.shiftKey;
			const isWinReplaceShortcut = !Platform.isMacOS && key === 'h' && evt.ctrlKey && !evt.metaKey && !evt.altKey && !evt.shiftKey;
			const isReplaceShortcut = isMacReplaceShortcut || isWinReplaceShortcut;
			if (
				!editorRoot
				|| (!isFindShortcut && !isReplaceShortcut)
				|| !isNode(evt.target)
				|| (!editorRoot.contains(evt.target) && !activeDocument.querySelector('.native-powerpoint-doc-editor-find-dialog')?.contains(evt.target))
			) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			openFindReplaceDialog(isFindShortcut ? 'find' : 'replace');
		};

		activeDocument.addEventListener('keydown', handleFindShortcut, true);
		return () => activeDocument.removeEventListener('keydown', handleFindShortcut, true);
	}, [openFindReplaceDialog]);

	const moveFindMatch = useCallback((direction: 1 | -1) => {
		if (findMatches.length === 0) {
			return;
		}

		const nextIndex = wrapMatchIndex(currentFindIndex, direction, findMatches.length);
		setCurrentFindIndex(nextIndex);
		publishFindHighlights(findMatches, nextIndex);
		selectFindMatch(findMatches, nextIndex);
	}, [currentFindIndex, findMatches, publishFindHighlights, selectFindMatch]);

	const replaceCurrentMatch = useCallback(() => {
		const match = findMatches[currentFindIndex];
		if (!match || !editorAdapter.replace(match, findReplaceText)) {
			return;
		}

		debugLog('search', 'Replaced current DOCX find match', {
			file: filePath,
			replacementLength: findReplaceText.length,
			matchCountBeforeReplace: findMatches.length,
		});
		refreshFindMatches(findSearchText, findMatchCase, findWholeWord, currentFindIndex);
	}, [currentFindIndex, editorAdapter, filePath, findMatchCase, findMatches, findReplaceText, findSearchText, findWholeWord, refreshFindMatches]);

	const replaceAllMatches = useCallback(() => {
		if (!editorAdapter.replaceAll(findMatches, findReplaceText)) {
			return;
		}

		debugLog('search', 'Replaced all DOCX find matches', {
			file: filePath,
			replacedCount: findMatches.length,
			replacementLength: findReplaceText.length,
		});
		refreshFindMatches(findSearchText, findMatchCase, findWholeWord, 0);
	}, [editorAdapter, filePath, findMatchCase, findMatches, findReplaceText, findSearchText, findWholeWord, refreshFindMatches]);

	const normalizeEditorModeDropdown = useCallback(() => {
		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		normalizeEditorFloatingLayers(editorRoot);
		const modeMenus = Array.from(activeDocument.querySelectorAll<HTMLElement>('div[style*="position: fixed"]'))
			.map((menu) => ({
				menu,
				buttons: Array.from(menu.querySelectorAll<HTMLButtonElement>(':scope > button'))
					.filter((button) => getEditorModeFromButton(button) !== null && button.querySelector(':scope span span')),
			}))
			.filter(({ buttons }) => {
				const modes = new Set(buttons.map((button) => getEditorModeFromButton(button)));
				return buttons.length === 3 && modes.has('editing') && modes.has('suggesting') && modes.has('viewing');
			});

		modeMenus.forEach(({ menu, buttons }) => {
			menu.dataset.nativePowerPointDocEditorModeMenu = 'true';
			markLightMenuSurface(menu, 'native-powerpoint-doc-editor-mode-menu');
			menu.addClass('native-powerpoint-doc-editor-mode-menu-normalized');

			buttons.forEach((button) => {
				const mode = getEditorModeFromButton(button);
				if (mode) {
					button.dataset.nativePowerPointDocEditorModeMenuItem = mode;
				}

				button.addClass('native-powerpoint-doc-editor-mode-menu-item');

				const icon = button.querySelector<HTMLElement>(':scope > svg:first-child');
				if (icon) {
					icon.addClass('native-powerpoint-doc-editor-mode-menu-icon');
				}

				const labelColumn = button.querySelector<HTMLElement>(':scope > span');
				if (labelColumn) {
					labelColumn.addClass('native-powerpoint-doc-editor-mode-menu-label');
				}

				const checkIcon = button.querySelector<HTMLElement>(':scope > svg:last-child:not(:first-child)');
				if (checkIcon) {
					checkIcon.addClass('native-powerpoint-doc-editor-mode-menu-check');
				}
			});
		});
	}, []);

	useEffect(() => {
		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		normalizeEditorModeDropdown();
		const observer = new MutationObserver(normalizeEditorModeDropdown);
		observer.observe(activeDocument.body, {
			childList: true,
			subtree: true,
		});

		return () => observer.disconnect();
	}, [buffer, filePath, isLoading, normalizeEditorModeDropdown]);

	useEffect(() => {
		scheduleCommentsSidebarToggleSync();

		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		const observer = new MutationObserver(() => {
			scheduleCommentsSidebarToggleSync();
		});
		observer.observe(editorRoot, {
			childList: true,
			subtree: true,
			attributes: true,
			attributeFilter: ['aria-pressed', 'disabled', 'class', 'aria-label'],
		});

		return () => observer.disconnect();
	}, [buffer, filePath, isLoading, scheduleCommentsSidebarToggleSync]);

	useEffect(() => {
		const suppressEvent = (evt: Event) => {
			evt.preventDefault();
			evt.stopImmediatePropagation();
			evt.stopPropagation();
		};

		const handleBlockedToggle = (evt: Event) => {
			if (!evt.isTrusted) {
				return;
			}

			if ('button' in evt && evt.button !== 0) {
				return;
			}

			if (!isElement(evt.target)) {
				return;
			}

			const button = evt.target.closest<HTMLButtonElement>(`[${COMMENTS_SIDEBAR_TOGGLE_ATTR}]`);
			if (!button) {
				return;
			}

			const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			if (!editorRoot?.contains(button)) {
				return;
			}

			if (getTopLevelCommentCount(editorRef.current) > 0) {
				return;
			}

			suppressEvent(evt);
		};

		activeDocument.addEventListener('pointerdown', handleBlockedToggle, true);
		activeDocument.addEventListener('click', handleBlockedToggle, true);

		return () => {
			activeDocument.removeEventListener('pointerdown', handleBlockedToggle, true);
			activeDocument.removeEventListener('click', handleBlockedToggle, true);
		};
	}, [buffer, filePath, isLoading]);

	useEffect(() => {
		const handleModePointerDown = (evt: PointerEvent) => {
			const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			if (!editorRoot || !isElement(evt.target)) {
				return;
			}

			const button = evt.target.closest('button');
			if (!isHTMLButtonElement(button)) {
				return;
			}

			const isEditorModeButton = editorRoot.contains(button) || button.closest('[data-native-powerpoint-doc-editor-mode-menu]');
			if (!isEditorModeButton) {
				return;
			}
			const mode = getEditorModeFromButton(button);
			if (mode) {
				window.setTimeout(() => setMode(mode), 0);
			}
		};

		activeDocument.addEventListener('pointerdown', handleModePointerDown, true);
		return () => activeDocument.removeEventListener('pointerdown', handleModePointerDown, true);
	}, [setMode]);

	const clearAutosaveTimeout = useCallback(() => {
		sessionRef.current?.clearAutosave();
	}, []);

	const clearRenameTimeout = useCallback(() => {
		if (renameTimeoutRef.current !== null) {
			window.clearTimeout(renameTimeoutRef.current);
			renameTimeoutRef.current = null;
		}
	}, []);

	useEffect(() => {
		clearRenameTimeout();
	}, [clearRenameTimeout, filePath]);

	const syncVerticalRulerMarkers = useCallback((docxDocument: DocxDocumentWithSectionProperties | null | undefined) => {
		if (!showRuler || !docxDocument) {
			return;
		}

		const documentProperties = docxDocument.package?.[DOCX_PACKAGE_DOCUMENT_KEY];
		const sectionProperties = {
			...documentProperties?.sections?.[0]?.properties,
			...documentProperties?.finalSectionProperties,
		};
		const pageHeight = sectionProperties.pageHeight ?? DEFAULT_PAGE_HEIGHT_TWIPS;
		const topMargin = sectionProperties.marginTop ?? DEFAULT_MARGIN_TWIPS;
		const bottomMargin = sectionProperties.marginBottom ?? DEFAULT_MARGIN_TWIPS;
		const ruler = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current} .docx-vertical-ruler`);

		if (!ruler || pageHeight <= 0) {
			return;
		}

		const pxPerTwip = ruler.getBoundingClientRect().height / pageHeight;
		const topMarker = ruler.querySelector<HTMLElement>('.docx-ruler-marker-topMargin');
		const bottomMarker = ruler.querySelector<HTMLElement>('.docx-ruler-marker-bottomMargin');

		if (topMarker) {
		topMarker.setCssProps({ top: `${Math.round(topMargin * pxPerTwip - 5)}px` });
		}
		if (bottomMarker) {
		bottomMarker.setCssProps({ top: `${Math.round((pageHeight - bottomMargin) * pxPerTwip - 5)}px` });
		}
	}, [showRuler]);

	const scheduleVerticalRulerMarkerSync = useCallback((sourceDocument: DocxDocumentWithSectionProperties | null | undefined) => {
		if (rulerSyncFrameRef.current !== null) {
			window.cancelAnimationFrame(rulerSyncFrameRef.current);
		}
		if (rulerSyncTimeoutRef.current !== null) {
			window.clearTimeout(rulerSyncTimeoutRef.current);
		}

		rulerSyncFrameRef.current = window.requestAnimationFrame(() => {
			rulerSyncFrameRef.current = null;
			syncVerticalRulerMarkers(sourceDocument);
			window.requestAnimationFrame(() => syncVerticalRulerMarkers(sourceDocument));
		});
		rulerSyncTimeoutRef.current = window.setTimeout(() => {
			rulerSyncTimeoutRef.current = null;
			syncVerticalRulerMarkers(sourceDocument);
		}, 50);
	}, [syncVerticalRulerMarkers]);

	const clearInitialDocumentCenter = useCallback(() => {
		if (initialCenterFrameRef.current !== null) {
			window.cancelAnimationFrame(initialCenterFrameRef.current);
			initialCenterFrameRef.current = null;
		}

		for (const timeout of initialCenterTimeoutsRef.current) {
			window.clearTimeout(timeout);
		}
		initialCenterTimeoutsRef.current = [];
	}, []);

	const centerInitialDocumentViewport = useCallback(() => {
		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return false;
		}

		return centerEditorViewport(editorRoot);
	}, []);

	const scheduleInitialDocumentCenter = useCallback(() => {
		clearInitialDocumentCenter();

		const runCenter = () => {
			centerInitialDocumentViewport();
		};

		initialCenterFrameRef.current = window.requestAnimationFrame(() => {
			initialCenterFrameRef.current = null;
			runCenter();
			window.requestAnimationFrame(runCenter);
		});
		initialCenterTimeoutsRef.current = [80, 240, 600].map((delay) => window.setTimeout(runCenter, delay));
	}, [centerInitialDocumentViewport, clearInitialDocumentCenter]);

	useEffect(() => {
		if (showRuler) {
			scheduleVerticalRulerMarkerSync(editorRef.current?.getDocument());
		}
	}, [showRuler, file, buffer, scheduleVerticalRulerMarkerSync]);

	useEffect(() => {
		if (isLoading || error || !file || !buffer || centeredDocumentKeyRef.current === documentKey) {
			return;
		}

		centeredDocumentKeyRef.current = documentKey;
		scheduleInitialDocumentCenter();
		return clearInitialDocumentCenter;
	}, [buffer, clearInitialDocumentCenter, documentKey, error, file, isLoading, scheduleInitialDocumentCenter]);

	useEffect(() => {
		if (!shouldEnableTouchPinchZoom()) {
			return;
		}

		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		const hostRoot = editorRoot.closest<HTMLElement>('.native-powerpoint-doc-editor-host') ?? editorRoot;
		editorRoot.addClass('native-powerpoint-doc-editor-touch-pinch-root');
		hostRoot.addClass('native-powerpoint-doc-editor-touch-pinch-root');

		const isEditorTarget = (target: EventTarget | null) => isNode(target) && hostRoot.contains(target);

		const shouldIgnoreGestureSource = (source: PinchZoomState['source']) => {
			const activeSource = pinchZoomStateRef.current?.source;
			return activeSource !== undefined && activeSource !== source;
		};

		const zoomAroundViewportPoint = (nextZoom: number, viewportPoint: PointerPoint, source: PinchZoomState['source']) => {
			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || pinchState.source !== source || Math.abs(nextZoom - pinchState.lastZoom) < TOUCH_ZOOM_MIN_DELTA) {
				return false;
			}

			const scrollContainer = getScrollableEditorElement(editorRoot);
			const rect = scrollContainer.getBoundingClientRect();
			const localX = viewportPoint.x - rect.left;
			const localY = viewportPoint.y - rect.top;
			const documentX = (scrollContainer.scrollLeft + localX) / pinchState.lastZoom;
			const documentY = (scrollContainer.scrollTop + localY) / pinchState.lastZoom;

			editorRef.current?.setZoom(nextZoom);
			pinchState.lastZoom = nextZoom;

			if (pinchZoomScrollFrameRef.current !== null) {
				window.cancelAnimationFrame(pinchZoomScrollFrameRef.current);
			}
			pinchZoomScrollFrameRef.current = window.requestAnimationFrame(() => {
				pinchZoomScrollFrameRef.current = null;
				scrollContainer.scrollLeft = Math.max(0, documentX * nextZoom - localX);
				scrollContainer.scrollTop = Math.max(0, documentY * nextZoom - localY);
				scheduleVerticalRulerMarkerSync(editorRef.current?.getDocument());
			});
			return true;
		};

		const handleTouchStart = (evt: TouchEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}
			if (shouldIgnoreGestureSource('touch')) {
				return;
			}
			if (evt.touches.length !== 2) {
				if (pinchZoomStateRef.current?.source === 'touch') {
					pinchZoomStateRef.current = null;
				}
				return;
			}

			const first = evt.touches.item(0);
			const second = evt.touches.item(1);
			if (!first || !second) {
				return;
			}
			const startZoom = editorRef.current?.getZoom() ?? 1;
			const startDistance = getTouchDistance(first, second);
			if (startDistance <= 0) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			pinchZoomStateRef.current = {
				source: 'touch',
				startDistance,
				lastDistance: startDistance,
				startZoom,
				lastZoom: startZoom,
			};
		};

		const handleTouchMove = (evt: TouchEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}
			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || pinchState.source !== 'touch' || evt.touches.length !== 2) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();

			const first = evt.touches.item(0);
			const second = evt.touches.item(1);
			if (!first || !second) {
				return;
			}
			const distance = getTouchDistance(first, second);
			if (distance <= 0) {
				return;
			}

			const center = getTouchCenter(first, second);
			const didZoom = zoomAroundViewportPoint(scaleTouchZoom(pinchState.lastZoom, distance / pinchState.lastDistance), center, 'touch');
			if (didZoom) {
				pinchState.lastDistance = distance;
			}
		};

		const handleTouchEnd = (evt: TouchEvent) => {
			if (evt.touches.length < 2 && pinchZoomStateRef.current?.source === 'touch') {
				pinchZoomStateRef.current = null;
			}
		};

		const handleGestureEnd = () => {
			if (pinchZoomStateRef.current?.source === 'gesture') {
				pinchZoomStateRef.current = null;
			}
		};

		const handleGestureStart = (evt: WebKitGestureEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}
			if (shouldIgnoreGestureSource('gesture')) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			const startZoom = editorRef.current?.getZoom() ?? 1;
			pinchZoomStateRef.current = {
				source: 'gesture',
				startDistance: 1,
				lastDistance: 1,
				startZoom,
				lastZoom: startZoom,
			};
		};

		const handleGestureChange = (evt: WebKitGestureEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}

			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || typeof evt.scale !== 'number' || evt.scale <= 0) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();

			const scrollContainer = getScrollableEditorElement(editorRoot);
			const rect = scrollContainer.getBoundingClientRect();
			zoomAroundViewportPoint(scaleTouchZoom(pinchState.startZoom, evt.scale), {
				x: evt.clientX ?? rect.left + rect.width / 2,
				y: evt.clientY ?? rect.top + rect.height / 2,
			}, 'gesture');
		};

		const handlePointerDown = (evt: PointerEvent) => {
			if (evt.pointerType !== 'touch' || !isEditorTarget(evt.target)) {
				return;
			}
			if (shouldIgnoreGestureSource('pointer')) {
				return;
			}

			activeTouchPointersRef.current.set(evt.pointerId, { x: evt.clientX, y: evt.clientY });
			if (activeTouchPointersRef.current.size !== 2) {
				return;
			}

			const [first, second] = Array.from(activeTouchPointersRef.current.values());
			if (!first || !second) {
				return;
			}
			const startDistance = getPointDistance(first, second);
			if (startDistance <= 0) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			const startZoom = editorRef.current?.getZoom() ?? 1;
			pinchZoomStateRef.current = {
				source: 'pointer',
				startDistance,
				lastDistance: startDistance,
				startZoom,
				lastZoom: startZoom,
			};
		};

		const handlePointerMove = (evt: PointerEvent) => {
			if (evt.pointerType !== 'touch' || !activeTouchPointersRef.current.has(evt.pointerId)) {
				return;
			}

			activeTouchPointersRef.current.set(evt.pointerId, { x: evt.clientX, y: evt.clientY });
			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || pinchState.source !== 'pointer' || activeTouchPointersRef.current.size !== 2) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			const [first, second] = Array.from(activeTouchPointersRef.current.values());
			if (!first || !second) {
				return;
			}
			const distance = getPointDistance(first, second);
			if (distance <= 0) {
				return;
			}

			const didZoom = zoomAroundViewportPoint(
				scaleTouchZoom(pinchState.lastZoom, distance / pinchState.lastDistance),
				getPointCenter(first, second),
				'pointer',
			);
			if (didZoom) {
				pinchState.lastDistance = distance;
			}
		};

		const handlePointerEnd = (evt: PointerEvent) => {
			if (evt.pointerType !== 'touch') {
				return;
			}

			activeTouchPointersRef.current.delete(evt.pointerId);
			if (activeTouchPointersRef.current.size < 2 && pinchZoomStateRef.current?.source === 'pointer') {
				pinchZoomStateRef.current = null;
			}
		};

		activeDocument.addEventListener('touchstart', handleTouchStart, { passive: false, capture: true });
		activeDocument.addEventListener('touchmove', handleTouchMove, { passive: false, capture: true });
		activeDocument.addEventListener('touchend', handleTouchEnd, { passive: true, capture: true });
		activeDocument.addEventListener('touchcancel', handleTouchEnd, { passive: true, capture: true });
		activeDocument.addEventListener('gesturestart', handleGestureStart, { passive: false, capture: true });
		activeDocument.addEventListener('gesturechange', handleGestureChange, { passive: false, capture: true });
		activeDocument.addEventListener('gestureend', handleGestureEnd, { passive: true, capture: true });
		activeDocument.addEventListener('pointerdown', handlePointerDown, { passive: false, capture: true });
		activeDocument.addEventListener('pointermove', handlePointerMove, { passive: false, capture: true });
		activeDocument.addEventListener('pointerup', handlePointerEnd, { passive: true, capture: true });
		activeDocument.addEventListener('pointercancel', handlePointerEnd, { passive: true, capture: true });

		return () => {
			editorRoot.removeClass('native-powerpoint-doc-editor-touch-pinch-root');
			hostRoot.removeClass('native-powerpoint-doc-editor-touch-pinch-root');
			activeDocument.removeEventListener('touchstart', handleTouchStart, true);
			activeDocument.removeEventListener('touchmove', handleTouchMove, true);
			activeDocument.removeEventListener('touchend', handleTouchEnd, true);
			activeDocument.removeEventListener('touchcancel', handleTouchEnd, true);
			activeDocument.removeEventListener('gesturestart', handleGestureStart, true);
			activeDocument.removeEventListener('gesturechange', handleGestureChange, true);
			activeDocument.removeEventListener('gestureend', handleGestureEnd, true);
			activeDocument.removeEventListener('pointerdown', handlePointerDown, true);
			activeDocument.removeEventListener('pointermove', handlePointerMove, true);
			activeDocument.removeEventListener('pointerup', handlePointerEnd, true);
			activeDocument.removeEventListener('pointercancel', handlePointerEnd, true);
			pinchZoomStateRef.current = null;
			activeTouchPointersRef.current.clear();
			if (pinchZoomScrollFrameRef.current !== null) {
				window.cancelAnimationFrame(pinchZoomScrollFrameRef.current);
				pinchZoomScrollFrameRef.current = null;
			}
		};
	}, [buffer, filePath, isLoading, scheduleVerticalRulerMarkerSync]);

	useEffect(() => () => {
		clearAutosaveTimeout();
		clearInitialDocumentCenter();
		clearRenameTimeout();
		if (rulerSyncFrameRef.current !== null) {
			window.cancelAnimationFrame(rulerSyncFrameRef.current);
			rulerSyncFrameRef.current = null;
		}
		if (rulerSyncTimeoutRef.current !== null) {
			window.clearTimeout(rulerSyncTimeoutRef.current);
			rulerSyncTimeoutRef.current = null;
		}
		if (commentsSidebarToggleFrameRef.current !== null) {
			window.cancelAnimationFrame(commentsSidebarToggleFrameRef.current);
			commentsSidebarToggleFrameRef.current = null;
		}
	}, [clearAutosaveTimeout, clearInitialDocumentCenter, clearRenameTimeout]);

	const prepareDocumentBufferForWrite = useCallback(async (output: ArrayBuffer, reason: 'save' | 'export') => {
		const startedAt = performance.now();
		try {
			const preserved = await preserveDocxTableCellFontSizes(sourceBufferRef.current, output);
			debugLog('font-preservation', 'DOCX table-cell font-size preservation check completed', {
				file: filePath,
				reason,
				status: preserved.status,
				sourceBytes: sourceBufferRef.current?.byteLength ?? 0,
				outputBytes: output.byteLength,
				preparedBytes: preserved.buffer.byteLength,
				sourceCellCount: preserved.sourceCellCount,
				outputCellCount: preserved.outputCellCount,
				matchedCellCount: preserved.matchedCellCount,
				skippedTextChangedCells: preserved.skippedTextChangedCells,
				skippedRunCountChangedCells: preserved.skippedRunCountChangedCells,
				sourceRunsWithDirectSize: preserved.sourceRunsWithDirectSize,
				restoredRuns: preserved.restoredRuns,
				restoredTags: preserved.restoredTags,
				ms: Math.round(performance.now() - startedAt),
			});
			return preserved.buffer;
		} catch (preserveError) {
			warnLog('font-preservation', 'DOCX table-cell font-size preservation check failed', {
				file: filePath,
				reason,
				outputBytes: output.byteLength,
				error: preserveError,
			});
			return output;
		}
	}, [filePath]);
	const prepareDocumentBufferRef = useRef(prepareDocumentBufferForWrite);
	prepareDocumentBufferRef.current = prepareDocumentBufferForWrite;

	const session = useMemo(() => {
		const assertCurrentSessionCanWrite = () => {
			if (externalReloadBlockedRef.current || activeDocumentKeyRef.current !== documentKey) {
				throw new Error('DOCX editor save was superseded by an external reload.');
			}
		};

		return new DocxSession<DocxSaveContext, ArrayBuffer | null, ArrayBuffer, ArrayBuffer>({
			adapter: {
				serialize: async () => {
					assertCurrentSessionCanWrite();
					const output = await editorAdapter.serialize();
					assertCurrentSessionCanWrite();
					return output;
				},
				prepareForWrite: async (output) => {
					assertCurrentSessionCanWrite();
					if (!output) throw new Error('DOCX editor serialization returned no document');
					const prepared = await prepareDocumentBufferRef.current(output, 'save');
					assertCurrentSessionCanWrite();
					return prepared;
				},
				validate: async (output) => {
					assertCurrentSessionCanWrite();
					return output;
				},
				persist: async (output, _validated, context, request) => {
					assertCurrentSessionCanWrite();
					const startedAt = performance.now();
					await context.persist(output);
					assertCurrentSessionCanWrite();
					sourceBufferRef.current = output;
					if (request.source === 'manual') {
						saveHostRef.current.showNotice('docx:notice.saved', { fileName: context.file.name });
					}
					const commentsSummary = summarizeDocxComments(editorRef.current?.getComments() ?? []);
					debugLog('save', 'DOCX vault write completed', {
						file: context.file.path,
						bytes: output.byteLength,
						source: request.source,
						targetVersion: request.targetVersion,
						ms: Math.round(performance.now() - startedAt),
						comments: commentsSummary,
					});
				},
			},
			getContext: () => {
				const { file: currentFile, onSave: persist, documentKey: hostDocumentKey } = saveHostRef.current;
				if (
					externalReloadBlockedRef.current
					|| activeDocumentKeyRef.current !== documentKey
					|| hostDocumentKey !== documentKey
				) {
					return null;
				}
				return currentFile ? { file: currentFile, persist } : null;
			},
			autosave: {
				enabled: () => (
					!externalReloadBlockedRef.current
					&& activeDocumentKeyRef.current === documentKey
					&& saveHostRef.current.documentKey === documentKey
					&& saveHostRef.current.autosave
				),
				delayMs: () => 1500,
				source: 'autosave',
			},
			onAutosaveScheduled: (delayMs, version) => debugLog('save', 'DOCX autosave scheduled', {
				file: saveHostRef.current.file?.path ?? null, delayMs, dirtyVersion: version,
			}),
			onAutosaveStarted: (version) => debugLog('save', 'DOCX autosave started', {
				file: saveHostRef.current.file?.path ?? null, dirtyVersion: version,
			}),
		});
	}, [documentKey, editorAdapter]);
	sessionRef.current = session;

	useEffect(() => session.subscribe(({ dirty, saveState, saveError }) => {
		if (externalReloadBlockedRef.current || activeDocumentKeyRef.current !== documentKey) {
			return;
		}
		onDirtyChange(dirty);
		setSaveStatus(saveState === 'clean' ? 'saved' : saveState === 'dirty' ? 'unsaved' : saveState);
		if (saveState !== 'failed') return;
		const currentFile = saveHostRef.current.file;
		const message = saveError instanceof Error ? saveError.message : 'Unknown save error';
		errorLog('save', 'DOCX save failed', { file: currentFile?.path ?? null, error: saveError });
		saveHostRef.current.showNotice(
			message === 'DOCX editor serialization returned no document' ? 'errors:saveNoDocument' : 'errors:saveFailed',
			message === 'DOCX editor serialization returned no document'
				? { fileName: currentFile?.name ?? 'document' }
				: { fileName: currentFile?.name ?? 'document', message },
		);
	}), [documentKey, onDirtyChange, session]);

	useEffect(() => () => {
		session.clearAutosave();
		if (activeDocumentKeyRef.current === documentKey) {
			externalReloadBlockedRef.current = true;
		}
	}, [documentKey, session]);

	const saveDocument = useCallback(async (source: DocxSaveSource = 'manual') => {
		if (externalReloadBlockedRef.current || activeDocumentKeyRef.current !== documentKey) {
			debugLog('save', 'Ignored stale DOCX save request during external reload', { file: file?.path ?? null, source });
			return false;
		}
		if (!file) {
			showNotice('docx:notice.noFileOpen');
			return false;
		}

		debugLog('save', 'DOCX save requested', {
			file: file.path,
			source,
			dirtyVersion: session.editVersion,
		});
		return session.save(source);
	}, [documentKey, file, session, showNotice]);

	const exportDocumentBuffer = useCallback(async (options?: ExportDocumentBufferOptions) => {
		if (!options?.preserveAutosave) {
			session.clearAutosave();
		}
		if (externalReloadBlockedRef.current || activeDocumentKeyRef.current !== documentKey) {
			debugLog('export', 'Ignored stale DOCX export during external reload', { file: file?.path ?? null });
			return null;
		}

		if (!file) {
			showNotice('docx:notice.noFileOpen');
			return null;
		}

		const output = await editorAdapter.serialize();
		if (externalReloadBlockedRef.current || activeDocumentKeyRef.current !== documentKey) {
			debugLog('export', 'Discarded stale DOCX export after external reload', { file: file.path });
			return null;
		}
		if (!output) {
			showNotice('errors:exportNoDocument', { path: file.name });
			return null;
		}

		const prepared = await prepareDocumentBufferForWrite(output, 'export');
		if (externalReloadBlockedRef.current || activeDocumentKeyRef.current !== documentKey) {
			debugLog('export', 'Discarded prepared stale DOCX export after external reload', { file: file.path });
			return null;
		}
		return prepared;
	}, [documentKey, editorAdapter, file, prepareDocumentBufferForWrite, session, showNotice]);

	const exportRenderedPdfBuffer = useCallback(async () => {
		if (externalReloadBlockedRef.current || activeDocumentKeyRef.current !== documentKey) {
			debugLog('export', 'Ignored stale DOCX PDF export during external reload', { file: filePath });
			return null;
		}
		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			warnLog('export', 'Could not export rendered PDF because the editor root is missing', {
				editorClassName: editorClassNameRef.current,
				hasRenderedDomContext: Boolean(renderedDomContextRef.current?.pagesContainer?.isConnected),
			});
			showNotice('docx:notice.editorNotReady');
			return null;
		}

		try {
			const renderedPagesContainer = renderedDomContextRef.current?.pagesContainer ?? null;
			const pdfBuffer = await exportRenderedPagesToPdf(editorRoot, renderedPagesContainer);
			if (externalReloadBlockedRef.current || activeDocumentKeyRef.current !== documentKey) {
				debugLog('export', 'Discarded stale DOCX PDF export after external reload', { file: filePath });
				return null;
			}
			if (!pdfBuffer) {
				showNotice('errors:exportNoPages');
				return null;
			}
			debugLog('export', 'Exported rendered DOCX pages to PDF', {
				bytes: pdfBuffer.byteLength,
			});
			return pdfBuffer;
		} catch (renderError) {
			errorLog('export', 'Could not export rendered DOCX pages to PDF', renderError);
			const message = renderError instanceof Error ? renderError.message : 'Unknown PDF render error';
			showNotice('errors:exportPdfRenderFailed', { message });
			return null;
		}
	}, [documentKey, filePath, showNotice]);

	const handleRenderedDomContextReady = useCallback((context: RenderedDomContext) => {
		renderedDomContextRef.current = context;
		onLoadPhase?.('editor-rendered-dom-ready', {
			pageCount: context.pagesContainer.querySelectorAll(DOCX_RENDERED_PAGE_SELECTOR).length,
			zoom: context.zoom,
		});
		debugLog('export', 'Rendered DOCX DOM context ready', {
			pageCount: context.pagesContainer.querySelectorAll(DOCX_RENDERED_PAGE_SELECTOR).length,
			zoom: context.zoom,
		});
		schedulePaginationDiagnostics('rendered-dom-ready');
	}, [onLoadPhase, schedulePaginationDiagnostics]);

	useEffect(() => {
		if (!autosave) {
			session.clearAutosave();
		}
	}, [autosave, session]);

	const scheduleRename = useCallback((name: string) => {
		clearRenameTimeout();
		const scheduledFilePath = file?.path ?? null;
		renameTimeoutRef.current = window.setTimeout(() => {
			renameTimeoutRef.current = null;
			void (async () => {
				try {
					await onDocumentNameChange(name, scheduledFilePath);
				} catch (renameError) {
					const message = renameError instanceof Error ? renameError.message : 'Unknown rename error';
					showNotice('errors:renameFailed', { fileName: file?.name ?? 'document', message });
					setDocumentName(file?.name ?? '');
				}
			})();
		}, 700);
	}, [clearRenameTimeout, file, onDocumentNameChange]);

	const insertCustomTable = useCallback((rows: number, columns: number) => {
		if (editorMode === 'viewing') {
			showNotice('docx:notice.switchToEditForTable');
			return;
		}

		if (!file || !buffer) {
			showNotice('docx:notice.openLoadedToInsertTable');
			return;
		}

		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view) {
			showNotice('docx:notice.editorStillLoading');
			return;
		}

		const inserted = insertTable(
			clampCustomTableSize(rows),
			clampCustomTableSize(columns),
		)(view.state, view.dispatch);

		if (!inserted) {
			warnLog('editor', 'DOCX table insertion was rejected', { file: filePath, rows, columns });
			showNotice('errors:insertTableFailed');
			return;
		}

		debugLog('editor', 'Inserted DOCX table', { file: filePath, rows, columns });
		view.focus();
		setCustomTableDialogOpen(false);
	}, [buffer, editorMode, file, filePath]);

	const openCustomTableDialog = useCallback(() => {
		setCustomTableDialogOpen(true);
	}, []);

	const decorateTableSizeDropdown = useCallback(() => {
		activeDocument.querySelectorAll<HTMLElement>('[role="grid"]').forEach((grid) => {
			appendCustomTableOption(grid, openCustomTableDialog);
		});
	}, [openCustomTableDialog]);

	useEffect(() => {
		let decorateFrame: number | null = null;
		const scheduleDecorateTableSizeDropdown = () => {
			if (decorateFrame !== null) {
				window.cancelAnimationFrame(decorateFrame);
			}
			decorateFrame = window.requestAnimationFrame(() => {
				decorateFrame = null;
				decorateTableSizeDropdown();
			});
		};

		decorateTableSizeDropdown();
		const observer = new MutationObserver(scheduleDecorateTableSizeDropdown);
		observer.observe(activeDocument.body, {
			childList: true,
			subtree: true,
		});

		return () => {
			if (decorateFrame !== null) {
				window.cancelAnimationFrame(decorateFrame);
			}
			observer.disconnect();
		};
	}, [buffer, decorateTableSizeDropdown, filePath, isLoading]);

	const applyFontFamilyToSelection = useCallback((fontFamily: string) => {
		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view || editorMode === 'viewing') {
			return false;
		}

		const result = applyFontFamilyToEditorView(
			view,
			fontFamily,
			preservedTextSelectionRef.current,
		);

		if (result.range) {
			preservedTextSelectionRef.current = result.range;
		}

		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (result.applied) {
			const syncVersion = ++fontFamilyDisplaySyncVersionRef.current;
			scheduleFontFamilySelectDisplaySync(
				editorRoot,
				fontFamily,
				fontFamiliesRef.current,
				() => fontFamilyDisplaySyncVersionRef.current === syncVersion,
			);
		}

		const storedFontMark = view.state.storedMarks?.find((mark) => mark.type.name === 'fontFamily');
		const emptyParagraphSummary = summarizeSelectedEmptyParagraphFonts(
			view.state.doc,
			result.range,
			fontFamily,
		);
		debugLog('editor', 'Applied DOCX font family', {
			file: filePath,
			fontFamily,
			applied: result.applied,
			restoredSelection: Boolean(result.range),
			...emptyParagraphSummary,
			displayName: resolveFontFamilyDisplayName(fontFamily, fontFamiliesRef.current),
			storedFontFamily: getFontFamilyFromMark(storedFontMark),
			displaySyncQueued: result.applied,
		});

		view.focus();
		return result.applied;
	}, [editorMode, filePath]);

	const stopFontSizeHold = useCallback(() => {
		const hold = fontSizeHoldRef.current;
		if (!hold) {
			return;
		}

		if (hold.startTimer !== null) {
			window.clearTimeout(hold.startTimer);
		}
		if (hold.repeatTimer !== null) {
			window.clearTimeout(hold.repeatTimer);
		}

		fontSizeHoldRef.current = null;
	}, []);

	const applyFontSizeStepToSelection = useCallback((direction: FontSizeStepDirection, control: HTMLElement | null) => {
		if (editorMode === 'viewing') {
			return false;
		}

		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view) {
			return false;
		}

		const hold = fontSizeHoldRef.current;
		const currentSize = hold && hold.control === control
			? hold.currentSize
			: readFontSizeControlPoints(control);
		const nextSize = clampFontSizePoints(currentSize + direction);
		if (nextSize === currentSize) {
			return false;
		}

		const applied = setFontSize(fontSizePointsToHalfPoints(nextSize))(view.state, view.dispatch);
		if (!applied) {
			return false;
		}

		if (hold && hold.control === control) {
			hold.currentSize = nextSize;
		}
		updateFontSizeControlDisplay(control, nextSize);
		debugLog('editor', 'Applied DOCX font-size step', {
			file: filePath,
			direction,
			fromPoints: currentSize,
			toPoints: nextSize,
		});
		view.focus();
		return true;
	}, [editorMode, filePath]);

	const startFontSizeHold = useCallback((button: HTMLButtonElement, direction: FontSizeStepDirection) => {
		if (button.disabled || editorMode === 'viewing') {
			return;
		}

		stopFontSizeHold();

		const control = getFontSizeControl(button);
		const currentSize = readFontSizeControlPoints(control);
		const nextSize = clampFontSizePoints(currentSize + direction);
		if (nextSize === currentSize) {
			return;
		}

		const hold: FontSizeHoldState = {
			control,
			currentSize,
			direction,
			repeatCount: 0,
			repeatTimer: null,
			startTimer: null,
		};
		fontSizeHoldRef.current = hold;

		const scheduleRepeat = (delay: number) => {
			hold.repeatTimer = window.setTimeout(() => {
				if (fontSizeHoldRef.current !== hold) {
					return;
				}

				hold.repeatCount += 1;
				if (!applyFontSizeStepToSelection(hold.direction, hold.control)) {
					stopFontSizeHold();
					return;
				}

				const nextDelay = Math.max(
					FONT_SIZE_HOLD_MIN_INTERVAL_MS,
					FONT_SIZE_HOLD_INITIAL_INTERVAL_MS * Math.pow(FONT_SIZE_HOLD_INTERVAL_DECAY, hold.repeatCount),
				);
				scheduleRepeat(nextDelay);
			}, delay);
		};

		if (!applyFontSizeStepToSelection(direction, control)) {
			stopFontSizeHold();
			return;
		}

		hold.startTimer = window.setTimeout(() => {
			if (fontSizeHoldRef.current === hold) {
				scheduleRepeat(FONT_SIZE_HOLD_INITIAL_INTERVAL_MS);
			}
		}, FONT_SIZE_HOLD_INITIAL_DELAY_MS);
	}, [applyFontSizeStepToSelection, editorMode, stopFontSizeHold]);

	const openFontPicker = useCallback(() => {
		if (!fontInputRef.current) {
			showNotice('docx:notice.fontPickerNotReady');
			return;
		}

		fontInputRef.current.click();
	}, []);

	const decorateFontDropdown = useCallback(() => {
		activeDocument.querySelectorAll<HTMLElement>('[role="listbox"]').forEach((listbox) => {
			appendImportFontOption(listbox, openFontPicker);
		});
	}, [openFontPicker]);

	useEffect(() => {
		let decorateFrame: number | null = null;
		const scheduleDecorateFontDropdown = () => {
			if (decorateFrame !== null) {
				window.cancelAnimationFrame(decorateFrame);
			}
			decorateFrame = window.requestAnimationFrame(() => {
				decorateFrame = null;
				decorateFontDropdown();
			});
		};

		decorateFontDropdown();
		const observer = new MutationObserver(scheduleDecorateFontDropdown);
		observer.observe(activeDocument.body, {
			childList: true,
			subtree: true,
		});

		return () => {
			if (decorateFrame !== null) {
				window.cancelAnimationFrame(decorateFrame);
			}
			observer.disconnect();
		};
	}, [buffer, decorateFontDropdown, filePath, isLoading]);

	useEffect(() => {
		const rememberSelectionForFontPicker = (evt: Event) => {
			if (!(evt.target instanceof Element)) {
				return;
			}

			const trigger = evt.target.closest<HTMLElement>(`[${FONT_FAMILY_TRIGGER_ATTRIBUTE}]`);
			const isFontItem = Boolean(isFontPickerMenuItem(evt.target));

			if (!trigger && !isFontItem) {
				return;
			}

			const view = editorRef.current?.getEditorRef()?.getView();
			if (!view) {
				return;
			}

			const remembered = rememberTextSelectionFromView(view);
			if (remembered) {
				preservedTextSelectionRef.current = remembered;
			}
		};

		activeDocument.addEventListener('pointerdown', rememberSelectionForFontPicker, true);
		return () => {
			activeDocument.removeEventListener('pointerdown', rememberSelectionForFontPicker, true);
		};
	}, [buffer, filePath, isLoading]);

	useEffect(() => {
		const syncFontFamilyAfterPicker = (evt: Event) => {
			if (!(evt.target instanceof Element)) {
				return;
			}

			const item = isFontPickerMenuItem(evt.target);
			if (!item) {
				return;
			}

			if (!activeDocument.querySelector(`.${editorClassNameRef.current}`)) {
				return;
			}

			const fontName = item.textContent?.replace(/\s+/g, ' ').trim() ?? '';
			if (!fontName || fontName === IMPORT_FONT_MENU_LABEL) {
				return;
			}

			const fontOption = resolveFontOptionByName(fontName, fontFamiliesRef.current);
			const primaryName = fontOption
				? parsePrimaryFontFamily(fontOption.fontFamily) ?? fontOption.name
				: fontName;

			debugLog('editor', 'Font picker item selected', {
				file: filePath,
				fontName,
				primaryName,
			});

			window.setTimeout(() => {
				window.requestAnimationFrame(() => {
					applyFontFamilyToSelection(primaryName);
				});
			}, 0);
		};

		activeDocument.addEventListener('click', syncFontFamilyAfterPicker, true);
		return () => {
			activeDocument.removeEventListener('click', syncFontFamilyAfterPicker, true);
		};
	}, [applyFontFamilyToSelection, buffer, editorMode, filePath, isLoading]);

	useEffect(() => {
		let lastPointerHandledAt = 0;

		const suppressEvent = (evt: Event) => {
			evt.preventDefault();
			evt.stopImmediatePropagation();
			evt.stopPropagation();
		};

		const getScopedTarget = (target: EventTarget | null) => {
			const stepTarget = getFontSizeStepTarget(target);
			if (!stepTarget) {
				return null;
			}

			const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			if (!editorRoot?.contains(stepTarget.button)) {
				return null;
			}

			return stepTarget;
		};

		const handlePressStart = (evt: MouseEvent | PointerEvent) => {
			if ('button' in evt && evt.button !== 0) {
				return;
			}

			const stepTarget = getScopedTarget(evt.target);
			if (!stepTarget) {
				return;
			}

			suppressEvent(evt);
			if (isPointerEvent(evt)) {
				lastPointerHandledAt = performance.now();
			}
			startFontSizeHold(stepTarget.button, stepTarget.direction);
		};

		const handleMouseDown = (evt: MouseEvent) => {
			const stepTarget = getScopedTarget(evt.target);
			if (stepTarget && performance.now() - lastPointerHandledAt < 100) {
				suppressEvent(evt);
				return;
			}

			handlePressStart(evt);
		};

		const handleClick = (evt: MouseEvent) => {
			if (getScopedTarget(evt.target)) {
				suppressEvent(evt);
			}
		};

		const handleKeyDown = (evt: KeyboardEvent) => {
			if (evt.key !== 'Enter' && evt.key !== ' ') {
				return;
			}

			const stepTarget = getScopedTarget(evt.target);
			if (!stepTarget) {
				return;
			}

			suppressEvent(evt);
			applyFontSizeStepToSelection(stepTarget.direction, getFontSizeControl(stepTarget.button));
		};

		activeDocument.addEventListener('pointerdown', handlePressStart, true);
		activeDocument.addEventListener('mousedown', handleMouseDown, true);
		activeDocument.addEventListener('click', handleClick, true);
		activeDocument.addEventListener('keydown', handleKeyDown, true);
		activeDocument.addEventListener('pointerup', stopFontSizeHold, true);
		activeDocument.addEventListener('pointercancel', stopFontSizeHold, true);
		activeDocument.addEventListener('mouseup', stopFontSizeHold, true);
		const activeView = activeDocument.defaultView ?? window;
		activeView.addEventListener('blur', stopFontSizeHold, false);

		return () => {
			activeDocument.removeEventListener('pointerdown', handlePressStart, true);
			activeDocument.removeEventListener('mousedown', handleMouseDown, true);
			activeDocument.removeEventListener('click', handleClick, true);
			activeDocument.removeEventListener('keydown', handleKeyDown, true);
			activeDocument.removeEventListener('pointerup', stopFontSizeHold, true);
			activeDocument.removeEventListener('pointercancel', stopFontSizeHold, true);
			activeDocument.removeEventListener('mouseup', stopFontSizeHold, true);
			activeView.removeEventListener('blur', stopFontSizeHold, false);
			stopFontSizeHold();
		};
	}, [applyFontSizeStepToSelection, startFontSizeHold, stopFontSizeHold]);

	useEffect(() => {
		if (editorMode === 'viewing') {
			return;
		}

		const suppressEvent = (evt: Event) => {
			evt.preventDefault();
			evt.stopImmediatePropagation();
			evt.stopPropagation();
		};

		const handleLineSpacingPick = (evt: Event) => {
			if ('button' in evt && evt.button !== 0) {
				return;
			}

			const optionTarget = getLineSpacingOptionTarget(evt.target);
			if (!optionTarget) {
				return;
			}

			if (!activeDocument.querySelector(`.${editorClassNameRef.current}`)) {
				return;
			}

			const view = editorRef.current?.getEditorRef()?.getView();
			if (!view) {
				return;
			}

			suppressEvent(evt);
			applyLineSpacingToEditor(view, optionTarget.twips);
			view.focus();
		};

		const handleLineSpacingKeyDown = (evt: KeyboardEvent) => {
			if (evt.key !== 'Enter' && evt.key !== ' ') {
				return;
			}

			handleLineSpacingPick(evt);
		};

		activeDocument.addEventListener('pointerdown', handleLineSpacingPick, true);
		activeDocument.addEventListener('click', handleLineSpacingPick, true);
		activeDocument.addEventListener('keydown', handleLineSpacingKeyDown, true);

		return () => {
			activeDocument.removeEventListener('pointerdown', handleLineSpacingPick, true);
			activeDocument.removeEventListener('click', handleLineSpacingPick, true);
			activeDocument.removeEventListener('keydown', handleLineSpacingKeyDown, true);
		};
	}, [buffer, editorMode, filePath, isLoading]);

	useEffect(() => {
		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		return attachDocxToolbarTooltipManager(editorRoot);
	}, [documentKey]);

	const importFontFile = useCallback(async (fontFile: File) => {
		if (!isSupportedFontFile(fontFile)) {
			showNotice('docx:notice.chooseFontFile');
			return;
		}

		const fontName = getUniqueImportedFontName(fontFile, fontFamilies);
		const fontFamily = cssFontFamilyName(fontName);

		try {
			const fontBuffer = await fontFile.arrayBuffer();
			const loaded = await loadFontFromBuffer(fontName, fontBuffer);
			if (!loaded) {
				showNotice('errors:fontImportFileFailed', { fileName: fontFile.name });
				return;
			}

			setImportedFonts((fonts) => [
				...fonts,
				{ name: fontName, fontFamily, category: 'other' },
			]);

			const wasApplied = applyFontFamilyToSelection(fontFamily);
			debugLog('editor', 'Imported DOCX font', {
				file: filePath,
				fontFileName: fontFile.name,
				fontName,
				bytes: fontBuffer.byteLength,
				appliedToSelection: wasApplied,
			});
			showNotice(wasApplied ? 'docx:notice.fontImportedApplied' : 'docx:notice.fontImported', { fontName });
		} catch (fontError) {
			const message = fontError instanceof Error ? fontError.message : 'Unknown font import error';
			errorLog('editor', `Could not import font ${fontFile.name}`, fontError);
			showNotice('errors:fontImportFailed', { message });
		}
	}, [applyFontFamilyToSelection, filePath, fontFamilies]);

	const handleFontInputChange = useCallback((evt: ChangeEvent<HTMLInputElement>) => {
		const fontFile = evt.currentTarget.files?.[0];
		evt.currentTarget.value = '';
		if (!fontFile) {
			return;
		}

		void importFontFile(fontFile);
	}, [importFontFile]);

	const insertImageFile = useCallback(async (imageFile: File) => {
		if (editorMode === 'viewing') {
			showNotice('docx:notice.switchToEditForImage');
			return;
		}

		if (!file || !buffer) {
			showNotice('docx:notice.openLoadedToInsertImage');
			return;
		}

		try {
			const inserted = await editorAdapter.insertImage(imageFile);
			if (!inserted) {
				showNotice('docx:notice.cannotInsertImageHere');
				return;
			}
			debugLog('editor', 'Inserted image into DOCX', {
				file: filePath,
				imageFileName: imageFile.name,
				imageType: imageFile.type,
				imageBytes: imageFile.size,
				width: inserted.width,
				height: inserted.height,
			});
		} catch (insertError) {
			const message = insertError instanceof Error ? insertError.message : 'Unknown image insert error';
			errorLog('editor', `Could not insert image into ${file.name}`, insertError);
			showNotice('errors:insertImageFailed', { message });
		}
	}, [buffer, editorAdapter, editorMode, file, filePath]);

	const openImagePicker = useCallback(() => {
		if (editorMode === 'viewing') {
			showNotice('docx:notice.switchToEditForImage');
			return;
		}

		if (!file || !buffer) {
			showNotice('docx:notice.openLoadedToInsertImage');
			return;
		}

		if (!imageInputRef.current) {
			showNotice('docx:notice.imagePickerNotReady');
			return;
		}

		imageInputRef.current.click();
	}, [buffer, editorMode, file]);

	const handleImageInputChange = useCallback((evt: ChangeEvent<HTMLInputElement>) => {
		const imageFile = evt.currentTarget.files?.[0];
		evt.currentTarget.value = '';
		if (!imageFile) {
			return;
		}

		void insertImageFile(imageFile);
	}, [insertImageFile]);

	const prepareForExternalReload = useCallback(async () => {
		externalReloadBlockedRef.current = true;
		session.clearAutosave();
		await session.waitForIdle();
	}, [session]);

	const flushPendingSave = useCallback(async () => {
		const current = sessionRef.current;
		if (!current) {
			return true;
		}

		const before = {
			dirty: current.dirty,
			editVersion: current.editVersion,
			saveState: current.saveState,
		};
		debugLog('save', 'DOCX flush pending save before close', {
			file: filePath,
			...before,
		});

		let saved = true;
		if (current.dirty) {
			// Persist scheduled/in-flight edits. Do not clearAutosave alone — that
			// drops the timer and leaves host isDirty true for the close prompt.
			saved = await current.save('autosave');
		}
		await current.waitForIdle();

		const afterDirty = current.dirty;
		debugLog('save', 'DOCX flush pending save settled', {
			file: filePath,
			saved,
			dirtyBefore: before.dirty,
			dirtyAfter: afterDirty,
			editVersion: current.editVersion,
			saveState: current.saveState,
		});
		return saved && !afterDirty;
	}, [filePath]);

	useImperativeHandle(ref, () => ({
		save: () => saveDocument(),
		flushPendingSave,
		prepareForExternalReload,
		exportBuffer: (options?: ExportDocumentBufferOptions) => exportDocumentBuffer(options),
		exportRenderedPdf: () => exportRenderedPdfBuffer(),
		getComments: () => {
			const comments = editorRef.current?.getComments() ?? [];
			return comments.map((comment) => ({
				id: comment.id,
				parentId: comment.parentId ?? null,
				author: comment.author,
				text: summarizeDocxComment(comment).text,
			}));
		},
		pasteFromClipboard: async (options: PasteClipboardOptions) => {
			const view = editorRef.current?.getEditorRef()?.getView();
			const pasted = view ? await pasteClipboardIntoEditor(view, options) : false;
			debugLog('clipboard', 'DOCX paste command completed', {
				file: filePath,
				preserveFormatting: options.preserveFormatting,
				pasted,
				editorReady: Boolean(view),
			});
			return pasted;
		},
		rewriteClipboardTextWithListMarkers: async () => {
			const view = editorRef.current?.getEditorRef()?.getView();
			return view ? await rewritePlainTextClipboardWithListMarkers(view) : false;
		},
		openFind: () => openFindReplaceDialog('find'),
		openFindReplace: () => openFindReplaceDialog('replace'),
		openImagePicker,
		openCustomTableDialog,
		openFontPicker,
		setMode,
		setZoom: (zoom: number) => {
			debugLog('editor', 'DOCX zoom changed', { file: filePath, zoom });
			editorRef.current?.setZoom(zoom);
		},
	}), [exportDocumentBuffer, exportRenderedPdfBuffer, filePath, flushPendingSave, openCustomTableDialog, openFindReplaceDialog, openFontPicker, openImagePicker, prepareForExternalReload, saveDocument, setMode]);

	if (isLoading) {
		return null;
	}

	if (error) {
		return <div>{error}</div>;
	}

	if (!file || !buffer) {
		return null;
	}

	return (
		<>
			<input
				ref={fontInputRef}
				type="file"
				accept={FONT_FILE_ACCEPT}
				aria-label="Import font file"
				style={{ display: 'none' }}
				onChange={handleFontInputChange}
			/>
			<input
				ref={imageInputRef}
				type="file"
				accept="image/*"
				aria-label="Import image file"
				style={{ display: 'none' }}
				onChange={handleImageInputChange}
			/>
			<DocxEditor
				key={documentKey}
				ref={editorRef}
				documentBuffer={buffer}
				mode={editorMode}
				onModeChange={setMode}
				author={authorName}
				i18n={i18n}
				commentsSidebarOpen={commentsSidebarOpen}
				onCommentsSidebarOpenChange={handleCommentsSidebarOpenChange}
				onCommentsChange={handleCommentsChange}
				onCommentAdd={handleCommentAdd}
				onCommentReply={handleCommentReply}
				onCommentDelete={handleCommentDelete}
				onCommentResolve={handleCommentResolve}
				onCommentUnresolve={handleCommentUnresolve}
				initialZoom={defaultZoom}
				className={editorClassNameRef.current}
				colorMode={resolvedEditorTheme}
				showRuler={showRuler}
				disableFindReplaceShortcuts
				externalPlugins={externalPlugins}
				fontFamilies={fontFamilies}
				documentName={documentName}
				documentNameEditable
				pluginSidebarItems={pluginSidebarItems.length > 0 ? pluginSidebarItems : undefined}
				onRenderedDomContextReady={handleRenderedDomContextReady}
				onEditorViewReady={() => {
					onLoadPhase?.('editor-view-ready');
					publishWordCount(editorRef.current?.getEditorRef()?.getView());
					scheduleListMarkerSelectionHighlightSync();
					scheduleCommentsSidebarToggleSync();
					schedulePaginationDiagnostics('editor-view-ready');
				}}
					onSelectionChange={() => {
						const view = editorRef.current?.getEditorRef()?.getView();
						publishWordCount(view);
						const remembered = view ? rememberTextSelectionFromView(view) : null;
						if (remembered) {
							preservedTextSelectionRef.current = remembered;
						}
						const syncVersion = ++fontFamilyDisplaySyncVersionRef.current;
						const fontFamily = view ? getFontFamilyNameFromEditorSelection(view) : null;
						const paragraph = view?.state.selection.$from.parent;
						if (view && paragraph?.type.name === 'paragraph' && paragraph.content.size === 0) {
							const defaults = paragraph.attrs.defaultTextFormatting as {
								fontFamily?: { ascii?: unknown; hAnsi?: unknown };
							} | null | undefined;
							const storedFontMark = view.state.storedMarks?.find((mark) => mark.type.name === 'fontFamily');
							debugLog('editor', 'DOCX empty paragraph font selection changed', {
								file: filePath,
								selectionStart: view.state.selection.from,
								defaultFontFamily: defaults?.fontFamily?.ascii ?? defaults?.fontFamily?.hAnsi ?? null,
								storedFontFamily: getFontFamilyFromMark(storedFontMark),
								resolvedFontFamily: fontFamily,
							});
						}
						if (fontFamily) {
							const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
							scheduleFontFamilySelectDisplaySync(
								editorRoot,
								fontFamily,
								fontFamiliesRef.current,
								() => fontFamilyDisplaySyncVersionRef.current === syncVersion,
							);
						}
						scheduleListMarkerSelectionHighlightSync();
						scheduleCommentsSidebarToggleSync();
					}}
				onDocumentNameChange={(name) => {
					setDocumentName(name);
					scheduleRename(name);
				}}
				renderLogo={() => (
					<SaveButton onClick={() => void saveDocument()} />
				)}
				renderTitleBarRight={() => (
					<SaveStatusIndicator status={saveStatus} />
				)}
					onChange={() => {
					if (
						dirtyTrackingEnabledRef.current
						&& !externalReloadBlockedRef.current
						&& activeDocumentKeyRef.current === documentKey
					) {
						session.markDirty();
					}
						scheduleVerticalRulerMarkerSync(editorRef.current?.getDocument());
						publishWordCount(editorRef.current?.getEditorRef()?.getView());
						scheduleListMarkerSelectionHighlightSync();
					scheduleCommentsSidebarToggleSync();
					schedulePaginationDiagnostics('document-change');
				}}
				onFontsLoaded={() => {
					onLoadPhase?.('editor-fonts-loaded');
					schedulePaginationDiagnostics('fonts-loaded');
				}}
				onSave={() => {}}
				onError={(docxError) => {
					errorLog('render', `Could not render ${file.name}`, docxError);
					warnLog('render', `DOCX editor render error for ${file.name}`, {
						message: docxError.message,
						name: docxError.name,
					});
					console.error('[Native PowerPoint Doc Editor] DOCX render error', docxError);
					showNotice('errors:renderFailed', { fileName: file.name, message: docxError.message });
				}}
			/>
			<CustomTableDialog
				isOpen={customTableDialogOpen}
				onClose={() => setCustomTableDialogOpen(false)}
				onInsert={insertCustomTable}
			/>
			{pluginI18n && findReplaceLabels ? (
			<FindReplaceDialog
				isOpen={findDialogMode !== null}
				labels={findReplaceLabels}
				mode={findDialogMode ?? 'find'}
				searchText={findSearchText}
				replaceText={findReplaceText}
				matchCase={findMatchCase}
				wholeWord={findWholeWord}
				matchCount={findMatches.length}
				currentIndex={currentFindIndex}
				onSearchTextChange={(value) => {
					setFindSearchText(value);
					refreshFindMatches(value, findMatchCase, findWholeWord, 0);
				}}
				onReplaceTextChange={setFindReplaceText}
				onMatchCaseChange={(value) => {
					setFindMatchCase(value);
					refreshFindMatches(findSearchText, value, findWholeWord, currentFindIndex);
				}}
				onWholeWordChange={(value) => {
					setFindWholeWord(value);
					refreshFindMatches(findSearchText, findMatchCase, value, currentFindIndex);
				}}
				onModeChange={setFindDialogMode}
				onNext={() => moveFindMatch(1)}
				onPrevious={() => moveFindMatch(-1)}
				onReplace={replaceCurrentMatch}
				onReplaceAll={replaceAllMatches}
				onClose={() => {
					setFindDialogMode(null);
					setFindMatches([]);
					publishFindHighlights([], 0);
				}}
			/>
			) : null}
		</>
	);
});
