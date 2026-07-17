import { FileView, Menu, Platform, TFile, WorkspaceLeaf, normalizePath, setIcon } from 'obsidian';

import { pptNotice, pptT } from '../../i18n/powerpointNotify';
import type { TranslateFn, TranslateValues } from '../../i18n/translate';

import {
  PresentationEngine,
  type EmptyPrecedingParagraphRemovalResult,
  type GeneratedTextKind,
  type ImageCrop,
  type InsertableShapeGeometry,
  type ParagraphAlignment,
  type ParagraphSplitResult,
  type ParagraphTextRange,
  type PrecedingParagraphMergeResult,
  type RunHighlightInfo,
  type RunStyleChange,
  type RunStyleInfo,
  type RunTarget,
  type TextBoxInsertOrigin,
  type TextRangeDeletionResult,
} from '../../PresentationEngine';
import {
  getImageMimeType,
  ImageCropModal,
  InsertTableModal,
  VaultImageSuggestModal,
  type ImageCropValues
} from '../../PowerPointInsertModals';
import type { ParagraphListStyle } from '../../SlideInsertions';
import {
  inspectPowerPointPackage,
  summarizePackageMessages,
  validatePowerPointPackageStructure,
  type PowerPointPackageInspection
} from '../../PowerPointPackage';
import { createSvgElementFromString, sanitizeSvg, summarizeSvgSecurityIssues, type SvgSecurityIssue } from '../../SvgSecurity';
import type { NativePowerPointSettings } from '../../settings';
import { countDocumentWords, type DocumentWordCount } from '../../documentWordCount';
import type { ShapeTransform } from 'pptx-svg';
import type { ChartDataGrid, ChartDataUpdate } from '../../ChartData';
import type { FontSubstitution } from '../../FontFidelity';
import type { SlideObjectClipboard } from '../../ShapeClipboard';
import { isElement, isNode, isSVGGElement, isSVGTextElement, isSVGTSpanElement } from '../../domGuards';
import { PowerPointPresentController } from '../../PowerPointPresent';
import { exportSlideToPng } from '../../PowerPointExport';
import { InlineTextGeometry } from '../inlineTextGeometry';
import { debugLog, errorLog, logPptxAction, warnLog } from '../../logger';
import { aiUndoStore } from '../../ai/aiUndoStore';
import { renameFileToSiblingName } from '../../vault/renameFlow';
import { scheduleIdleWork } from '../../idleSchedule';

import {
  EDITABLE_POWERPOINT_EXTENSIONS,
  LEGACY_POWERPOINT_EXTENSIONS,
  MACRO_ENABLED_POWERPOINT_EXTENSIONS,
  MODERN_POWERPOINT_EXTENSIONS,
  NATIVE_POWERPOINT_VIEW_TYPE,
  POWERPOINT_EXTENSIONS,
  isEditablePowerPointExtension,
  isMacroEnabledPowerPointExtension,
  isModernPowerPointExtension,
  isPowerPointExtension
} from '../extensions';
import {
  GENERATED_GRID_SELECTOR,
  MAX_ZOOM,
  MIN_ZOOM,
  OBSIDIAN_DOWNLOAD_URL,
  TEXT_TOOLBAR_FONTS,
  TEXT_TOOLBAR_MAX_FONT_SIZE,
  TEXT_TOOLBAR_MIN_FONT_SIZE,
  TEXT_TOOLBAR_SWATCHES
} from '../constants';
import {
  cleanError,
  isWasmGcUnsupportedError
} from '../runtimeCompat';
import { getChromiumVersion } from '../../obsidianRuntime';
import { bindPopoverDismissHandlers, createMenuItem, createPopoverShell, createToolbarIconButton, positionPopoverBelow } from '../../menuControls';
import {
	PPTX_EDITOR_CHROME_CONTEXT_TOOLBAR_CLASS,
	PPTX_EDITOR_CHROME_DOCUMENT_CONTENT_CLASS,
	PPTX_EDITOR_CHROME_DOCUMENT_SURFACE_CLASS,
	PPTX_EDITOR_CHROME_HEADER_CLASS,
	PPTX_EDITOR_CHROME_TOOLBAR_CLASS,
	PPTX_EDITOR_FORMATTING_SURFACE_SELECTOR,
} from '../../editorChromeRegions';
import { FindReplaceController, type FindReplaceHost } from '../findReplaceController';
import { HistoryController, type HistoryHost } from '../historyController';
import { ExportController, type ExportHost } from '../exportController';
import { MenuBarController } from '../menuBarController';
import { ToolbarTooltipController } from '../toolbarTooltipController';
import { SnapController, type SnapHost } from '../snapController';
import { SlideFilmstripController, type SlideFilmstripHost } from '../slideFilmstripController';
import { ArrangeController, type ArrangeHost } from '../arrangeController';
import { InsertController, type InsertHost } from '../insertController';
import { InspectorController, type InspectorHost } from '../inspectorController';
import { SelectionDragController, type SelectionDragHost } from '../selectionDragController';
import { TextToolbarController, type TextToolbarHost } from '../textToolbarController';
import { SaveController, type SaveHost } from '../saveController';
import { PresentationSession, type PresentationSessionEvent } from '../session/PresentationSession';
import { PptxMutationService } from '../mutations/PptxMutationService';
import {
  annotateShapeGroupTextOffsets,
  annotateSlideTextOffsets as stampSlideTextOffsets,
  collectRunSpansByParagraph,
} from '../annotateTextOffsets';
import {
  getInlineWordRange,
  isPrimaryFindShortcut,
  mapEditorOffsetToOoxmlOffset,
  mapEditorRangeToOoxml,
  mapFlatOffsetToRunLine,
  mapFlatRangeToRunLineSegments,
  parsePrimaryFontFamily,
  previousTextForInlineApply,
  redistributeTextAcrossVisualRuns,
  wrapTextForPreview,
  type RunTspanOffset
} from '../textUtils';
import {
  cloneTransform,
  getShapeIndex,
  getSvgIntrinsicSize,
  isEditableShapeIndex,
  isSelectableShapeIndex,
  normalizeSvgForDisplay,
  transformsMatch
} from '../svgUtils';
import type {
  CanvasScrollPosition,
  DragState,
  GeneratedTextEditTarget,
  GroupDragState,
  HandleName,
  HistoryEntry,
  HistorySlideXmlEntry,
  HistoryTransformChange,
  InlineCaretPlacement,
  InlineCaretRow,
  InlineRangeSelection,
  InlineSelectionDrag,
  MarqueeState,
  MenuDropdownEntry,
  PointerPoint,
  ShapeTextEditTarget,
  SvgInlineCaretGeometry,
  SvgInlineSelectionBox,
  SvgRectLike,
  SvgSecurityDecision,
  TextEditTarget,
  TextStyleContext,
  TextToolbarControls,
  ToolbarFormattingSnapshot
} from '../types';

export {
  NATIVE_POWERPOINT_VIEW_TYPE,
  MODERN_POWERPOINT_EXTENSIONS,
  LEGACY_POWERPOINT_EXTENSIONS,
  MACRO_ENABLED_POWERPOINT_EXTENSIONS,
  EDITABLE_POWERPOINT_EXTENSIONS,
  POWERPOINT_EXTENSIONS,
  isPowerPointExtension,
  isModernPowerPointExtension,
  isEditablePowerPointExtension,
  isMacroEnabledPowerPointExtension
};
/**
 * A snapshot of the inline text editor used for in-place (mid-edit) undo/redo.
 * Captures the textarea value plus the selection so undoing a deletion restores
 * both the text and the exact selection that was removed.
 */
interface InlineEditSnapshot {
  value: string;
  selectionStart: number;
  selectionEnd: number;
}

interface SvgFactoryWindow {
  createSvg(tagName: 'line'): SVGLineElement;
  createSvg(tagName: 'rect'): SVGRectElement;
}

interface InlineWholeShapeReplacement {
  shapeIndex: number;
  /** SVG before the temporary all-text preview was cleared; Escape restores it. */
  textFrame: SVGTextElement;
  /**
   * Plain text captured before preview mutation. Commit compares against this —
   * never against live SVG after preview recovery (that falsely skips OOXML writes).
   */
  baselineText: string;
}

const INLINE_EDIT_HISTORY_LIMIT = 200;
const ROTATION_SNAP_THRESHOLD_DEGREES = 3;

export class NativePowerPointView extends FileView {
  private readonly t: TranslateFn = (key, values) => pptT(key, values);
  private tb(suffix: string, values?: TranslateValues): string {
    return this.t(`powerpoint:toolbar.${suffix}`, values);
  }

  private readonly getSettings: () => NativePowerPointSettings;

  private engine: PresentationEngine | null = null;
  private loadedFile: TFile | null = null;
  private sourcePackage: PowerPointPackageInspection | null = null;
  private sourceBuffer: ArrayBuffer | null = null;
  private readonly session: PresentationSession;

  /** Test/debug access to the session-owned save controller. */
  get saveController(): SaveController {
    return this.session.saveController;
  }
  private get currentSlide(): number {
    return this.session.currentSlide;
  }
  private set currentSlide(value: number) {
    this.session.setCurrentSlide(value);
  }
  private zoomLevel = 1;
  private selectedShapeIndex: number | null = null;
  private selectedShapeIndices = new Set<number>();
  private lastInteractionRegion: 'canvas' | 'thumbnails' = 'canvas';
  private selectedTransform: ShapeTransform | null = null;
  private marquee: MarqueeState | null = null;
  private marqueeEl: HTMLElement | null = null;
  /** Passive union outline for the shapes currently hit by an active marquee. */
  private marqueeSelectionPreview: HTMLElement | null = null;
  private groupDrag: GroupDragState | null = null;
  private multiSelectionBoxes: HTMLElement[] = [];
  private suppressNextClick = false;
  private isViewOnly = false;
  private viewOnlyReason = '';
  private isLoading = false;
  private filmstripRenderScheduled = false;
  private filmstripRendered = false;
  private isNavigatingSlide = false;
  private isTearingDownEditor = false;
  private slideRenderGeneration = 0;
  private textCommitPromise: Promise<void> | null = null;
  /** Prevent repeated Enter presses from racing a structural text mutation. */
  private paragraphSplitPromise: Promise<void> | null = null;
  /** Prevent repeated Backspace presses from racing a structural text mutation. */
  private paragraphRemovalPromise: Promise<void> | null = null;
  /** Prevent repeated Delete/Backspace presses from racing an inline range mutation. */
  private rangeDeletionPromise: Promise<void> | null = null;
  private dragState: DragState | null = null;
  private activeEditor: HTMLTextAreaElement | null = null;
  /** Last text intentionally written by this inline-edit transaction. */
  private activeEditorCanonicalText: string | null = null;
  private activeEditorInitialWordCount = 0;
  /** True only after a user text mutation, never after a formatting re-render. */
  private activeEditorTextDirty = false;
  private activeEditorCommit: (() => Promise<void>) | null = null;
  private activeInlineCaret: SVGLineElement | null = null;
  private activeInlineSelectionRects: SVGRectElement[] = [];
  private inlineWholeShapeSelection: string | null = null;
  private inlineWholeShapeSelected = false;
  /** A Cmd/Ctrl+A edit commits through updateShapeText, never one stale paragraph. */
  private inlineWholeShapeReplacement: InlineWholeShapeReplacement | null = null;
  private inlineRangeSelection: InlineRangeSelection | null = null;
  private inlineSelectionDrag: InlineSelectionDrag | null = null;
  private lastInlineCaretPlacement: InlineCaretPlacement | null = null;
  private suppressNextTextClick = false;
  private activeInlineCaretRow: InlineCaretRow | null = null;
  /** Visual line count of the live SVG preview; logged only when it changes. */
  private activeInlinePreviewLineCount: number | null = null;
  private activeEditorTarget: SVGTextElement | SVGTSpanElement | null = null;
  private inlineUndoStack: InlineEditSnapshot[] = [];
  private inlineRedoStack: InlineEditSnapshot[] = [];
  /** Prevent a failed inline restore from falling through to document history. */
  private lastInlineHistoryRestoreFailed = false;
  private lastSelectionOverlayDragLogAt = 0;
  private skippedSelectionOverlayDragLogs = 0;
  private lastInlineSelectionRenderedLogAt = 0;
  private lastInlineSelectionRenderedLogKey: string | null = null;
  private skippedInlineSelectionRenderedLogs = 0;
  private svgSecurityDecision: SvgSecurityDecision = null;
  private presentationWordCount = 0;
  private presentationWordCountEditVersion = -1;

  private layoutEl: HTMLElement | null = null;
  private rootEl: HTMLElement | null = null;
  private headerTitleEl: HTMLElement | null = null;
  private zoomLevelEl: HTMLElement | null = null;
  private thumbnailContainer: HTMLElement | null = null;
  private canvasPane: HTMLElement | null = null;
  private readonly inlineGeometry = new InlineTextGeometry(() => this.canvasPane);
  private slideSurface: HTMLElement | null = null;
  private inspectorEl: HTMLElement | null = null;
  private statusEl: HTMLElement | null = null;
  private slideCounterEl: HTMLElement | null = null;
  private svgEl: SVGSVGElement | null = null;
  private selectionOverlay: HTMLElement | null = null;
  private canvasResizeObserver: ResizeObserver | null = null;
  private editButtons: HTMLButtonElement[] = [];
  private xInput: HTMLInputElement | null = null;
  private yInput: HTMLInputElement | null = null;
  private widthInput: HTMLInputElement | null = null;
  private heightInput: HTMLInputElement | null = null;
  private rotationInput: HTMLInputElement | null = null;
  private hasShownGeneratedTextNotice = false;
  private fontSubstitutions: FontSubstitution[] = [];
  private objectClipboard: SlideObjectClipboard | null = null;
  private copyButton: HTMLButtonElement | null = null;
  private pasteButton: HTMLButtonElement | null = null;
  private duplicateButton: HTMLButtonElement | null = null;
  private distributeButtons: HTMLButtonElement[] = [];
  private zOrderButtons: HTMLButtonElement[] = [];
  private groupButton: HTMLButtonElement | null = null;
  private ungroupButton: HTMLButtonElement | null = null;
  private imageFileInput: HTMLInputElement | null = null;
  private replaceImageFileInput: HTMLInputElement | null = null;
  private insertTableButton: HTMLButtonElement | null = null;
  private pendingReplaceShapeIndex: number | null = null;
  private activeInsertMenu: HTMLElement | null = null;
  private activeShapeTextTarget: ShapeTextEditTarget | null = null;
  private activeTextStyleTarget: ShapeTextEditTarget | null = null;
  private textToolbarEl: HTMLElement | null = null;
  private textToolbarControls: TextToolbarControls | null = null;
  private topFontButton: HTMLButtonElement | null = null;
  private topFontLabel: HTMLElement | null = null;
  private runHighlightRects: SVGRectElement[] = [];
  private pendingHighlightClear: {
    slide: number;
    shapeIndex: number;
    paragraphs: Set<number>;
    ranges: ParagraphTextRange[] | null;
  } | null = null;
  private textToolbarShapeIndex: number | null = null;
  private currentRunStyle: RunStyleInfo | null = null;
  private textColorValue = '000000';
  private textHighlightValue = 'FFFF00';
  private activeToolbarPopover: HTMLElement | null = null;
  private toolbarPopoverCleanup: (() => void) | null = null;
  private toolbarFormattingSnapshot: ToolbarFormattingSnapshot | null = null;
  private presentController: PowerPointPresentController | null = null;
  private readonly findController: FindReplaceController;
  private readonly historyController: HistoryController;
  private readonly exportController: ExportController;
  private readonly snapController: SnapController;
  private readonly slideFilmstripController: SlideFilmstripController;
  private readonly arrangeController: ArrangeController;
  private readonly insertController: InsertController;
  private readonly inspectorController: InspectorController;
  private readonly selectionDragController: SelectionDragController;
  private readonly textToolbarController: TextToolbarController;
  private readonly menuBar = new MenuBarController();
  private readonly toolbarTooltips = new ToolbarTooltipController();

  constructor(
    leaf: WorkspaceLeaf,
    getSettings: () => NativePowerPointSettings,
    private onWordCountChange: (wordCount: DocumentWordCount) => void = () => {},
    private onWordCountClear: () => void = () => {},
  ) {
    super(leaf);
    this.getSettings = getSettings;
    this.historyController = new HistoryController(this.createHistoryHost());
    this.session = new PresentationSession(this.createSaveHost(), {
      history: {
        undo: () => {
          if (this.activeEditor || !this.historyController.canUndo || this.historyController.isRestoring) {
            return false;
          }
          void this.historyController.undo();
          return true;
        },
        redo: () => {
          if (this.activeEditor || !this.historyController.canRedo || this.historyController.isRestoring) {
            return false;
          }
          void this.historyController.redo();
          return true;
        }
      },
      mutationExecutor: new PptxMutationService(() => this.engine),
    });
    this.session.subscribe((event) => this.handlePresentationSessionEvent(event));
    this.findController = new FindReplaceController(this.createFindReplaceHost());
    this.exportController = new ExportController(this.createExportHost());
    this.snapController = new SnapController(this.createSnapHost());
    this.slideFilmstripController = new SlideFilmstripController(this.createSlideFilmstripHost());
    this.arrangeController = new ArrangeController(this.createArrangeHost());
    this.insertController = new InsertController(this.createInsertHost());
    this.inspectorController = new InspectorController(this.createInspectorHost());
    this.selectionDragController = new SelectionDragController(this.createSelectionDragHost());
    this.textToolbarController = new TextToolbarController(this.createTextToolbarHost());
    this.addChild(this.menuBar);
    this.addChild(this.toolbarTooltips);
  }

  /**
   * Bridges the slide filmstrip subsystem to the view's shared editor state.
   * Built as an adapter object (rather than `implements SlideFilmstripHost`) so
   * the view's own members can remain `private`; this closure can read them
   * because it is lexically inside the class.
   */
  private createSlideFilmstripHost(): SlideFilmstripHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get engine() { return getView().engine; },
      get thumbnailContainer() { return getView().thumbnailContainer; },
      get isLoading() { return getView().isLoading; },
      get currentSlide() { return getView().session.currentSlide; },
      set currentSlide(value: number) { getView().session.setCurrentSlide(value); },
      get lastInteractionRegion() { return getView().lastInteractionRegion; },
      set lastInteractionRegion(value: 'canvas' | 'thumbnails') { getView().lastInteractionRegion = value; },
      get selectedShapeIndex() { return getView().selectedShapeIndex; },
      set selectedShapeIndex(value: number | null) { getView().selectedShapeIndex = value; },
      get selectedTransform() { return getView().selectedTransform; },
      set selectedTransform(value: ShapeTransform | null) { getView().selectedTransform = value; },
      get slideRenderGeneration() { return getView().slideRenderGeneration; },
      set slideRenderGeneration(value: number) { getView().slideRenderGeneration = value; },
      get isNavigatingSlide() { return getView().isNavigatingSlide; },
      set isNavigatingSlide(value: boolean) { getView().isNavigatingSlide = value; },
      canEdit: () => getView().canEdit(),
      ensureEditable: (action) => getView().ensureEditable(action),
      finishInlineTextEditing: (reason) => getView().finishInlineTextEditing(reason),
      captureHistoryEntry: (label) => getView().captureHistoryEntry(label),
      recordHistoryEntry: (entry) => getView().recordHistoryEntry(entry),
      markDirty: () => getView().markDirty(),
      renderCurrentSlide: (keepSelection, expectedGeneration) => getView().renderCurrentSlide(keepSelection, expectedGeneration),
      clearSelection: () => getView().clearSelection(),
      renderInspector: () => getView().renderInspector(),
      prepareSvgForRender: (svg, isThumbnail) => getView().prepareSvgForRender(svg, isThumbnail),
      createNativeMenu: () => getView().createNativeMenu()
    };
  }

  /**
   * Bridges snapping to the view's slide geometry. Built as an adapter object
   * (rather than `implements SnapHost`) so the view's own members can remain
   * `private`; this closure can read them because it is lexically inside the
   * class.
   */
  private createSnapHost(): SnapHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get engine() { return getView().engine; },
      get svgEl() { return getView().svgEl; },
      get canvasPane() { return getView().canvasPane; },
      get slideSurface() { return getView().slideSurface; },
      emuPointToPane: (emuX, emuY) => getView().emuPointToPane(emuX, emuY),
      getElementBox: (element) => getView().getElementBox(element)
    };
  }

  private createTextToolbarHost(): TextToolbarHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get engine() { return getView().engine; },
      get svgEl() { return getView().svgEl; },
      get canvasPane() { return getView().canvasPane; },
      get currentSlide() { return getView().currentSlide; },
      set currentSlide(value: number) { getView().currentSlide = value; },
      get activeEditor() { return getView().activeEditor; },
      get activeEditorTarget() { return getView().activeEditorTarget; },
      get activeTextStyleTarget() { return getView().activeTextStyleTarget; },
      get currentRunStyle() { return getView().currentRunStyle; },
      set currentRunStyle(value: RunStyleInfo | null) { getView().currentRunStyle = value; },
      ensureEditable: (action) => getView().ensureEditable(action),
      canEdit: () => getView().canEdit(),
      getTextStyleContext: () => getView().getTextStyleContext(),
      getElementBox: (element) => getView().getElementBox(element),
      getSelectedBox: () => getView().getSelectedBox(),
      getStoredInlineSelectionRanges: (shapeIndex) => getView().getStoredInlineSelectionRanges(shapeIndex),
      getSelectedRangeFontSizePt: (shapeIndex, ranges) => getView().getSelectedRangeFontSizePt(shapeIndex, ranges),
      applyRunStyle: (change) => getView().applyRunStyle(change),
      applyAlignment: (align) => getView().applyAlignment(align),
      flushActiveEditor: () => getView().flushActiveEditor()
    };
  }

  private createSelectionDragHost(): SelectionDragHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get session() { return getView().session; },
      get engine() { return getView().engine; },
      get svgEl() { return getView().svgEl; },
      get canvasPane() { return getView().canvasPane; },
      get selectedShapeIndex() { return getView().selectedShapeIndex; },
      set selectedShapeIndex(value: number | null) { getView().selectedShapeIndex = value; },
      get selectedShapeIndices() { return getView().selectedShapeIndices; },
      get selectedTransform() { return getView().selectedTransform; },
      set selectedTransform(value: ShapeTransform | null) { getView().selectedTransform = value; },
      get snapController() { return getView().snapController; },
      get suppressNextClick() { return getView().suppressNextClick; },
      set suppressNextClick(value: boolean) { getView().suppressNextClick = value; },
      get currentSlide() { return getView().currentSlide; },
      set currentSlide(value: number) { getView().currentSlide = value; },
      ensureEditable: (action) => getView().ensureEditable(action),
      canEdit: () => getView().canEdit(),
      getSelectedShapeElement: () => getView().getSelectedShapeElement(),
      getElementBox: (element) => getView().getElementBox(element),
      emuPointToPane: (x, y) => getView().emuPointToPane(x, y),
      applyMultiSelection: (indices) => getView().applyMultiSelection(indices),
      clearSelection: (options) => getView().clearSelection(options),
      captureHistoryEntry: (label) => getView().captureHistoryEntry(label),
      recordHistoryEntry: (entry) => getView().recordHistoryEntry(entry),
      markDirty: () => getView().markDirty(),
      renderCurrentSlide: (keepSelection) => getView().renderCurrentSlide(keepSelection),
      renderEditedShape: (shapeIndex) => getView().renderEditedShape(shapeIndex),
      renderThumbnails: () => getView().renderThumbnails(),
      updateInspectorValues: () => getView().updateInspectorValues(),
      updateTextToolbar: () => getView().updateTextToolbar()
    };
  }

  private createArrangeHost(): ArrangeHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get session() { return getView().session; },
      get engine() { return getView().engine; },
      get svgEl() { return getView().svgEl; },
      get currentSlide() { return getView().currentSlide; },
      set currentSlide(value: number) { getView().currentSlide = value; },
      ensureEditable: (action) => getView().ensureEditable(action),
      canEdit: () => getView().canEdit(),
      getSelectedIndices: () => getView().getSelectedIndices(),
      captureHistoryEntry: (label) => getView().captureHistoryEntry(label),
      recordHistoryEntry: (entry) => getView().recordHistoryEntry(entry),
      markDirty: () => getView().markDirty(),
      renderCurrentSlide: (keepSelection) => getView().renderCurrentSlide(keepSelection),
      renderThumbnails: () => getView().renderThumbnails(),
      applyMultiSelection: (indices) => getView().applyMultiSelection(indices),
      selectShape: (shapeIndex) => getView().selectShape(shapeIndex),
      commitGroupTransforms: (updates, label) => getView().commitGroupTransforms(updates, label),
      createIconButton: (container, icon, label, onClick) => getView().createIconButton(container, icon, label, onClick),
      updateToolbarButton: (button, enabled) => getView().updateObjectClipboardButton(button, enabled)
    };
  }

  private createInsertHost(): InsertHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get session() { return getView().session; },
      get engine() { return getView().engine; },
      get app() { return getView().app; },
      get layoutEl() { return getView().layoutEl; },
      get selectedShapeIndex() { return getView().selectedShapeIndex; },
      get activeEditorTarget() { return getView().activeEditorTarget; },
      get currentSlide() { return getView().currentSlide; },
      set currentSlide(value: number) { getView().currentSlide = value; },
      ensureEditable: (action) => getView().ensureEditable(action),
      captureHistoryEntry: (label) => getView().captureHistoryEntry(label),
      recordHistoryEntry: (entry) => getView().recordHistoryEntry(entry),
      markDirty: () => getView().markDirty(),
      renderCurrentSlide: (keepSelection) => getView().renderCurrentSlide(keepSelection),
      renderEditedShape: (shapeIndex) => getView().renderEditedShape(shapeIndex),
      renderThumbnails: () => getView().renderThumbnails(),
      selectShape: (shapeIndex) => getView().selectShape(shapeIndex),
      selectShapeForTextEditing: (shapeIndex) => getView().selectShapeForTextEditing(shapeIndex),
      startTextEditor: () => getView().startTextEditor(),
      createEditIconButton: (container, icon, label, onClick) => getView().createEditIconButton(container, icon, label, onClick),
      getTextEditTarget: (target) => getView().getTextEditTarget(target),
      registerDocumentPointerDown: (handler, capture) => getView().registerDomEvent(activeDocument, 'pointerdown', handler, capture),
      openToolbarPopover: (anchor, build) => getView().openToolbarPopover(anchor, build),
      bindToolbarButton: (button, action) => getView().bindToolbarButton(button, action),
      closeToolbarPopover: () => getView().closeToolbarPopover()
    };
  }

  private createInspectorHost(): InspectorHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get session() { return getView().session; },
      get engine() { return getView().engine; },
      get inspectorEl() { return getView().inspectorEl; },
      get selectedShapeIndex() { return getView().selectedShapeIndex; },
      get selectedShapeIndices() { return getView().selectedShapeIndices; },
      get selectedTransform() { return getView().selectedTransform; },
      get currentSlide() { return getView().currentSlide; },
      set currentSlide(value: number) { getView().currentSlide = value; },
      get isViewOnly() { return getView().isViewOnly; },
      get viewOnlyReason() { return getView().viewOnlyReason; },
      get fontSubstitutions() { return getView().fontSubstitutions; },
      ensureEditable: (action) => getView().ensureEditable(action),
      canEdit: () => getView().canEdit(),
      getSelectedShapeElement: () => getView().getSelectedShapeElement(),
      captureHistoryEntry: (label) => getView().captureHistoryEntry(label),
      recordHistoryEntry: (entry) => getView().recordHistoryEntry(entry),
      renderCurrentSlide: (keepSelection) => getView().renderCurrentSlide(keepSelection),
      renderThumbnails: () => getView().renderThumbnails(),
      renderEditedShape: (shapeIndex) => getView().renderEditedShape(shapeIndex),
      commitTransform: (transform) => getView().commitTransform(transform)
    };
  }

  /** Bridges saving to the view's loaded presentation and status element. */
  private createSaveHost(): SaveHost {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get app() { return getView().app; },
      get statusEl() { return getView().statusEl; },
      getSettings: () => getView().getSettings(),
      getFile: () => getView().loadedFile || getView().file,
      getLoadedFile: () => getView().loadedFile,
      getEngine: () => getView().engine,
      getSourcePackage: () => getView().sourcePackage,
      getSourceBuffer: () => getView().sourceBuffer,
      setSource: (sourcePackage, sourceBuffer) => {
        getView().sourcePackage = sourcePackage;
        getView().sourceBuffer = sourceBuffer;
      },
      isCurrentPresentation: (engine, file) =>
        getView().engine === engine && getView().loadedFile?.path === file.path,
      ensureEditable: (action) => getView().ensureEditable(action),
      getViewOnlyReason: () => getView().viewOnlyReason
    };
  }

  /**
   * Bridges the export subsystem to the view's shared state. Built as an adapter
   * object (rather than `implements ExportHost`) so the view's own members can
   * remain `private`; this closure can read them because it is lexically inside
   * the class.
   */
  private createExportHost(): ExportHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get engine() { return getView().engine; },
      get currentSlide() { return getView().currentSlide; },
      get ownerDocument() { return getView().contentEl.ownerDocument; },
      get app() { return getView().app; },
      get sourceFile() { return getView().loadedFile || getView().file; },
      buildSlideSvgElement: (index) => getView().buildSlideSvgElement(index),
      collectSvgElements: (indices) => getView().collectExportSvgElements(indices),
      createNativeMenu: () => getView().createNativeMenu()
    };
  }

  /**
   * Bridges the undo/redo subsystem to the view's shared editor state. Built as
   * an adapter object (rather than `implements HistoryHost`) so the view's own
   * members can remain `private`; this closure can read them because it is
   * lexically inside the class.
   */
  private createHistoryHost(): HistoryHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get engine() { return getView().engine; },
      get currentSlide() { return getView().currentSlide; },
      set currentSlide(value: number) { getView().currentSlide = value; },
      get activeEditor() { return getView().activeEditor; },
      canUndoInlineEdit: () => getView().hasInlineHistory('undo'),
      canRedoInlineEdit: () => getView().hasInlineHistory('redo'),
      ensureEditable: (action) => getView().ensureEditable(action),
      canEdit: () => getView().canEdit(),
      clearAutosave: () => getView().clearAutosave(),
      clearDragState: () => {
        getView().dragState = null;
        getView().selectionDragController.clearDragState();
      },
      clearSelection: () => getView().clearSelection(),
      markDirty: () => getView().markDirty(),
      renderCurrentSlide: (keepSelection) => getView().renderCurrentSlide(keepSelection),
      renderThumbnails: () => getView().renderThumbnails(),
      scheduleThumbnailRefresh: (indices) => getView().scheduleThumbnailRefresh(indices)
    };
  }

  /**
   * Bridges the find/replace subsystem to the view's shared editor state. Built
   * as an adapter object (rather than `implements FindReplaceHost`) so the
   * view's own members can remain `private`; this closure can read them because
   * it is lexically inside the class.
   */
  private createFindReplaceHost(): FindReplaceHost & { t: TranslateFn } {
    const getView = (): NativePowerPointView => this;
    return {
      t: this.t,
      get session() { return getView().session; },
      get engine() { return getView().engine; },
      get isLoading() { return getView().isLoading; },
      get currentSlide() { return getView().currentSlide; },
      get activeEditor() { return getView().activeEditor; },
      get svgEl() { return getView().svgEl; },
      ensureEditable: (action) => getView().ensureEditable(action),
      clearSelection: () => getView().clearSelection(),
      captureHistoryEntry: (label) => getView().captureHistoryEntry(label),
      recordHistoryEntry: (entry) => getView().recordHistoryEntry(entry),
      renderCurrentSlide: (keepSelection) => getView().renderCurrentSlide(keepSelection),
      renderThumbnails: () => getView().renderThumbnails(),
      getShapeTextParagraphs: (shape) => getView().getShapeTextParagraphs(shape),
      getParagraphLeafText: (element) => getView().getParagraphLeafText(element),
      getSvgInlineSelectionBoxes: (element, start, end) => getView().getSvgInlineSelectionBoxes(element, start, end),
      formatSvgNumber: (value) => getView().formatSvgNumber(value)
    };
  }

  /** Re-render lightweight chrome when shared presentation state changes. */
  private handlePresentationSessionEvent(event: PresentationSessionEvent): void {
	if (event.type === 'save' && event.snapshot.editVersion !== this.presentationWordCountEditVersion) {
		this.refreshPresentationWordCount();
	}
	if (event.type === 'selection' || event.type === 'slide') {
		this.publishPresentationWordCount();
	}
    if (event.type === 'slide') {
      this.updateSlideCounter();
      return;
    }
    this.updateEditingAvailability();
  }

	private refreshPresentationWordCount(): void {
		const engine = this.engine;
		const renderSlide = (engine as Partial<PresentationEngine> | null)?.renderSlide;
		if (!engine || typeof renderSlide !== 'function' || typeof DOMParser === 'undefined') {
			return;
		}

		const parser = new DOMParser();
		const text = Array.from({ length: engine.slideCount }, (_, slideIndex) => {
			const svg = renderSlide.call(engine, slideIndex).svg;
			const slide = parser.parseFromString(svg, 'image/svg+xml');
			return slide.documentElement.textContent ?? '';
		}).join('\n');
		this.presentationWordCount = countDocumentWords(text);
		this.presentationWordCountEditVersion = this.session.editVersion;
		this.publishPresentationWordCount();
	}

	private getSelectedPresentationText(): string | null {
		if (this.inlineWholeShapeSelection !== null) {
			return this.inlineWholeShapeSelection;
		}

		if (this.activeEditor) {
			const start = Math.min(this.activeEditor.selectionStart ?? 0, this.activeEditor.selectionEnd ?? 0);
			const end = Math.max(this.activeEditor.selectionStart ?? 0, this.activeEditor.selectionEnd ?? 0);
			return end > start ? this.activeEditor.value.slice(start, end) : null;
		}

		const selectedIndexes = this.getSelectedIndices();
		if (selectedIndexes.length === 0) {
			return null;
		}

		return selectedIndexes.map((shapeIndex) => (
			this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`)?.textContent ?? ''
		)).join('\n');
	}

	private publishPresentationWordCount(): void {
		if (!this.engine) {
			return;
		}

		const selectedText = this.getSelectedPresentationText();
		const liveTotal = this.activeEditor
			? Math.max(0, this.presentationWordCount - this.activeEditorInitialWordCount + countDocumentWords(this.activeEditor.value))
			: this.presentationWordCount;
		this.onWordCountChange({
			totalWords: liveTotal,
			selectedWords: selectedText === null ? null : countDocumentWords(selectedText),
		});
	}

  canAcceptExtension(extension: string): boolean {
    return isPowerPointExtension(extension);
  }

  getViewType(): string {
    return NATIVE_POWERPOINT_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.loadedFile?.basename || this.file?.basename || this.tb('nativePowerPoint');
  }

  getIcon(): string {
    return 'presentation';
  }

  async onOpen(): Promise<void> {
    debugLog('view', 'Opening PowerPoint view');
    this.contentEl.empty();
    this.contentEl.addClass('native-powerpoint-view');
    this.createLayout();
    this.registerKeyboardHandlers();
    if (this.getSettings().showInspector) {
      this.renderInspector();
    }
  }

  async onLoadFile(file: TFile): Promise<void> {
    debugLog('file', 'PowerPoint file load requested', { file: file.path });
    if (this.engine && this.loadedFile && this.loadedFile.path !== file.path) {
      const preserved = await this.preserveUnsavedChangesForTeardown('switching files');
      if (!preserved) {
        pptNotice('powerpoint:notice.switchFilesPreserveFailed', { fileName: this.loadedFile.name });
        return;
      }

      this.resetLoadedPresentation();
    }

    await this.loadPresentation(file);
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    debugLog('file', 'PowerPoint file unload requested', { file: _file.path });
    const preserved = await this.preserveUnsavedChangesForTeardown('switching files');
    if (preserved) {
      this.resetLoadedPresentation();
    }
  }

  async onClose(): Promise<void> {
    debugLog('view', 'Closing PowerPoint view', {
      file: this.loadedFile?.path ?? this.file?.path ?? null,
      dirty: this.session.dirty
    });
    this.findController.dispose();
    this.presentController?.dispose();
    this.presentController = null;
    this.slideFilmstripController.dispose();

    const preserved = await this.preserveUnsavedChangesForTeardown('closing the view');
    if (!preserved) {
      pptNotice('powerpoint:notice.closeUnsafe');
      return;
    }

    this.resetLoadedPresentation();
    this.contentEl.removeClass('native-powerpoint-view');
    this.file = null;
    debugLog('view', 'Closed PowerPoint view');
  }

  async saveCurrentPresentation(source: 'manual' | 'autosave' = 'manual'): Promise<boolean> {
    return this.session.save(source);
  }

  getAgentSaveError(): string | null {
    return this.saveController.lastSaveError;
  }

  /**
   * Flush dirty presentation state before the development hot-reloader disables
   * this plugin. Obsidian view teardown is asynchronous, so the reloader must
   * await this barrier instead of relying on `onClose()` racing plugin startup.
   *
   * This is an explicit save boundary even when background autosave is disabled:
   * if the source file cannot be updated, the caller must abort the reload and
   * leave the live view intact.
   */
  async saveBeforePluginReload(): Promise<boolean> {
    this.clearAutosave();
    if (!this.session.dirty) return true;

    debugLog('save', 'Saving PowerPoint before plugin reload', {
      file: this.loadedFile?.path ?? this.file?.path ?? null,
      editVersion: this.session.editVersion
    });
    return this.saveCurrentPresentation('manual');
  }

  getLoadedPresentationPath(): string | null {
    return (this.loadedFile || this.file)?.path ?? null;
  }

  getPresentationEngineForAgent(): PresentationEngine | null {
    return this.engine;
  }

  canAgentEdit(): boolean {
    return this.canEdit();
  }

  async runAgentEditBatch(
    label: string,
    mutate: (engine: PresentationEngine) => Promise<number[]>,
  ): Promise<void> {
    if (!this.engine || !this.ensureEditable('apply AI edit')) {
      throw new Error('Cannot apply AI edit to this presentation.');
    }

    const history = await this.captureHistoryEntry(label);
    try {
      const affectedSlideIndices = await mutate(this.engine);
      this.recordHistoryEntry(history);
      this.markDirty();
      await this.renderCurrentSlide();
      if (affectedSlideIndices.length > 0) {
        this.slideFilmstripController.scheduleThumbnailRefresh(affectedSlideIndices);
      } else {
        this.slideFilmstripController.scheduleThumbnailRefresh(this.currentSlide);
      }
    } catch (error) {
      if (history.kind === 'snapshot') {
        await this.engine.restoreSnapshot(history.buffer);
        await this.renderCurrentSlide();
      }
      throw error;
    }
  }

  canUndoAgentEdit(): boolean {
    return this.historyController.canUndo;
  }

  canRedoAgentEdit(): boolean {
    return this.historyController.canRedo;
  }

  async undoAgentEdit(): Promise<boolean> {
    if (!this.historyController.canUndo) {
      return false;
    }
    this.session.undo();
    return true;
  }

  async redoAgentEdit(): Promise<boolean> {
    if (!this.historyController.canRedo) {
      return false;
    }
    this.session.redo();
    return true;
  }

  private importPendingAgentUndoHistory(): void {
    const file = this.loadedFile || this.file;
    if (!file || !this.engine) {
      return;
    }

    for (const entry of aiUndoStore.drainUndo(file.path)) {
      if (entry.before.kind !== 'pptx') {
        continue;
      }
      this.historyController.record({
        kind: 'snapshot',
        buffer: entry.before.buffer,
        currentSlide: entry.before.currentSlide,
        label: entry.label,
      });
    }
  }

  private canEdit(): boolean {
    const file = this.loadedFile || this.file;
    return Boolean(
      this.engine &&
      file &&
      !this.isViewOnly &&
      isEditablePowerPointExtension(file.extension)
    );
  }

  private ensureEditable(action: string): boolean {
    if (this.canEdit()) return true;

    const reason = this.viewOnlyReason || pptT('powerpoint:notice.viewOnlyFile');
    debugLog('view', 'Blocked PowerPoint edit action', {
      action,
      file: (this.loadedFile || this.file)?.path,
      reason
    });
    pptNotice('powerpoint:notice.cannotAction', { action, reason });
    return false;
  }

  private shouldOpenViewOnly(file: TFile, sourcePackage: PowerPointPackageInspection): boolean {
    return isMacroEnabledPowerPointExtension(file.extension) || sourcePackage.hasVbaProject;
  }

  private getViewOnlyReason(file: TFile, sourcePackage: PowerPointPackageInspection): string {
    if (isMacroEnabledPowerPointExtension(file.extension)) {
      return this.t('powerpoint:notice.macroEnabledViewOnly');
    }

    if (sourcePackage.hasVbaProject) {
      return this.t('powerpoint:notice.hasVbaProjectViewOnly');
    }

    return '';
  }

  private createLayout(): void {
    this.contentEl.empty();

    const root = this.contentEl.createDiv({ cls: 'native-powerpoint-root' });
    this.rootEl = root;
    this.applyThemeClass();
    this.createHeaderBar(root);
    this.layoutEl = root.createDiv({ cls: 'native-powerpoint-layout' });

    const sidebar = this.layoutEl.createDiv({ cls: 'native-powerpoint-sidebar' });
    this.registerDomEvent(sidebar, 'pointerdown', () => {
      this.lastInteractionRegion = 'thumbnails';
    }, true);
    const sidebarHeader = sidebar.createDiv({ cls: 'native-powerpoint-sidebar-header', text: this.tb('slides') });
    const addSlideButton = createToolbarIconButton(sidebarHeader, {
      className: 'native-powerpoint-sidebar-add',
      icon: 'plus',
      label: this.t('powerpoint:accessibility.newSlide')
    });
    this.registerDomEvent(addSlideButton, 'click', () => void this.slideFilmstripController.addSlideWithLayout('blank'));
    this.thumbnailContainer = sidebar.createDiv({ cls: 'native-powerpoint-thumbnails' });

    const main = this.layoutEl.createDiv({ cls: 'native-powerpoint-main-content' });
    this.createToolbar(main);
    this.canvasPane = main.createDiv({ cls: PPTX_EDITOR_CHROME_DOCUMENT_SURFACE_CLASS });
    this.slideSurface = this.canvasPane.createDiv({ cls: PPTX_EDITOR_CHROME_DOCUMENT_CONTENT_CLASS });
    this.registerDomEvent(this.canvasPane, 'pointerdown', this.handleCanvasPanePointerDown, true);
    this.registerDomEvent(this.canvasPane, 'contextmenu', this.handleCanvasContextMenu);
    this.registerCanvasWheelZoom();
    this.observeCanvasPane();
    this.insertController.registerMenus();
    this.registerReplaceImageInput();
    this.textToolbarController.reset();
    this.register(() => this.textToolbarController.closeToolbarPopover());

    this.inspectorEl = this.layoutEl.createDiv({ cls: 'native-powerpoint-inspector' });
    this.applyInspectorVisibility();
    this.toolbarTooltips.attach(root);
  }

  private applyInspectorVisibility(): void {
    const show = this.getSettings().showInspector;
    this.inspectorEl?.toggleClass('native-powerpoint-inspector-hidden', !show);
  }

  refreshSettings(): void {
    this.applyThemeClass();
    this.applyInspectorVisibility();
    if (this.getSettings().showInspector) {
      this.renderInspector();
    }
  }

  private applyThemeClass(): void {
    if (!this.rootEl) return;

    this.rootEl.removeClasses([
      'native-powerpoint-theme-system',
      'native-powerpoint-theme-light',
      'native-powerpoint-theme-dark',
      'native-powerpoint-theme-resolved-light',
      'native-powerpoint-theme-resolved-dark'
    ]);
    const editorTheme = this.getSettings().editorTheme;
    this.rootEl.addClass(`native-powerpoint-theme-${editorTheme}`);
    this.rootEl.addClass(`native-powerpoint-theme-resolved-${this.getSettings().resolvedEditorTheme}`);
  }

  private createHeaderBar(root: HTMLElement): void {
    const headerBar = root.createDiv({ cls: PPTX_EDITOR_CHROME_HEADER_CLASS });

    const saveButton = createToolbarIconButton(headerBar, {
      className: 'native-powerpoint-header-save',
      icon: 'save',
      label: this.t('powerpoint:accessibility.save')
    });
    saveButton.addEventListener('click', () => void this.saveCurrentPresentation());

    const headerMain = headerBar.createDiv({ cls: 'native-powerpoint-headerbar-main' });
    this.createHeader(headerMain);
    this.createMenuBar(headerMain);
    this.createSaveStatus(headerBar);
  }

  private createHeader(root: HTMLElement): void {
    const header = root.createDiv({ cls: 'native-powerpoint-header' });

    const title = header.createDiv({ cls: 'native-powerpoint-header-title' });
    this.headerTitleEl = title.createSpan({
      cls: 'native-powerpoint-header-name',
      text: this.getDisplayText()
    });
    this.headerTitleEl.setAttribute('role', 'button');
    this.headerTitleEl.setAttribute('tabindex', '0');
    this.headerTitleEl.setAttribute('aria-label', pptT('powerpoint:accessibility.renamePresentation'));
    this.headerTitleEl.title = pptT('powerpoint:accessibility.clickToRename');
    this.headerTitleEl.addEventListener('click', () => this.beginRenameTitle());
    this.headerTitleEl.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        this.beginRenameTitle();
      }
    });
  }

  private createSaveStatus(root: HTMLElement): void {
    this.statusEl = root.createDiv({ cls: 'native-powerpoint-save-status', text: this.t('powerpoint:save.ready') });
    this.statusEl.setAttribute('role', 'status');
    this.statusEl.addEventListener('click', () => {
      if (this.session.saveState === 'failed') {
        void this.saveCurrentPresentation();
      }
    });
  }

  private updateHeaderTitle(): void {
    this.headerTitleEl?.setText(this.getDisplayText());
  }

  private beginRenameTitle(): void {
    const file = this.loadedFile || this.file;
    const titleEl = this.headerTitleEl;
    const parent = titleEl?.parentElement ?? null;
    if (!file || !titleEl || !parent) return;
    if (parent.querySelector('.native-powerpoint-header-name-input')) return;

    const input = parent.createEl('input', {
      cls: 'native-powerpoint-header-name-input',
      type: 'text',
      value: file.basename
    });
    parent.insertBefore(input, titleEl);
    titleEl.hide();
    input.focus();
    input.select();

    let finished = false;
    const cleanup = () => {
      input.remove();
      titleEl.show();
    };
    const commit = () => {
      if (finished) return;
      finished = true;
      const newName = input.value;
      cleanup();
      void this.renameLoadedFile(file, newName);
    };
    const cancel = () => {
      if (finished) return;
      finished = true;
      cleanup();
    };

    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        commit();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        cancel();
      }
    });
    input.addEventListener('blur', () => commit());
  }

  private async renameLoadedFile(file: TFile, rawName: string): Promise<void> {
    const sanitized = rawName.replace(/[\\/:*?"<>|]/g, '').trim();
    if (!sanitized || sanitized === file.basename) return;

    debugLog('file', 'Renaming PowerPoint presentation', {
      file: file.path,
      name: sanitized
    });
    try {
      await renameFileToSiblingName(this.app, file, `${sanitized}.${file.extension}`);
      debugLog('file', 'Renamed PowerPoint presentation', { file: file.path, name: sanitized });
    } catch (error) {
      pptNotice('powerpoint:notice.couldNotRename', { message: error instanceof Error ? error.message : String(error) });
    } finally {
      this.updateHeaderTitle();
    }
  }

  private createMenuBar(root: HTMLElement): void {
    this.menuBar.build(root, [
      { kind: 'dropdown', label: this.tb('file'), getItems: () => this.getFileMenuItems() },
      { kind: 'dropdown', label: this.tb('edit'), getItems: () => this.getEditMenuItems() },
      { kind: 'dropdown', label: this.tb('insert'), getItems: () => this.getInsertMenuItems() },
      { kind: 'action', label: this.tb('search'), action: () => this.findController.open() },
      { kind: 'action', label: this.tb('settings'), action: () => this.openPluginSettings() }
    ]);
  }

  private showMenuUnder(menu: Menu, anchor: HTMLElement): void {
    const rect = anchor.getBoundingClientRect();
    menu.showAtPosition({ x: rect.left, y: rect.bottom + 4 });
  }

  private createNativeMenu(): Menu {
    const instance = new Menu();
    const dom = (instance as unknown as { dom?: HTMLElement }).dom;
    dom?.addClass('native-powerpoint-light-surface');
    return instance;
  }

  private getFileMenuItems(): MenuDropdownEntry[] {
    return [
      { label: this.tb('save'), icon: 'save', onClick: () => void this.saveCurrentPresentation() },
      { label: this.tb('duplicate'), icon: 'copy', onClick: () => void this.duplicatePresentation() },
      'separator',
      { label: this.tb('print'), icon: 'printer', onClick: () => void this.printPresentation() },
      'separator',
      {
        label: this.tb('exportCurrentSlidePng'),
        icon: 'image',
        onClick: () => void this.exportController.exportCurrentSlideAsPng()
      },
      {
        label: this.tb('exportCurrentSlidePdf'),
        icon: 'file-output',
        onClick: () => void this.exportController.exportDeckAsPdf(true)
      },
      {
        label: this.tb('exportDeckPdf'),
        icon: 'file-output',
        onClick: () => void this.exportController.exportDeckAsPdf(false)
      },
      {
        label: this.tb('exportDeckPngsZip'),
        icon: 'file-archive',
        onClick: () => void this.exportController.exportDeckAsPngZip()
      },
      'separator',
      {
        label: this.tb('present'),
        icon: 'play',
        onClick: () => this.startPresentation()
      }
    ];
  }

  private getEditMenuItems(): MenuDropdownEntry[] {
    const canEdit = this.canEdit();
    const canUseHistory = canEdit && !this.historyController.isRestoring;
    const canUndo = this.hasInlineHistory('undo') || this.historyController.canUndo;
    const canRedo = this.hasInlineHistory('redo') || this.historyController.canRedo;
    const hasSelection = this.selectedShapeIndex !== null || this.selectedShapeIndices.size > 0;
    const hasClipboard = Boolean(this.objectClipboard);

    return [
      {
        label: this.tb('undo'),
        icon: 'undo',
        onClick: () => void this.requestHistoryAction('undo', 'menu'),
        disabled: !canUseHistory || !canUndo
      },
      {
        label: this.tb('redo'),
        icon: 'redo',
        onClick: () => void this.requestHistoryAction('redo', 'menu'),
        disabled: !canUseHistory || !canRedo
      },
      'separator',
      {
        label: this.tb('cut'),
        icon: 'scissors',
        onClick: () => void this.cutSelectedShape(),
        disabled: !canEdit || this.selectedShapeIndex === null
      },
      {
        label: this.tb('copy'),
        icon: 'copy',
        onClick: () => void this.copySelectedShape(),
        disabled: !hasSelection
      },
      {
        label: this.tb('paste'),
        icon: 'clipboard-paste',
        onClick: () => void this.pasteCopiedShape(),
        disabled: !canEdit || !hasClipboard
      },
      {
        label: this.tb('pasteWithoutFormatting'),
        icon: 'clipboard-type',
        onClick: () => void this.pasteWithoutFormatting(),
        disabled: !canEdit || (!this.activeEditor && !hasClipboard)
      },
      {
        label: this.tb('delete'),
        icon: 'trash-2',
        onClick: () => void this.deleteSelectedShape(),
        disabled: !canEdit || !hasSelection
      },
      'separator',
      {
        label: this.tb('selectAll'),
        icon: 'box-select',
        onClick: () => this.selectAllShapes()
      },
      'separator',
      {
        label: this.tb('findAndReplace'),
        icon: 'replace',
        onClick: () => this.findController.open({ replace: true })
      }
    ];
  }

  private getInsertMenuItems(): MenuDropdownEntry[] {
    return [
      { label: this.tb('imageFromVault'), icon: 'image', onClick: () => this.openVaultImagePicker() },
      { label: this.tb('uploadImage'), icon: 'upload', onClick: () => this.imageFileInput?.click() },
      'separator',
      { label: this.tb('textBox'), icon: 'type', onClick: () => void this.insertController.insertTextBox(true) },
      { label: this.tb('rectangle'), icon: 'square', onClick: () => void this.insertShape('rect') },
      { label: this.tb('ellipse'), icon: 'circle', onClick: () => void this.insertShape('ellipse') },
      { label: this.tb('line'), icon: 'minus', onClick: () => void this.insertShape('line') },
      { label: this.tb('arrow'), icon: 'move-right', onClick: () => void this.insertShape('rightArrow') },
      'separator',
      { label: this.tb('table'), icon: 'table', onClick: () => this.openTableSizePicker(this.insertTableButton) },
      { label: this.tb('chart'), icon: 'bar-chart-3', onClick: () => void this.insertChart() },
      'separator',
      { label: this.tb('bulletedList'), icon: 'list', onClick: () => void this.applyListStyle('bullet') },
      {
        label: this.tb('numberedList'),
        icon: 'list-ordered',
        onClick: () => void this.applyListStyle('number')
      },
      'separator',
      { label: this.tb('newSlide'), icon: 'plus', onClick: () => void this.slideFilmstripController.addSlideWithLayout('blank') }
    ];
  }

  private async duplicatePresentation(): Promise<void> {
    const file = this.loadedFile || this.file;
    if (!file) {
      pptNotice('powerpoint:notice.openToDuplicate');
      return;
    }

    debugLog('file', 'Duplicating PowerPoint presentation', { file: file.path });
    try {
      if (this.session.dirty && this.canEdit()) {
        await this.saveCurrentPresentation();
      }
      const copyPath = this.getAvailableCopyPath(file);
      const data = await this.app.vault.readBinary(file);
      const created = await this.app.vault.createBinary(copyPath, data);
      debugLog('file', 'Duplicated PowerPoint presentation', { file: file.path, copyPath: created.path });
      pptNotice('powerpoint:notice.duplicatedTo', { fileName: created.name });
    } catch (error) {
      pptNotice('powerpoint:notice.couldNotDuplicatePresentation', { message: cleanError(error) });
    }
  }

  private getAvailableCopyPath(file: TFile): string {
    const folderPath = file.parent?.path;
    const folderPrefix = folderPath && folderPath !== '/' ? `${folderPath}/` : '';
    for (let index = 1; index < 1000; index += 1) {
      const suffix = index === 1 ? 'copy' : `copy ${index}`;
      const candidate = normalizePath(`${folderPrefix}${file.basename} ${suffix}.${file.extension}`);
      if (!this.app.vault.getAbstractFileByPath(candidate)) return candidate;
    }
    return normalizePath(`${folderPrefix}${file.basename} copy ${Date.now()}.${file.extension}`);
  }

  private async printPresentation(): Promise<void> {
    if (!this.engine || this.engine.slideCount === 0) {
      pptNotice('powerpoint:notice.openToPrint');
      return;
    }

    debugLog('export', 'Printing PowerPoint presentation', { slideCount: this.engine.slideCount });
    try {
      const indices = Array.from({ length: this.engine.slideCount }, (_, index) => index);
      const elements = this.collectExportSvgElements(indices);
      if (elements.length === 0) {
        throw new Error('No slides could be rendered for printing.');
      }

      pptNotice('powerpoint:notice.preparingPrint');
      const urls: string[] = [];
      for (const element of elements) {
        const bytes = await exportSlideToPng(element, this.contentEl.ownerDocument);
        urls.push(URL.createObjectURL(new Blob([bytes], { type: 'image/png' })));
      }
      this.printSlideImages(urls);
      debugLog('export', 'Opened PowerPoint print dialog', { slideCount: elements.length });
    } catch (error) {
      pptNotice('powerpoint:notice.couldNotPrint', { message: cleanError(error) });
    }
  }

  private printSlideImages(urls: string[]): void {
    const iframe = activeDocument.body.createEl('iframe', { cls: 'native-powerpoint-print-frame' });
    let cleaned = false;
    const cleanup = () => {
      if (cleaned) return;
      cleaned = true;
      urls.forEach((url) => URL.revokeObjectURL(url));
      iframe.remove();
    };

    const doc = iframe.contentDocument;
    const win = iframe.contentWindow;
    if (!doc || !win) {
      cleanup();
      pptNotice('powerpoint:notice.couldNotOpenPrintView');
      return;
    }

    // Print iframe needs ephemeral @page CSS; styles.css cannot target this document.
    const printSheet = new CSSStyleSheet();
    printSheet.replaceSync(
      '@page { size: landscape; margin: 12mm; }' +
      'html, body { margin: 0; padding: 0; background: #ffffff; }' +
      '.native-powerpoint-print-slide { page-break-after: always; text-align: center; }' +
      '.native-powerpoint-print-slide:last-child { page-break-after: auto; }' +
      '.native-powerpoint-print-slide img { width: 100%; height: auto; display: block; }',
    );
    doc.adoptedStyleSheets = [...doc.adoptedStyleSheets, printSheet];

    let remaining = urls.length;
    const onReady = () => {
      remaining -= 1;
      if (remaining > 0) return;
      win.focus();
      win.print();
    };

    win.addEventListener('afterprint', cleanup, { once: true });
    window.setTimeout(cleanup, 60000);

    for (const url of urls) {
      const wrap = doc.createDiv();
      wrap.className = 'native-powerpoint-print-slide';
      const img = doc.createEl('img');
      img.addEventListener('load', onReady, { once: true });
      img.addEventListener('error', onReady, { once: true });
      img.src = url;
      wrap.appendChild(img);
      doc.body.appendChild(wrap);
    }
  }

  private openPluginSettings(): void {
    const setting = (
      this.app as unknown as {
        setting?: { open?: () => void; openTabById?: (id: string) => void };
      }
    ).setting;
    if (!setting?.open || !setting.openTabById) {
      pptNotice('powerpoint:notice.settingsUnavailable');
      return;
    }
    setting.open();
    setting.openTabById('native-powerpoint-doc-editor');
  }

  private updateZoomLabel(): void {
    this.zoomLevelEl?.setText(this.tb('zoomPercent', { percent: Math.round(this.zoomLevel * 100) }));
  }

  private observeCanvasPane(): void {
    this.canvasResizeObserver?.disconnect();
    this.canvasResizeObserver = null;

    if (!this.canvasPane || typeof ResizeObserver === 'undefined') return;

    const observer = new ResizeObserver(() => this.updateSlideScale());
    observer.observe(this.canvasPane);
    this.canvasResizeObserver = observer;
    this.register(() => observer.disconnect());
  }

  private registerCanvasWheelZoom(): void {
    if (!this.canvasPane) return;

    const pane = this.canvasPane;
    const handleWheel = (event: WheelEvent) => this.handleCanvasWheel(event);
    pane.addEventListener('wheel', handleWheel, { passive: false });
    this.register(() => pane.removeEventListener('wheel', handleWheel));
  }

  private createToolbar(main: HTMLElement): void {
    this.editButtons = [];
    const toolbar = main.createDiv({ cls: PPTX_EDITOR_CHROME_TOOLBAR_CLASS });

    // Layout mirrors Google Slides' toolbar order: history first, then zoom,
    // then slide operations, then insert/object operations, then find. Slide
    // navigation sits on the right like Slides' top-right controls.
    const historyGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    const undoLabel = this.t('powerpoint:accessibility.undoShortcut');
    const redoLabel = this.t('powerpoint:accessibility.redoShortcut');
    const undoButton = this.createIconButton(historyGroup, 'undo', undoLabel, () => void this.requestHistoryAction('undo', 'toolbar'));
    const redoButton = this.createIconButton(historyGroup, 'redo', redoLabel, () => void this.requestHistoryAction('redo', 'toolbar'));
    this.historyController.attachButtons(undoButton, redoButton);

    const zoomGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.createIconButton(zoomGroup, 'zoom-out', this.tb('zoomOut'), () => this.setZoom(this.zoomLevel - 0.1));
    this.zoomLevelEl = zoomGroup.createDiv({ cls: 'native-powerpoint-zoom-level', text: this.tb('zoomPercent', { percent: 100 }) });
    this.createIconButton(zoomGroup, 'zoom-in', this.tb('zoomIn'), () => this.setZoom(this.zoomLevel + 0.1));
    this.updateZoomLabel();

    this.insertController.createToolbarGroup(toolbar);

    const slideGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    // Primary click adds a blank slide immediately (Google Slides "+" behavior).
    this.createEditIconButton(slideGroup, 'plus', this.tb('newSlide'), () => void this.slideFilmstripController.addSlideWithLayout('blank'));
    // A caret opens the layout choices without blocking the quick-add action.
    const newSlideLayoutButton = this.createEditIconButton(slideGroup, 'chevron-down', this.tb('newSlideLayout'), () => {
      this.toggleInsertMenu(newSlideLayoutButton, [
        { label: this.tb('blank'), onClick: () => void this.slideFilmstripController.addSlideWithLayout('blank') },
        { label: this.tb('title'), onClick: () => void this.slideFilmstripController.addSlideWithLayout('title') },
        { label: this.tb('titleBody'), onClick: () => void this.slideFilmstripController.addSlideWithLayout('titleBody') }
      ]);
    });
    this.createEditIconButton(slideGroup, 'files', this.tb('duplicateSlide'), () => void this.slideFilmstripController.duplicateSlide());
    this.createEditIconButton(slideGroup, 'trash-2', this.tb('deleteSlide'), () => void this.slideFilmstripController.deleteSlide());
    this.createEditIconButton(slideGroup, 'arrow-left-to-line', this.tb('moveSlideLeft'), () => void this.slideFilmstripController.moveSlide(-1));
    this.createEditIconButton(slideGroup, 'arrow-right-to-line', this.tb('moveSlideRight'), () => void this.slideFilmstripController.moveSlide(1));

    const objectGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.copyButton = this.createIconButton(objectGroup, 'copy', this.tb('copySelectedObject'), () => void this.copySelectedShape());
    this.pasteButton = this.createIconButton(objectGroup, 'clipboard-paste', this.tb('pasteObject'), () => void this.pasteCopiedShape());
    this.duplicateButton = this.createIconButton(objectGroup, 'copy-plus', this.tb('duplicateSelectedObject'), () => void this.duplicateSelectedShape());
    this.createEditIconButton(objectGroup, 'eraser', this.tb('deleteSelectedObject'), () => void this.deleteSelectedShape());

    const fontGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    const fontFamilyLabel = this.t('powerpoint:accessibility.fontFamily');
    const fontButton = fontGroup.createEl('button', {
      cls: 'native-powerpoint-toolbar-btn native-powerpoint-text-toolbar-font',
      attr: { 'aria-label': fontFamilyLabel, 'data-tooltip': fontFamilyLabel }
    });
    const fontLabel = fontButton.createSpan({ cls: 'native-powerpoint-text-toolbar-font-label', text: this.tb('font') });
    this.topFontLabel = fontLabel;
    setIcon(fontButton.createSpan({ cls: 'native-powerpoint-text-toolbar-caret' }), 'chevron-down');
    this.textToolbarController.wireFontFamilyButton(fontButton, fontLabel);

    this.arrangeController.createToolbarGroups(toolbar);

    const searchGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    const findButton = this.createIconButton(searchGroup, 'search', this.tb('findInPresentation'), () => this.findController.toggle());
    // The find/replace UI is a floating dropdown anchored to the search button
    // rather than an inline element inside the horizontally scrolling toolbar.
    this.findController.createPanel(findButton);

    const shareGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.createIconButton(shareGroup, 'play', this.tb('present'), () => this.startPresentation());
    const exportButton = this.createIconButton(shareGroup, 'download', this.tb('exportSlides'), () =>
      this.exportController.openMenu(exportButton)
    );

    const navGroup = toolbar.createDiv({
      cls: 'native-powerpoint-toolbar-group native-powerpoint-toolbar-group-end'
    });
    this.createIconButton(navGroup, 'chevron-left', this.tb('previousSlide'), () => this.slideFilmstripController.navigateToSlide(this.currentSlide - 1, 'toolbar-prev'));
    this.slideCounterEl = navGroup.createDiv({ cls: 'native-powerpoint-page-counter', text: this.t('powerpoint:present.slideCount', { current: 0, total: 0 }) });
    this.createIconButton(navGroup, 'chevron-right', this.tb('nextSlide'), () => this.slideFilmstripController.navigateToSlide(this.currentSlide + 1, 'toolbar-next'));

    this.updateEditingAvailability();
    this.historyController.updateAvailability();
    this.updateObjectClipboardAvailability();
  }

  private registerReplaceImageInput(): void {
    if (!this.layoutEl) return;
    const replaceInput = this.layoutEl.createEl('input', {
      type: 'file',
      cls: 'native-powerpoint-image-file-input'
    });
    replaceInput.accept = 'image/png,image/jpeg,image/gif,image/webp,image/bmp';
    replaceInput.addEventListener('change', () => {
      const file = replaceInput.files?.[0];
      replaceInput.value = '';
      const shapeIndex = this.pendingReplaceShapeIndex;
      this.pendingReplaceShapeIndex = null;
      if (file && shapeIndex !== null) void this.replaceImageWithLocalFile(shapeIndex, file);
    });
    this.replaceImageFileInput = replaceInput;
  }





  private getSelectedIndices(): number[] {
    if (this.selectedShapeIndices.size > 0) return [...this.selectedShapeIndices];
    if (this.selectedShapeIndex !== null) return [this.selectedShapeIndex];
    return [];
  }

  /** Text inside an existing multi-selection is a drag surface, not an edit trigger. */
  private shouldStartGroupDragFromText(shapeIndex: number | null, additive: boolean): boolean {
    return !additive
      && shapeIndex !== null
      && this.selectedShapeIndices.size > 1
      && this.selectedShapeIndices.has(shapeIndex);
  }

  private toggleShapeInSelection(shapeIndex: number): void {
    if (!isSelectableShapeIndex(shapeIndex)) return;
    const next = new Set(this.getSelectedIndices());
    if (next.has(shapeIndex)) {
      next.delete(shapeIndex);
    } else {
      next.add(shapeIndex);
    }
    this.applyMultiSelection([...next]);
  }











  private updateArrangeAvailability(): void {
    this.arrangeController.updateArrangeAvailability();
  }

  private async nudgeSelection(key: string, large: boolean): Promise<void> {
    await this.arrangeController.nudgeSelection(key, large);
  }

  private emuPointToPane(emuX: number, emuY: number): PointerPoint | null {
    if (!this.engine || !this.svgEl || !this.canvasPane) return null;
    const ctm = this.svgEl.getScreenCTM();
    const scale = this.engine.getSlideScale(this.svgEl);
    if (!ctm || !scale) return null;

    const screenX = (emuX / scale) * ctm.a + ctm.e;
    const screenY = (emuY / scale) * ctm.d + ctm.f;
    const paneRect = this.canvasPane.getBoundingClientRect();
    return {
      x: screenX - paneRect.left + this.canvasPane.scrollLeft,
      y: screenY - paneRect.top + this.canvasPane.scrollTop
    };
  }

  /** Pane pixels per EMU at the current zoom; used during drag to avoid repeated layout reads. */
  private getPaneEmuScale(): { x: number; y: number } | null {
    if (!this.engine || !this.svgEl) return null;
    const ctm = this.svgEl.getScreenCTM();
    const scale = this.engine.getSlideScale(this.svgEl);
    if (!ctm || !scale || ctm.a === 0 || ctm.d === 0) return null;
    return { x: ctm.a / scale, y: ctm.d / scale };
  }

  /**
   * Move the selection overlay while dragging without resizing it. Uses the
   * shape's fitted start box (as computed on selection) and only translates it
   * by the snapped move delta, so an image keeps its tight outline instead of
   * jumping to the square OOXML frame that {@link positionOverlayFromTransform}
   * would produce.
   */
  private positionOverlayDuringMove(transform: ShapeTransform): void {
    if (!this.selectionOverlay || !this.dragState) return;

    const shape = this.getSelectedShapeElement();
    if (this.isPictureShape(shape) && shape) {
      const box = this.getPictureSelectionBox(shape);
      if (box) {
        this.selectionOverlay.style.removeProperty('transform');
        this.selectionOverlay.style.removeProperty('transform-origin');
        this.selectionOverlay.setCssProps({
          left: `${box.left}px`,
          top: `${box.top}px`,
          width: `${box.width}px`,
          height: `${box.height}px`,
        });
        return;
      }
    }

    const box = this.dragState.startBox;
    const paneScaleX = this.dragState.paneEmuScaleX;
    const paneScaleY = this.dragState.paneEmuScaleY;
    if (paneScaleX && paneScaleY) {
      const dEmuX = transform.x - this.dragState.startTransform.x;
      const dEmuY = transform.y - this.dragState.startTransform.y;
      this.selectionOverlay.setCssProps({
        left: `${box.left + dEmuX * paneScaleX}px`,
        top: `${box.top + dEmuY * paneScaleY}px`,
        width: `${box.width}px`,
        height: `${box.height}px`
      });
      this.syncOrientedSelectionOverlayRotation(transform);
      return;
    }

    const startPane = this.emuPointToPane(this.dragState.startTransform.x, this.dragState.startTransform.y);
    const nextPane = this.emuPointToPane(transform.x, transform.y);
    if (!startPane || !nextPane) {
      this.positionOverlayFromTransform(transform);
      this.syncOrientedSelectionOverlayRotation(transform);
      return;
    }

    this.selectionOverlay.setCssProps({
      left: `${box.left + (nextPane.x - startPane.x)}px`,
      top: `${box.top + (nextPane.y - startPane.y)}px`,
      width: `${box.width}px`,
      height: `${box.height}px`
    });
    this.syncOrientedSelectionOverlayRotation(transform);
  }

  private positionOverlayFromTransform(transform: ShapeTransform): void {
    if (!this.selectionOverlay) return;
    const box = this.getTransformSelectionBox(transform);
    if (!box) return;
    this.selectionOverlay.setCssProps({
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: `${box.width}px`,
      height: `${box.height}px`
    });
  }

  private shapeHasRotation(transform: ShapeTransform): boolean {
    if (!this.engine) return false;
    return Math.abs(this.engine.ooxmlToDegrees(transform.rot)) > 0.001;
  }

  private isPictureShape(shape: Element | null): boolean {
    return shape?.getAttribute('data-ooxml-shape-type') === 'picture';
  }

  /** Rotated pictures letterbox inside the OOXML frame; fit the visible image bounds. */
  private pictureUsesImageSelectionBounds(shape: Element | null): boolean {
    return this.isPictureShape(shape);
  }

  private getPictureImageElement(shape: Element | null): SVGImageElement | null {
    if (!shape || !this.isPictureShape(shape)) return null;
    const image = shape.querySelector('image');
    return image instanceof SVGImageElement ? image : null;
  }

  private getPictureSelectionBox(
    shape: SVGGElement,
  ): { left: number; top: number; width: number; height: number } | null {
    const image = this.getPictureImageElement(shape);
    return image ? this.getElementBox(image) : null;
  }

  private getShapeSelectionElement(shape: SVGGElement): Element {
    return this.getPictureImageElement(shape) ?? shape;
  }

  private logSelectionOverlayLayout(
    reason: string,
    shape: SVGGElement | null,
    transform: ShapeTransform | null,
    box: { left: number; top: number; width: number; height: number } | null,
    strategy: string,
  ): void {
    if (!shape) return;
    let skippedDragLogs = 0;
    if (reason === 'drag') {
      const now = performance.now();
      if (now - this.lastSelectionOverlayDragLogAt < 250) {
        this.skippedSelectionOverlayDragLogs += 1;
        return;
      }
      this.lastSelectionOverlayDragLogAt = now;
      skippedDragLogs = this.skippedSelectionOverlayDragLogs;
      this.skippedSelectionOverlayDragLogs = 0;
    }

    const groupBox = this.getElementBox(shape);
    const image = this.getPictureImageElement(shape);
    const imageBox = image ? this.getElementBox(image) : null;
    const ooxmlBox = transform ? this.getTransformSelectionBox(transform) : null;
    debugLog('selection', 'PowerPoint selection overlay layout', {
      reason,
      slide: this.currentSlide,
      shapeIndex: this.selectedShapeIndex,
      strategy,
      rotation: transform?.rot ?? 0,
      groupBox,
      imageBox,
      ooxmlBox,
      box,
      imageTransform: image?.getAttribute('transform') ?? null,
      skippedDragLogs,
    });
  }

  private normalizeDegrees(degrees: number): number {
    return ((degrees % 360) + 360) % 360;
  }

  /** Snap a nearly horizontal or vertical shape to its exact cardinal angle. */
  private getCardinalRotationSnap(degrees: number): number | null {
    const target = Math.round(degrees / 90) * 90;
    return Math.abs(degrees - target) <= ROTATION_SNAP_THRESHOLD_DEGREES ? target : null;
  }

  /**
   * Keep a drag's rotation continuous as atan2 wraps from +180° to -180°.
   * The returned value is deliberately unbounded so the selection outline can
   * make full turns in either direction before the OOXML angle is normalized.
   */
  private advanceContinuousRotation(
    rotation: {
      startAngle?: number;
      lastAngle?: number;
      accumulatedRotationDegrees?: number;
    },
    angle: number,
  ): number {
    const previousAngle = rotation.lastAngle ?? rotation.startAngle ?? angle;
    const deltaRadians = Math.atan2(
      Math.sin(angle - previousAngle),
      Math.cos(angle - previousAngle),
    );
    rotation.lastAngle = angle;
    rotation.accumulatedRotationDegrees = (rotation.accumulatedRotationDegrees ?? 0)
      + (deltaRadians * 180) / Math.PI;
    return rotation.accumulatedRotationDegrees;
  }

  private solveRotatedRectDimensions(
    desiredWidth: number,
    desiredHeight: number,
    rotationDegrees: number,
    fallbackWidth: number,
    fallbackHeight: number,
    scaleX: number,
    scaleY: number,
  ): { width: number; height: number } {
    const radians = (rotationDegrees * Math.PI) / 180;
    const cos = Math.abs(Math.cos(radians));
    const sin = Math.abs(Math.sin(radians));
    const determinant = cos * cos - sin * sin;

    if (Math.abs(determinant) > 0.000001) {
      const width = (cos * desiredWidth - sin * desiredHeight) / determinant;
      const height = (-sin * desiredWidth + cos * desiredHeight) / determinant;
      if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
        return { width, height };
      }
    }

    const normalized = this.normalizeDegrees(rotationDegrees);
    const swapAxes = (normalized > 45 && normalized <= 135) || (normalized > 225 && normalized <= 315);
    return {
      width: fallbackWidth * (swapAxes ? scaleY : scaleX),
      height: fallbackHeight * (swapAxes ? scaleX : scaleY),
    };
  }

  private parseSvgRotate(transform: string | null | undefined): {
    raw: string;
    degrees: number;
    centerX: number | null;
    centerY: number | null;
  } | null {
    if (!transform) return null;
    const match = transform.match(/\brotate\(\s*([-+]?\d*\.?\d+(?:e[-+]?\d+)?)(?:[\s,]+([-+]?\d*\.?\d+(?:e[-+]?\d+)?)[\s,]+([-+]?\d*\.?\d+(?:e[-+]?\d+)?))?\s*\)/i);
    if (!match) return null;
    const degrees = Number(match[1]);
    const centerX = match[2] === undefined ? null : Number(match[2]);
    const centerY = match[3] === undefined ? null : Number(match[3]);
    return Number.isFinite(degrees)
      ? {
          raw: match[0],
          degrees,
          centerX: Number.isFinite(centerX) ? centerX : null,
          centerY: Number.isFinite(centerY) ? centerY : null,
        }
      : null;
  }

  private getTransformSelectionBox(
    transform: ShapeTransform
  ): { left: number; top: number; width: number; height: number } | null {
    const topLeft = this.emuPointToPane(transform.x, transform.y);
    const bottomRight = this.emuPointToPane(transform.x + transform.cx, transform.y + transform.cy);
    if (!topLeft || !bottomRight) return null;
    return {
      left: topLeft.x,
      top: topLeft.y,
      width: Math.max(0, bottomRight.x - topLeft.x),
      height: Math.max(0, bottomRight.y - topLeft.y)
    };
  }

  private syncOrientedSelectionOverlayRotation(transform: ShapeTransform): void {
    if (!this.selectionOverlay || !this.engine) return;
    const shape = this.getSelectedShapeElement();
    if (this.isPictureShape(shape)) {
      this.selectionOverlay.style.removeProperty('transform');
      this.selectionOverlay.style.removeProperty('transform-origin');
      return;
    }
    if (!this.shapeHasRotation(transform)) {
      this.selectionOverlay.style.removeProperty('transform');
      this.selectionOverlay.style.removeProperty('transform-origin');
      return;
    }
    const degrees = this.engine.ooxmlToDegrees(transform.rot);
		this.selectionOverlay.setCssProps({ transform: `rotate(${degrees}deg)` });
  }

  /** OOXML frame + CSS rotation for rotated autoshapes; tight image bounds for pictures. */
  private applySelectionOverlayLayout(transform?: ShapeTransform): boolean {
    if (!this.selectionOverlay || !this.engine) return false;
    const next = transform ?? this.selectedTransform;
    if (!next) return false;

    const selected = this.getSelectedShapeElement();
    if (this.pictureUsesImageSelectionBounds(selected) && selected) {
      const box = this.getPictureSelectionBox(selected);
      if (!box) return false;
      this.selectionOverlay.style.removeProperty('transform');
      this.selectionOverlay.style.removeProperty('transform-origin');
      this.selectionOverlay.setCssProps({
        left: `${box.left}px`,
        top: `${box.top}px`,
        width: `${box.width}px`,
        height: `${box.height}px`,
      });
      this.logSelectionOverlayLayout('apply', selected, next, box, 'picture-image-bounds');
      return true;
    }

    if (this.shapeHasRotation(next)) {
      this.positionOverlayFromTransform(next);
      this.syncOrientedSelectionOverlayRotation(next);
      this.logSelectionOverlayLayout('apply', selected, next, null, 'ooxml-frame-rotate');
      return true;
    }

    const target = selected ? this.getShapeSelectionElement(selected) : null;
    const box = target ? this.getElementBox(target) : null;
    if (!box) return false;
    this.selectionOverlay.style.removeProperty('transform');
    this.selectionOverlay.style.removeProperty('transform-origin');
    this.selectionOverlay.setCssProps({
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: `${box.width}px`,
      height: `${box.height}px`
    });
    this.logSelectionOverlayLayout('apply', selected, next, box, 'shape-bounds');
    return true;
  }

  private startRotateDrag(event: PointerEvent): void {
    if (!this.engine || this.selectedTransform === null || !this.selectionOverlay) return;
    if (!this.ensureEditable('rotate object')) return;

    const rect = this.selectionOverlay.getBoundingClientRect();
    const centerClientX = rect.left + rect.width / 2;
    const centerClientY = rect.top + rect.height / 2;
    const startBox = this.getSelectedBox();
    if (!startBox) return;
    const previewElement = this.getSelectedShapeElement();
    if (previewElement) {
      previewElement.classList.add('native-powerpoint-shape-drag-preview');
    }

    this.dragState = {
      mode: 'rotate',
      pointerId: event.pointerId,
      startPoint: { x: event.clientX, y: event.clientY },
      startClientX: event.clientX,
      startClientY: event.clientY,
      startBox,
      startTransform: cloneTransform(this.selectedTransform),
      latestTransform: cloneTransform(this.selectedTransform),
      centerClientX,
      centerClientY,
      startAngle: Math.atan2(event.clientY - centerClientY, event.clientX - centerClientX),
      lastAngle: Math.atan2(event.clientY - centerClientY, event.clientX - centerClientX),
      accumulatedRotationDegrees: 0,
      previewElement,
      previewOriginalTransform: previewElement?.getAttribute('transform') ?? null,
    };
    logPptxAction('selection', 'rotate', {
      slide: this.currentSlide,
      shapeIndexes: this.selectedShapeIndex === null ? [] : [this.selectedShapeIndex],
      startTransform: cloneTransform(this.selectedTransform),
      preview: 'shape-group',
      rotationSnapTarget: null,
    });
  }

  private toggleInsertMenu(
    anchor: HTMLButtonElement,
    items: { label: string; onClick: () => void }[]
  ): void {
    if (!anchor.dataset.menuId) {
      anchor.dataset.menuId = `insert-menu-${Math.random().toString(36).slice(2)}`;
    }

    // Only treat a repeat click as "toggle closed" when a menu is actually open
    // for this anchor. Without the explicit null check, a first click compares
    // two `undefined`s (no open menu, no anchor id yet) and wrongly closes.
    if (this.activeInsertMenu && this.activeInsertMenu.dataset.anchorId === anchor.dataset.menuId) {
      this.closeInsertMenus();
      return;
    }

    this.closeInsertMenus();
    anchor.classList.add('native-powerpoint-insert-menu-anchor');

    const menu = createPopoverShell(activeDocument.body, {
      className: 'native-powerpoint-insert-menu native-powerpoint-light-surface'
    });
    menu.dataset.anchorId = anchor.dataset.menuId;
    for (const item of items) {
      createMenuItem(menu, {
        className: 'native-powerpoint-insert-menu-item',
        text: item.label,
        preventDefaultOnClick: true,
        stopClickPropagation: true,
        onClick: () => {
        this.closeInsertMenus();
        item.onClick();
        }
      });
    }

    positionPopoverBelow(menu, anchor);
    this.activeInsertMenu = menu;
  }

  private closeInsertMenus(): void {
    this.activeInsertMenu?.remove();
    this.activeInsertMenu = null;
  }

  private openVaultImagePicker(): void {
    new VaultImageSuggestModal(this.app, (file) => void this.insertImageFromVaultFile(file), this.t).open();
  }

  private openInsertTableModal(): void {
    if (!this.ensureEditable('insert table')) return;
    new InsertTableModal(this.app, (rows, cols) => void this.insertTable(rows, cols), this.t).open();
  }

  // Google Slides-style size picker: a hover grid that matches the look of the
  // other toolbar popovers (color, font) instead of a separate modal dialog.
  private openTableSizePicker(anchor: HTMLElement | null): void {
    if (!this.ensureEditable('insert table')) return;
    if (!anchor) {
      this.openInsertTableModal();
      return;
    }

    const cols = 10;
    const rows = 8;
    this.openToolbarPopover(anchor, (popover) => {
      popover.addClass('native-powerpoint-table-picker');

      const grid = popover.createDiv({ cls: 'native-powerpoint-table-picker-grid' });
      const label = popover.createDiv({
        cls: 'native-powerpoint-table-picker-label',
        text: this.tb('tableSizePicker')
      });

      const cells: HTMLButtonElement[] = [];
      const highlight = (activeCols: number, activeRows: number): void => {
        cells.forEach((cell, index) => {
          const c = index % cols;
          const r = Math.floor(index / cols);
          cell.toggleClass('is-active', c < activeCols && r < activeRows);
        });
        label.setText(activeCols > 0 && activeRows > 0
          ? this.tb('tableSize', { columns: activeCols, rows: activeRows })
          : this.tb('tableSizePicker'));
      };

      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cell = grid.createEl('button', {
            cls: 'native-powerpoint-table-picker-cell',
            attr: { 'aria-label': this.t('powerpoint:accessibility.tableCellSize', { columns: c + 1, rows: r + 1 }) }
          });
          cell.addEventListener('pointerenter', () => highlight(c + 1, r + 1));
          this.bindToolbarButton(cell, () => {
            this.closeToolbarPopover();
            void this.insertTable(r + 1, c + 1);
          });
          cells.push(cell);
        }
      }

      grid.addEventListener('pointerleave', () => highlight(0, 0));
    });
  }

  private async insertImageFromVaultFile(file: TFile): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert image')) return;

    try {
      const bytes = await this.app.vault.readBinary(file);
      const history = await this.captureHistoryEntry('Insert image');
      const shapeIndex = await this.engine.addImage(
        this.currentSlide,
        new Uint8Array(bytes),
        getImageMimeType(file.extension)
      );
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint image from vault', {
        slide: this.currentSlide,
        shapeIndex,
        file: file.path,
        bytes: bytes.byteLength
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint vault image insertion failed', { file: file.path, error });
      pptNotice('powerpoint:notice.couldNotInsertImage', { message: cleanError(error) });
    }
  }

  private async insertImageFromLocalFile(file: File): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert image')) return;

    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const history = await this.captureHistoryEntry('Insert image');
      const shapeIndex = await this.engine.addImage(
        this.currentSlide,
        bytes,
        file.type || getImageMimeType(file.name.split('.').pop() ?? 'png')
      );
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint image from local file', {
        slide: this.currentSlide,
        shapeIndex,
        fileName: file.name,
        mimeType: file.type || null,
        bytes: bytes.byteLength
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint local image insertion failed', {
        fileName: file.name,
        mimeType: file.type || null,
        error
      });
      pptNotice('powerpoint:notice.couldNotInsertImage', { message: cleanError(error) });
    }
  }

  private async insertShape(geometry: InsertableShapeGeometry): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert shape')) return;

    try {
      const history = await this.captureHistoryEntry('Insert shape');
      const shapeIndex = await this.engine.addShapeGeometry(this.currentSlide, geometry);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint shape', {
        slide: this.currentSlide,
        shapeIndex,
        geometry
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint shape insertion failed', { geometry, error });
      pptNotice('powerpoint:notice.couldNotInsertShape', { message: cleanError(error) });
    }
  }

  private async insertTable(rows: number, cols: number): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert table')) return;

    try {
      const history = await this.captureHistoryEntry('Insert table');
      const shapeIndex = await this.engine.addTable(this.currentSlide, rows, cols);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint table', {
        slide: this.currentSlide,
        shapeIndex,
        rows,
        columns: cols
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint table insertion failed', { rows, columns: cols, error });
      pptNotice('powerpoint:notice.couldNotInsertTable', { message: cleanError(error) });
    }
  }

  private async insertChart(): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert chart')) return;

    try {
      const history = await this.captureHistoryEntry('Insert chart');
      const shapeIndex = await this.engine.addChart(this.currentSlide);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint chart', {
        slide: this.currentSlide,
        shapeIndex
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint chart insertion failed', { slide: this.currentSlide, error });
      pptNotice('powerpoint:notice.couldNotInsertChart', { message: cleanError(error) });
    }
  }

  private async applyListStyle(style: ParagraphListStyle): Promise<void> {
    if (!this.engine || !this.ensureEditable('format text')) return;

    if (this.textCommitPromise) {
      await this.finishInlineTextEditing('before-list-formatting');
    }

    const textTarget = this.getTextEditTarget(this.activeEditorTarget);
    const shapeIndex = textTarget?.shapeIndex ?? this.selectedShapeIndex;
    if (shapeIndex === null) {
      pptNotice('powerpoint:notice.selectTextBoxFirst');
      return;
    }

    const paragraphIndex = textTarget?.kind === 'shape-paragraph' ? textTarget.paragraphIndex : 0;
    const editor = this.activeEditor;
    const shapeTextTarget = this.activeShapeTextTarget;
    let pendingText: string | null = null;
    let savedStart = 0;
    let savedEnd = 0;
    if (
      editor
      && shapeTextTarget
      && shapeTextTarget.shapeIndex === shapeIndex
      && shapeTextTarget.paragraphIndex === paragraphIndex
    ) {
      savedStart = Math.min(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      savedEnd = Math.max(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      if (this.activeEditorTextDirty) {
        const normalizedEditorText = this.paragraphEditorTextFromDom(
          shapeIndex,
          paragraphIndex,
          editor.value
        );
        pendingText = normalizedEditorText !== shapeTextTarget.text ? normalizedEditorText : null;
      }
    }

    try {
      const history = await this.captureHistoryEntry(
        style === 'bullet' ? 'Bulleted list' : style === 'number' ? 'Numbered list' : 'Remove list'
      );
      const scrollPosition = this.captureCanvasScroll();
      if (pendingText !== null) {
        await this.engine.updateParagraphText(this.currentSlide, shapeIndex, paragraphIndex, pendingText);
        this.activeEditorTextDirty = false;
      }
      await this.engine.applyListStyle(this.currentSlide, shapeIndex, paragraphIndex, style);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderEditedShape(shapeIndex);
      if (rendered) {
        this.restoreCanvasScrollSoon(scrollPosition);
        await this.renderThumbnails();
        if (editor && this.activeEditor === editor && shapeTextTarget?.shapeIndex === shapeIndex) {
          if (this.refreshActiveShapeEditorAfterRender()) {
            const length = editor.value.length;
            editor.setSelectionRange(Math.min(savedStart, length), Math.min(savedEnd, length));
            this.refreshInlineEditorGeometry();
          } else {
            this.removeActiveEditor(editor);
          }
        }
        this.updateTextToolbar();
      }
      debugLog('insert', 'Applied PowerPoint list style', {
        slide: this.currentSlide,
        shapeIndex,
        paragraphIndex,
        style
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint list-style update failed', {
        slide: this.currentSlide,
        shapeIndex,
        paragraphIndex,
        style,
        error
      });
      pptNotice('powerpoint:notice.couldNotUpdateListStyle', { message: cleanError(error) });
    }
  }

  private createIconButton(container: HTMLElement, icon: string, label: string, onClick: () => void): HTMLButtonElement {
    return createToolbarIconButton(container, {
      className: 'native-powerpoint-toolbar-btn',
      icon,
      label,
      onClick
    });
  }

  private createEditIconButton(container: HTMLElement, icon: string, label: string, onClick: () => void): HTMLButtonElement {
    const button = this.createIconButton(container, icon, label, () => {
      if (this.ensureEditable(label.toLowerCase())) {
        onClick();
      }
    });
    button.dataset.baseTitle = label;
    this.editButtons.push(button);
    return button;
  }

  private buildSlideSvgElement(index: number): SVGSVGElement | null {
    if (!this.engine || index < 0 || index >= this.engine.slideCount) return null;

    const safeSvg = this.prepareSvgForRender(this.engine.renderSlide(index).svg, true);
    if (!safeSvg.allowed) return null;

    const element = createSvgElementFromString(safeSvg.svg, this.contentEl.ownerDocument);
    if (!element) return null;

    this.engine.applyFontFidelity(element);
    this.engine.formatChartAxisLabels(element, index);
    normalizeSvgForDisplay(element);
    return element;
  }

  private startPresentation(): void {
    if (!this.engine || this.engine.slideCount === 0) {
      pptNotice('powerpoint:notice.openToPresent');
      return;
    }

    debugLog('view', 'Starting PowerPoint presentation', {
      slide: this.currentSlide,
      slideCount: this.engine.slideCount
    });
    this.presentController?.dispose();

    const controller = new PowerPointPresentController({
      ownerDocument: this.contentEl.ownerDocument,
      slideCount: this.engine.slideCount,
      startIndex: this.currentSlide,
      renderSlide: (index) => this.buildSlideSvgElement(index),
      onExit: (lastIndex) => {
        this.presentController = null;
        if (this.engine && lastIndex >= 0 && lastIndex < this.engine.slideCount) {
          this.slideFilmstripController.navigateToSlide(lastIndex, 'presentation-exit');
        }
      },
      t: this.t
    });

    this.presentController = controller;
    controller.start();
  }

  private collectExportSvgElements(indices: number[]): SVGSVGElement[] {
    const elements: SVGSVGElement[] = [];
    for (const index of indices) {
      const element = this.buildSlideSvgElement(index);
      if (element) elements.push(element);
    }
    return elements;
  }

  private captureHistoryEntry(label: string): Promise<HistoryEntry> {
    return this.historyController.capture(label);
  }

  private captureSlideXmlHistoryEntry(slideIndex: number, label: string): HistorySlideXmlEntry {
    return this.historyController.captureSlideXml(slideIndex, label);
  }

  private completeSlideXmlHistoryEntry(entry: HistorySlideXmlEntry): HistorySlideXmlEntry {
    return this.historyController.completeSlideXml(entry);
  }

  private recordHistoryEntry(entry: HistoryEntry): void {
    this.historyController.record(entry);
  }

  private registerKeyboardHandlers(): void {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (!this.containerEl.isShown()) return;

      if (isPrimaryFindShortcut(event) && this.isActivePowerPointView()) {
        const target = isElement(event.target) ? event.target : null;
        if (!target?.closest('.modal')) {
          event.preventDefault();
          event.stopImmediatePropagation();
          this.findController.open();
          return;
        }
      }

      if (!this.isActivePowerPointView()) return;
      if (this.activeEditor && activeDocument.activeElement === this.activeEditor) {
        // Route undo/redo to the in-place editor history (which restores the
        // selection) and fall back to document history when the edit session is
        // exhausted. Capture-phase + stopImmediatePropagation keeps Obsidian and
        // the textarea's native undo from also firing. All other keys pass
        // through to the textarea unchanged.
        const lowerKey = event.key.toLowerCase();
        if ((event.metaKey || event.ctrlKey) && lowerKey === 'z') {
          event.preventDefault();
          event.stopImmediatePropagation();
          void (event.shiftKey ? this.handleInlineRedo() : this.handleInlineUndo());
          return;
        }
        if (event.ctrlKey && !event.metaKey && lowerKey === 'y') {
          event.preventDefault();
          event.stopImmediatePropagation();
          void this.handleInlineRedo();
          return;
        }
        return;
      }

      const target = isElement(event.target) ? event.target : null;
      if (target?.closest('input, textarea, select, [contenteditable="true"]')) {
        if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 's') {
          event.preventDefault();
          event.stopImmediatePropagation();
          void this.saveCurrentPresentation();
        }
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        event.stopImmediatePropagation();
        void this.saveCurrentPresentation();
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'z') {
        event.preventDefault();
        event.stopImmediatePropagation();
        void this.requestHistoryAction(event.shiftKey ? 'redo' : 'undo', 'keyboard');
        return;
      }

      if (event.ctrlKey && !event.metaKey && event.key.toLowerCase() === 'y') {
        event.preventDefault();
        event.stopImmediatePropagation();
        void this.requestHistoryAction('redo', 'keyboard');
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'c') {
        if (this.getSelectedIndices().length > 0) {
          event.preventDefault();
          event.stopImmediatePropagation();
          void this.copySelectedShape();
        }
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'v') {
        if (this.objectClipboard) {
          event.preventDefault();
          event.stopImmediatePropagation();
          void this.pasteCopiedShape();
        }
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'd') {
        if (this.selectedShapeIndex !== null) {
          event.preventDefault();
          event.stopImmediatePropagation();
          void this.duplicateSelectedShape();
        }
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'a') {
        event.preventDefault();
        event.stopImmediatePropagation();
        if (this.lastInteractionRegion === 'thumbnails') {
          this.slideFilmstripController.selectAllSlides();
        } else {
          this.selectAllShapes();
        }
        return;
      }

      const isArrowKey =
        event.key === 'ArrowUp'
        || event.key === 'ArrowDown'
        || event.key === 'ArrowLeft'
        || event.key === 'ArrowRight';
      const hasShapeSelection = this.selectedShapeIndex !== null || this.selectedShapeIndices.size > 0;
      if (isArrowKey && hasShapeSelection) {
        event.preventDefault();
        event.stopImmediatePropagation();
        void this.nudgeSelection(event.key, event.shiftKey);
        return;
      }

      if (event.key === 'Escape' && this.slideFilmstripController.selectedSlideIndices.size > 0) {
        event.preventDefault();
        this.slideFilmstripController.clearSlideSelection();
        return;
      }

      if (
        (event.key === 'Delete' || event.key === 'Backspace')
        && this.lastInteractionRegion === 'thumbnails'
      ) {
        if (this.slideFilmstripController.selectedSlideIndices.size > 0) {
          event.preventDefault();
          void this.slideFilmstripController.deleteSelectedSlides();
          return;
        }
        event.preventDefault();
        void this.slideFilmstripController.deleteSlide();
        return;
      }

      if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
        event.preventDefault();
        this.slideFilmstripController.navigateToSlide(this.currentSlide - 1, 'keyboard-prev');
      } else if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
        event.preventDefault();
        this.slideFilmstripController.navigateToSlide(this.currentSlide + 1, 'keyboard-next');
      } else if (event.key === 'Delete' || event.key === 'Backspace') {
        if (this.selectedShapeIndex !== null || this.selectedShapeIndices.size > 0) {
          event.preventDefault();
          void this.deleteSelectedShape();
        }
      }
    };

    this.registerDomEvent(window, 'keydown', handleKeyDown, true);
    this.registerDomEvent(activeDocument, 'keydown', handleKeyDown, true);

    this.registerDomEvent(window, 'resize', () => this.updateSlideScale());
    this.registerDomEvent(activeDocument, 'pointermove', this.handleDragMove, true);
    this.registerDomEvent(activeDocument, 'pointerup', this.handleDragEnd, true);
    this.registerDomEvent(activeDocument, 'pointerdown', this.handleOutsideSlidePointerDown, true);
  }

  private isActivePowerPointView(): boolean {
    if (this.app.workspace.getActiveViewOfType(NativePowerPointView) === this) {
      return true;
    }

    if (this.contentEl.closest('.workspace-leaf.mod-active')) {
      return true;
    }

    const activeElement = activeDocument.activeElement;
    return Boolean(isNode(activeElement) && this.contentEl.contains(activeElement));
  }

  private async loadPresentation(file: TFile): Promise<void> {
    const loadStartedAt = performance.now();
    debugLog('file', 'PowerPoint load started', { file: file.path, extension: file.extension });
    this.session.reset();
		this.presentationWordCount = 0;
		this.presentationWordCountEditVersion = -1;
		this.onWordCountClear();
    this.removeActiveEditor();
    this.historyController.clear();
    this.dragState = null;
    this.isLoading = true;
    this.engine = null;
    this.loadedFile = file;
    this.sourcePackage = null;
    this.sourceBuffer = null;
    this.selectedShapeIndex = null;
    this.selectedTransform = null;
    this.svgSecurityDecision = null;
    this.isViewOnly = false;
    this.viewOnlyReason = '';
    this.hasShownGeneratedTextNotice = false;
    this.fontSubstitutions = [];
    this.findController.reset();
    this.session.setSaveStatus('idle');
    this.updateEditingAvailability();
    if (this.getSettings().showInspector) {
      this.renderInspector();
    }
    this.showLoading(this.t('powerpoint:loading.loadingFile', { fileName: file.name }));

    if (!isModernPowerPointExtension(file.extension)) {
      this.isLoading = false;
      warnLog('file', 'PowerPoint load rejected unsupported format', {
        file: file.path,
        extension: file.extension
      });
      this.showUnsupported(file);
      return;
    }

    try {
      const buffer = await this.app.vault.readBinary(file);
      const sourcePackage = inspectPowerPointPackage(buffer);
      const sourceValidation = validatePowerPointPackageStructure(sourcePackage);
      if (!sourceValidation.ok) {
        throw new Error(summarizePackageMessages(sourceValidation.errors));
      }

      this.sourcePackage = sourcePackage;
      this.sourceBuffer = buffer;
      this.isViewOnly = this.shouldOpenViewOnly(file, sourcePackage);
      this.viewOnlyReason = this.getViewOnlyReason(file, sourcePackage);
      this.engine = await PresentationEngine.load(buffer);
		this.refreshPresentationWordCount();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.scheduleFilmstripRender();
      }
      this.importPendingAgentUndoHistory();
      this.session.setSaveStatus(this.isViewOnly ? 'view-only' : 'saved');
      debugLog('file', 'PowerPoint load completed', {
        file: file.path,
        bytes: buffer.byteLength,
        slideCount: this.engine.slideCount,
        viewOnly: this.isViewOnly,
        hasVbaProject: sourcePackage.hasVbaProject,
        ms: Math.round(performance.now() - loadStartedAt)
      });
      if (this.isViewOnly) {
        pptNotice('powerpoint:notice.openedViewOnly', { fileName: file.name, reason: this.viewOnlyReason });
      }
    } catch (error) {
      errorLog('file', 'PowerPoint load failed', {
        file: file.path,
        extension: file.extension,
        error
      });
      if (isWasmGcUnsupportedError(error)) {
        this.showRuntimeUnsupportedError(file);
      } else {
        this.showError(this.t('powerpoint:loading.couldNotOpen', { fileName: file.name, message: cleanError(error) }));
      }
    } finally {
      this.isLoading = false;
      this.updateSlideCounter();
      this.updateEditingAvailability();
      if (this.getSettings().showInspector) {
        this.renderInspector();
      }
    }
  }

  private showLoading(message: string): void {
    if (!this.slideSurface) return;
    this.resetSlideSurfaceSizing();
    this.slideSurface.empty();
    this.slideSurface.createDiv({ cls: 'native-powerpoint-loading', text: message });
    this.thumbnailContainer?.empty();
  }

  private showUnsupported(file: TFile): void {
    if (!this.slideSurface) return;
    this.resetSlideSurfaceSizing();
    this.slideSurface.empty();
    this.slideSurface.createDiv({
      cls: 'native-powerpoint-error',
      text: this.t('powerpoint:loading.legacyFormat', { extension: file.extension.toUpperCase() })
    });
    this.thumbnailContainer?.empty();
    this.session.setSaveStatus('idle');
    this.updateEditingAvailability();
  }

  private showError(message: string): void {
    if (!this.slideSurface) return;
    this.resetSlideSurfaceSizing();
    this.slideSurface.empty();
    this.slideSurface.createDiv({ cls: 'native-powerpoint-error', text: message });
    this.thumbnailContainer?.empty();
    this.session.setSaveStatus('failed');
    this.updateEditingAvailability();
  }

  private showRuntimeUnsupportedError(file: TFile): void {
    if (!this.slideSurface) return;
    this.resetSlideSurfaceSizing();
    this.slideSurface.empty();
    this.thumbnailContainer?.empty();

    const chromeVersion = getChromiumVersion();
    const isMobile = Platform.isMobileApp;
    const notice = this.slideSurface.createDiv({ cls: 'native-powerpoint-runtime-error' });
    notice.createDiv({
      cls: 'native-powerpoint-runtime-error-title',
      text: this.t('powerpoint:runtime.title')
    });
    const chromeVersionSuffix = chromeVersion
      ? this.t('powerpoint:runtime.chromeVersionSuffix', { version: chromeVersion })
      : '';
    const desktopChromeSuffix = chromeVersion
      ? this.t('powerpoint:runtime.desktopChromeSuffix', { version: chromeVersion })
      : '';
    notice.createEl('p', {
      text: isMobile
        ? this.t('powerpoint:runtime.mobileDescription', { fileName: file.name, chromeVersion: chromeVersionSuffix })
        : this.t('powerpoint:runtime.desktopDescription', { fileName: file.name, chromeVersion: desktopChromeSuffix })
    });

    if (!isMobile) {
      notice.createEl('p', {
        text: this.t('powerpoint:runtime.reinstallHint')
      });

      const actions = notice.createDiv({ cls: 'native-powerpoint-runtime-error-actions' });
      actions.createEl('a', {
        cls: 'native-powerpoint-runtime-error-link',
        text: this.t('powerpoint:runtime.downloadObsidian'),
        href: OBSIDIAN_DOWNLOAD_URL
      });

      notice.createEl('p', {
        cls: 'native-powerpoint-runtime-error-hint',
        text: this.t('powerpoint:runtime.installerVersionHint')
      });
    } else {
      notice.createEl('p', {
        cls: 'native-powerpoint-runtime-error-hint',
        text: this.t('powerpoint:runtime.mobileFallbackHint')
      });
    }

    this.session.setSaveStatus('failed');
    this.updateEditingAvailability();
  }

  private async renderCurrentSlide(keepSelection = false, expectedGeneration?: number): Promise<boolean> {
    if (!this.engine || !this.slideSurface) return false;
    if (expectedGeneration !== undefined && expectedGeneration !== this.slideRenderGeneration) {
      debugLog('render', 'skipped stale slide render', {
        slide: this.currentSlide,
        expectedGeneration,
        currentGeneration: this.slideRenderGeneration
      });
      return false;
    }

    const renderStarted = performance.now();
    const slideIndex = this.currentSlide;
    debugLog('render', 'renderCurrentSlide start', { slide: slideIndex, keepSelection, expectedGeneration });

    const selectedShape = keepSelection ? this.selectedShapeIndex : null;
    const { svg } = this.engine.renderSlide(slideIndex);
    const safeSvg = this.prepareSvgForRender(svg);

    if (!safeSvg.allowed) {
      this.showUnsafeSvgWarning(safeSvg.issues);
      return false;
    }

    const svgElement = createSvgElementFromString(safeSvg.svg, this.slideSurface.ownerDocument);
    if (!svgElement) {
      this.showError(this.t('powerpoint:loading.couldNotRenderSlide'));
      return false;
    }

    this.slideSurface.empty();
    this.slideSurface.appendChild(svgElement);
    this.svgEl = svgElement;

    if (this.svgEl) {
      this.fontSubstitutions = this.engine.applyFontFidelity(this.svgEl);
      this.engine.formatChartAxisLabels(this.svgEl, this.currentSlide);
      normalizeSvgForDisplay(this.svgEl);
      this.markGeneratedTextEditability(this.svgEl);
      this.annotateSlideTextOffsets();
      this.svgEl.addClass('native-powerpoint-slide-svg');
      this.slideSurface.addClass('is-rendered');
      this.updateSlideScale();
      window.requestAnimationFrame(() => this.updateSlideScale());
      this.attachSvgEvents();
      this.applyRunHighlights();
    }

    if (selectedShape !== null) {
      this.selectShape(selectedShape);
    } else {
      this.clearSelection();
    }

    this.findController.refreshHighlight();
    this.updateSlideCounter();
    const renderMs = Math.round(performance.now() - renderStarted);
    debugLog('render', 'renderCurrentSlide complete', { slide: slideIndex, keepSelection, ms: renderMs });
    if (renderMs > 750) {
      warnLog('render', 'slow renderCurrentSlide', { slide: slideIndex, ms: renderMs });
    }
    return true;
  }

  /**
   * Re-render a single edited shape and swap it into the live slide SVG, instead
   * of tearing down and rebuilding the whole slide. The engine already exposes a
   * single-shape render (`renderShape` -> `renderer.renderShapeSvg`) that returns
   * exactly the `<g data-ooxml-shape-idx>` node that lives as a direct child of
   * the slide `<svg>`, so we can parse it, sanitize it through the same funnel as
   * a full render, and replace only that one group.
   *
   * This kills the flicker (no `slideSurface.empty()` + re-append + re-scale) and
   * most of the save/restore-selection choreography: every other shape's DOM,
   * the inline-editor textarea (it lives in the canvas pane, not the SVG), the
   * delegated SVG event listeners (bound on the root `<svg>`), the slide scale,
   * and the scroll position all survive untouched. The caller still re-points an
   * open inline editor at the freshly rendered run nodes for the edited shape.
   *
   * Returns false when an incremental swap is not safe/possible (shape group
   * missing, renderer error, or the sanitizer would suppress content), so the
   * caller can fall back to a full {@link renderCurrentSlide}.
   */
  private renderShapeInPlace(shapeIndex: number): boolean {
    const svg = this.svgEl;
    const engine = this.engine;
    if (!svg || !engine) return false;

    const existing = svg.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!isSVGGElement(existing)) return false;
    // Nested groups can carry the same idx attribute; only swap a top-level shape.
    if (existing.parentElement?.closest('g[data-ooxml-shape-idx]')) return false;

    let shapeSvg: string;
    try {
      shapeSvg = engine.renderShape(this.currentSlide, shapeIndex);
    } catch (error) {
      warnLog('render', 'renderShapeInPlace render failed; falling back', {
        slide: this.currentSlide,
        shapeIndex,
        error: cleanError(error)
      });
      return false;
    }

    // Wrap the bare `<g>` fragment in a minimal SVG so it runs through the same
    // sanitizer/parse path as a full slide render; bail to a full render if the
    // sanitizer would hide content (the full path surfaces the security prompt).
    const wrapper = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">${shapeSvg}</svg>`;
    const safeSvg = this.prepareSvgForRender(wrapper);
    if (!safeSvg.allowed) return false;

    const wrapperEl = createSvgElementFromString(safeSvg.svg, svg.ownerDocument);
    const newGroup = wrapperEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`) ?? null;
    if (!isSVGGElement(newGroup)) return false;

    existing.replaceWith(newGroup);

    // Re-run the per-shape render passes on the swapped node. Font fidelity is
    // scoped to the new group; the grid/halo/chart-axis/editable passes are
    // idempotent (attribute-only) so running them slide-wide is safe and keeps
    // unchanged shapes correct.
    engine.applyFontFidelity(newGroup);
    engine.formatChartAxisLabels(svg, this.currentSlide);
    normalizeSvgForDisplay(svg);
    this.markGeneratedTextEditability(svg);
    this.annotateShapeTextOffsets(newGroup);
    this.applyRunHighlights();

    // The replaced group dropped its `native-powerpoint-shape-selected` class and
    // any selection box geometry; re-apply selection state without re-running the
    // full select choreography.
    this.applySelectionClasses();
    if (this.selectedShapeIndex === shapeIndex || this.selectedShapeIndices.has(shapeIndex)) {
      this.updateSelectionOverlay();
    }

    this.findController.refreshHighlight();
    debugLog('render', 'renderShapeInPlace swapped shape', { slide: this.currentSlide, shapeIndex });
    return true;
  }

  /**
   * Re-render after a single-shape edit: try the incremental in-place swap and
   * fall back to a full keep-selection render when that is not possible. Drop-in
   * for `renderCurrentSlide(true)` at single-shape edit sites.
   */
  private async renderEditedShape(shapeIndex: number): Promise<boolean> {
    if (this.renderShapeInPlace(shapeIndex)) return true;
    return this.renderCurrentSlide(true);
  }

  private async renderThumbnails(): Promise<void> {
    return this.slideFilmstripController.renderThumbnails();
  }

  private scheduleThumbnailRefresh(indices: number | number[]): void {
    this.slideFilmstripController.scheduleThumbnailRefresh(indices);
  }

  private scheduleFilmstripRender(force = false): void {
    if (!force && (this.filmstripRenderScheduled || this.filmstripRendered || !this.engine || !this.thumbnailContainer)) {
      return;
    }

    this.filmstripRenderScheduled = true;
    const cancelIdle = scheduleIdleWork(() => {
      this.filmstripRenderScheduled = false;
      if (!this.engine || !this.thumbnailContainer?.isConnected) {
        return;
      }

		void this.renderThumbnails()
			.then(() => {
				this.filmstripRendered = true;
			})
			.catch((error) => {
				errorLog('render', 'PowerPoint filmstrip render failed', { error: cleanError(error) });
			});
    }, { timeout: 3000 });

    this.register(() => cancelIdle());
  }

  private navigateToSlide(index: number, reason: string): void {
    this.slideFilmstripController.navigateToSlide(index, reason);
  }


  private async finishInlineTextEditing(reason: string): Promise<void> {
    if (this.textCommitPromise) {
      debugLog('text-edit', 'awaiting in-flight text commit', { reason });
      await this.textCommitPromise.catch((error: unknown) => {
        errorLog('text-edit', 'text commit failed while waiting', { reason, error: cleanError(error) });
      });
      return;
    }

    if (!this.activeEditor) return;

    debugLog('text-edit', 'finishInlineTextEditing', {
      reason,
      slide: this.currentSlide,
      hasCommit: this.activeEditorCommit !== null
    });

    const commit = this.activeEditorCommit;
    if (!commit) {
      this.removeActiveEditor();
      return;
    }

    this.isTearingDownEditor = true;
    this.textCommitPromise = commit()
      .catch((error: unknown) => {
        errorLog('text-edit', 'inline text commit failed', { reason, error: cleanError(error) });
      })
      .finally(() => {
        this.textCommitPromise = null;
        this.isTearingDownEditor = false;
      });
    await this.textCommitPromise;
  }

  private prepareSvgForRender(svg: string, isThumbnail = false): { svg: string; issues: SvgSecurityIssue[]; allowed: boolean } {
    const settings = this.getSettings();
    if (settings.openWithYoloMode || this.svgSecurityDecision === 'yolo') {
      return {
        svg,
        issues: [],
        allowed: true
      };
    }

    const sanitizerMode =
      this.svgSecurityDecision === 'compatibility' || !settings.hideUnsupportedSvgContent
        ? 'compatibility'
        : 'strict';
    const scannedSvg = sanitizeSvg(svg, { mode: sanitizerMode });
    const shouldWarn =
      this.svgSecurityDecision === null &&
      scannedSvg.issues.length > 0 &&
      !isThumbnail;
    return {
      svg: scannedSvg.svg,
      issues: scannedSvg.issues,
      allowed: !shouldWarn
    };
  }

  private showUnsafeSvgWarning(issues: SvgSecurityIssue[]): void {
    if (!this.slideSurface) return;

    this.svgEl = null;
    this.clearSelection();
    this.resetSlideSurfaceSizing();
    this.slideSurface.empty();
    this.thumbnailContainer?.empty();

    const warning = this.slideSurface.createDiv({ cls: 'native-powerpoint-security-warning' });
    warning.createDiv({ cls: 'native-powerpoint-security-title', text: this.t('powerpoint:security.title') });
    warning.createEl('p', {
      text: this.t('powerpoint:security.description')
    });

    const summary = summarizeSvgSecurityIssues(issues);
    const list = warning.createEl('ul', { cls: 'native-powerpoint-security-list' });
    for (const item of summary.slice(0, 6)) {
      list.createEl('li', { text: item });
    }

    if (summary.length > 6) {
      list.createEl('li', { text: this.t('powerpoint:security.moreIssuesHidden', { count: summary.length - 6 }) });
    }

    const actions = warning.createDiv({ cls: 'native-powerpoint-security-actions' });
    const openCompatibility = actions.createEl('button', { text: this.t('powerpoint:security.openCompatibilityMode') });
    openCompatibility.addClass('mod-warning');
    openCompatibility.addEventListener('click', () => {
      this.svgSecurityDecision = 'compatibility';
      pptNotice('powerpoint:notice.openingCompatibilityMode');
		void this.renderCurrentSlide()
			.then((rendered) => {
				if (rendered) void this.renderThumbnails();
			})
			.catch((error) => {
				errorLog('render', 'PowerPoint compatibility-mode render failed', { error: cleanError(error) });
			});
    });

    const openYolo = actions.createEl('button', { text: this.t('powerpoint:security.openYoloMode') });
    openYolo.addClass('mod-warning');
    openYolo.addEventListener('click', () => {
      this.svgSecurityDecision = 'yolo';
      pptNotice('powerpoint:notice.openingYoloMode');
		void this.renderCurrentSlide()
			.then((rendered) => {
				if (rendered) void this.renderThumbnails();
			})
			.catch((error) => {
				errorLog('render', 'PowerPoint YOLO-mode render failed', { error: cleanError(error) });
			});
    });

    const rememberYolo = actions.createEl('button', { text: this.t('powerpoint:security.alwaysUseYoloMode') });
    rememberYolo.addClass('mod-warning');
    rememberYolo.addEventListener('click', () => {
      this.svgSecurityDecision = 'yolo';
      void this.getSettings().setOpenWithYoloMode(true)
        .then(() => {
          pptNotice('powerpoint:notice.yoloModeRemembered');
        })
        .catch((error) => {
          pptNotice('powerpoint:notice.couldNotRememberYoloMode', { message: cleanError(error) });
        })
        .finally(() => {
			void this.renderCurrentSlide()
				.then((rendered) => {
					if (rendered) void this.renderThumbnails();
				})
				.catch((error) => {
					errorLog('render', 'PowerPoint remembered YOLO-mode render failed', { error: cleanError(error) });
				});
		});
    });
  }

  private resetSlideSurfaceSizing(): void {
    if (!this.slideSurface) return;

    this.slideSurface.removeClass('is-rendered');
    this.slideSurface.removeClass('is-scaled');
    this.slideSurface.style.removeProperty('--native-powerpoint-slide-width');
    this.slideSurface.style.removeProperty('--native-powerpoint-slide-height');

    if (this.svgEl) {
      this.svgEl.style.removeProperty('width');
      this.svgEl.style.removeProperty('height');
      this.svgEl.style.removeProperty('transform');
      this.svgEl.style.removeProperty('transform-origin');
    }
  }

  private attachSvgEvents(): void {
    if (!this.svgEl) return;

    this.svgEl.addEventListener('click', (event) => {
      if (this.suppressNextClick) {
        this.suppressNextClick = false;
        event.preventDefault();
        event.stopPropagation();
        return;
      }
      const target = isElement(event.target) ? event.target : null;
      if (this.consumeInlineTextClick(event, target)) return;
      this.suppressNextTextClick = false;

      if (target?.closest('text')) {
        const shape = target?.closest('g[data-ooxml-shape-idx]') ?? null;
        const rawShapeIndex = getShapeIndex(shape);
        const shapeIndex = isSelectableShapeIndex(rawShapeIndex) ? rawShapeIndex : null;
        const additive = event.shiftKey || event.ctrlKey || event.metaKey;
        if (shapeIndex === null) {
          return;
        }
        if (additive) {
          event.preventDefault();
          event.stopPropagation();
          this.toggleShapeInSelection(shapeIndex);
          return;
        }
        if (this.selectedShapeIndex !== shapeIndex) {
          this.selectShape(shapeIndex);
        }
        return;
      }

      const shape = target?.closest('g[data-ooxml-shape-idx]') ?? null;
      const rawShapeIndex = getShapeIndex(shape);
      const shapeIndex = isSelectableShapeIndex(rawShapeIndex) ? rawShapeIndex : null;
      const additive = event.shiftKey || event.ctrlKey || event.metaKey;
      if (shapeIndex === null) {
        if (!additive) this.clearSelection();
        return;
      }

      if (additive) {
        event.preventDefault();
        event.stopPropagation();
        this.toggleShapeInSelection(shapeIndex);
        return;
      }

      this.selectShape(shapeIndex);
    });

    this.svgEl.addEventListener('dblclick', (event) => {
      const target = isElement(event.target) ? event.target : null;
      const shape = target?.closest('g[data-ooxml-shape-idx]') ?? null;
      const shapeIndex = getShapeIndex(shape);
      if (shapeIndex !== null) {
        event.preventDefault();
        if (target?.closest('text')) {
          this.selectShapeForTextEditing(shapeIndex);
        } else {
          this.selectShape(shapeIndex);
        }
        if (target?.closest('text') && target.closest(GENERATED_GRID_SELECTOR)) {
          const textTarget = this.getGeneratedTextEditTarget(target);
          if (textTarget && this.ensureEditable('edit text')) {
            this.startTextEditor(textTarget, event.clientX, event.clientY);
            this.applyInlineMultiClickSelectionAtPoint(textTarget, event.clientX, event.clientY, 2);
          } else {
            this.showGeneratedTextNotice();
          }
          return;
        }

        const textTarget = this.getTextEditTarget(target);
        if (this.ensureEditable('edit text')) {
          this.startTextEditor(textTarget, event.clientX, event.clientY);
          if (textTarget) {
            this.applyInlineMultiClickSelectionAtPoint(textTarget, event.clientX, event.clientY, 2);
          }
        }
      }
    });

    this.svgEl.addEventListener('pointerdown', (event) => {
      if (event.button !== 0) return;

      const target = isElement(event.target) ? event.target : null;
      if (target?.closest('text')) {
        if (this.activeEditor) {
          const shape = target.closest('g[data-ooxml-shape-idx]');
          const editingShape = this.activeEditorTarget?.closest('g[data-ooxml-shape-idx]');
          if (shape === editingShape) {
            this.handleInlineTextPointerDown(event, target);
            return;
          }
          this.commitActiveTextEditing();
        }

        const shape = target.closest('g[data-ooxml-shape-idx]');
        const rawShapeIndex = getShapeIndex(shape);
        const shapeIndex = isSelectableShapeIndex(rawShapeIndex) ? rawShapeIndex : null;
        const additive = event.shiftKey || event.ctrlKey || event.metaKey;
        if (this.shouldStartGroupDragFromText(shapeIndex, additive)) {
          event.preventDefault();
          event.stopPropagation();
          this.suppressNextClick = true;
          debugLog('selection', 'Starting multi-object drag from PowerPoint text', {
            slide: this.currentSlide,
            shapeIndex,
            count: this.selectedShapeIndices.size,
            shapeIndexes: [...this.selectedShapeIndices],
          });
          if (this.ensureEditable('move objects')) {
            this.startGroupDrag(event);
          }
          return;
        }
        if (additive) return;
        if (shapeIndex === null) return;

        this.handleInlineTextPointerDown(event, target);
        return;
      }

      if (this.activeEditor) {
        this.commitActiveTextEditing();
      }

      const shape = target?.closest('g[data-ooxml-shape-idx]') ?? null;
      const rawShapeIndex = getShapeIndex(shape);
      const shapeIndex = isSelectableShapeIndex(rawShapeIndex) ? rawShapeIndex : null;
      const additive = event.shiftKey || event.ctrlKey || event.metaKey;

      if (shapeIndex === null) {
        event.preventDefault();
        if (!additive) this.clearSelection();
        this.beginMarquee(event, additive);
        return;
      }

      if (additive) {
        // Let the click handler toggle this shape in/out of the selection
        // instead of starting a drag.
        event.preventDefault();
        return;
      }

      if (this.selectedShapeIndices.size > 1 && this.selectedShapeIndices.has(shapeIndex)) {
        event.preventDefault();
        if (this.ensureEditable('move objects')) {
          this.startGroupDrag(event);
        }
        return;
      }

      if (this.selectedShapeIndex === shapeIndex && this.selectedTransform !== null) {
        event.preventDefault();
        if (this.ensureEditable('move object')) {
          this.startDrag(event, 'move');
        }
        return;
      }
    });
  }

  /**
   * Consume the click generated by an inline text pointer sequence. Pointer
   * capture retargets that click to the SVG root once a drag leaves its glyph,
   * so checking for a text target here would let the canvas clear the editor.
   */
  private consumeInlineTextClick(event: MouseEvent, target: Element | null): boolean {
    if (!this.suppressNextTextClick) return false;

    event.preventDefault();
    event.stopPropagation();
    this.suppressNextTextClick = false;

    const wasRetargetedByCapture = !target?.closest('text');
    if (wasRetargetedByCapture) {
      debugLog('text-select', 'suppressed captured inline click', {
        slide: this.currentSlide,
        shapeIndex: this.activeShapeTextTarget?.shapeIndex ?? this.selectedShapeIndex,
        targetTag: target?.tagName.toLowerCase() ?? null,
        detail: event.detail,
      });
    }

    if (event.detail < 2) return true;

    const textTarget = target?.closest('text')
      ? target.closest(GENERATED_GRID_SELECTOR)
        ? this.getGeneratedTextEditTarget(target)
        : this.getTextEditTarget(target)
      : this.activeShapeTextTarget;
    if (textTarget) {
      this.applyInlineMultiClickSelectionAtPoint(textTarget, event.clientX, event.clientY, event.detail);
    }
    return true;
  }


  private async deleteSelectedShape(): Promise<void> {
    if (!this.engine) return;
    if (this.selectedShapeIndices.size > 1) {
      await this.deleteSelectedShapes();
      return;
    }
    if (this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('delete object')) return;

    const shapeIndex = this.selectedShapeIndex;
    debugLog('selection', 'Deleting PowerPoint object', { slide: this.currentSlide, shapeIndex });
    try {
      const history = await this.captureHistoryEntry('Delete object');
      await this.engine.deleteShape(this.currentSlide, shapeIndex);
      this.clearSelection();
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
      debugLog('selection', 'Deleted PowerPoint object', { slide: this.currentSlide, shapeIndex });
    } catch (error) {
      errorLog('selection', 'PowerPoint object deletion failed', { slide: this.currentSlide, shapeIndex, error });
      pptNotice('powerpoint:notice.couldNotDeleteObject', { message: cleanError(error) });
    }
  }

  private async deleteSelectedShapes(): Promise<void> {
    if (!this.engine || this.selectedShapeIndices.size === 0) return;
    if (!this.ensureEditable('delete objects')) return;

    const indices = [...this.selectedShapeIndices].sort((a, b) => b - a);
    debugLog('selection', 'Deleting PowerPoint objects', {
      slide: this.currentSlide,
      count: indices.length,
      shapeIndexes: indices
    });
    const startedAt = Date.now();
    try {
      const history = await this.captureHistoryEntry('Delete objects');
      await this.engine.deleteShapes(this.currentSlide, indices);
      this.clearSelection();
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) this.scheduleThumbnailRefresh(this.currentSlide);
      debugLog('selection', 'Deleted PowerPoint objects', {
        slide: this.currentSlide,
        count: indices.length,
        shapeIndexes: indices,
        ms: Date.now() - startedAt,
      });
    } catch (error) {
      errorLog('selection', 'PowerPoint multi-object deletion failed', {
        slide: this.currentSlide,
        shapeIndexes: indices,
        error
      });
      pptNotice('powerpoint:notice.couldNotDeleteObjects', { message: cleanError(error) });
    }
  }

  private async copySelectedShape(): Promise<void> {
    const shapeIndexes = this.getSelectedIndices();
    if (!this.engine || shapeIndexes.length === 0) {
      pptNotice('powerpoint:notice.selectObjectToCopy');
      return;
    }

    debugLog('clipboard', 'Copying PowerPoint objects', {
      slide: this.currentSlide,
      count: shapeIndexes.length,
      shapeIndexes,
    });
    try {
      this.objectClipboard = await this.engine.copyShapes(this.currentSlide, shapeIndexes);
      this.updateObjectClipboardAvailability();
      debugLog('clipboard', 'Copied PowerPoint objects', {
        slide: this.currentSlide,
        count: this.objectClipboard.shapeIndexes.length,
        shapeIndexes: this.objectClipboard.shapeIndexes,
      });
      pptNotice('powerpoint:notice.copiedSlideObject');
    } catch (error) {
      errorLog('clipboard', 'PowerPoint object copy failed', {
        slide: this.currentSlide,
        count: shapeIndexes.length,
        shapeIndexes,
        error
      });
      pptNotice('powerpoint:notice.couldNotCopyObject', { message: cleanError(error) });
    }
  }

  private async pasteCopiedShape(): Promise<void> {
    if (!this.engine || !this.objectClipboard) {
      pptNotice('powerpoint:notice.copyObjectFirst');
      return;
    }
    if (!this.ensureEditable('paste object')) return;

    debugLog('clipboard', 'Pasting PowerPoint objects', {
      slide: this.currentSlide,
      count: this.objectClipboard.shapeIndexes.length,
      shapeIndexes: this.objectClipboard.shapeIndexes,
    });
    try {
      const history = await this.captureHistoryEntry('Paste objects');
      const shapeIndexes = await this.engine.pasteShapes(this.objectClipboard, this.currentSlide);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.applyMultiSelection(shapeIndexes);
        await this.renderThumbnails();
      }
      debugLog('clipboard', 'Pasted PowerPoint objects', {
        slide: this.currentSlide,
        count: shapeIndexes.length,
        shapeIndexes,
      });
    } catch (error) {
      errorLog('clipboard', 'PowerPoint object paste failed', {
        slide: this.currentSlide,
        count: this.objectClipboard.shapeIndexes.length,
        shapeIndexes: this.objectClipboard.shapeIndexes,
        error,
      });
      pptNotice('powerpoint:notice.couldNotPasteObject', { message: cleanError(error) });
    }
  }

  private async duplicateSelectedShape(): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) {
      pptNotice('powerpoint:notice.selectObjectToDuplicate');
      return;
    }
    if (!this.ensureEditable('duplicate object')) return;

    const sourceShapeIndex = this.selectedShapeIndex;
    debugLog('clipboard', 'Duplicating PowerPoint object', {
      slide: this.currentSlide,
      sourceShapeIndex
    });
    try {
      const history = await this.captureHistoryEntry('Duplicate object');
      const shapeIndex = await this.engine.duplicateShape(this.currentSlide, sourceShapeIndex);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
      debugLog('clipboard', 'Duplicated PowerPoint object', {
        slide: this.currentSlide,
        sourceShapeIndex,
        shapeIndex
      });
    } catch (error) {
      errorLog('clipboard', 'PowerPoint object duplication failed', {
        slide: this.currentSlide,
        sourceShapeIndex,
        error
      });
      pptNotice('powerpoint:notice.couldNotDuplicateObject', { message: cleanError(error) });
    }
  }

  private async cutSelectedShape(): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) {
      pptNotice('powerpoint:notice.selectObjectToCut');
      return;
    }
    if (!this.ensureEditable('cut object')) return;

    debugLog('clipboard', 'Cutting PowerPoint object', {
      slide: this.currentSlide,
      shapeIndex: this.selectedShapeIndex
    });
    try {
      this.objectClipboard = await this.engine.copyShape(this.currentSlide, this.selectedShapeIndex);
      this.updateObjectClipboardAvailability();
      await this.deleteSelectedShape();
      debugLog('clipboard', 'Cut PowerPoint object', { slide: this.currentSlide });
      pptNotice('powerpoint:notice.cutSlideObject');
    } catch (error) {
      errorLog('clipboard', 'PowerPoint object cut failed', { slide: this.currentSlide, error });
      pptNotice('powerpoint:notice.couldNotCutObject', { message: cleanError(error) });
    }
  }

  /**
   * Paste without formatting. With an active inline text editor this inserts
   * the clipboard's plain text at the caret. Otherwise it falls back to the
   * regular object paste (slide objects carry their own formatting, so there is
   * no plain-text variant for them).
   */
  private async pasteWithoutFormatting(): Promise<void> {
    if (this.activeEditor) {
      if (!this.ensureEditable('paste text')) return;
      debugLog('clipboard', 'Pasting PowerPoint plain text', { slide: this.currentSlide });
      try {
        const text = await navigator.clipboard.readText();
        if (text) this.insertPlainTextIntoActiveEditor(text);
      } catch {
        pptNotice('powerpoint:notice.plainTextUnavailable');
      }
      return;
    }

    await this.pasteCopiedShape();
  }

  private insertPlainTextIntoActiveEditor(text: string): void {
    const editor = this.activeEditor;
    if (!editor) return;

    const start = editor.selectionStart ?? editor.value.length;
    const end = editor.selectionEnd ?? editor.value.length;
    this.setInlineEditorValue(editor, `${editor.value.slice(0, start)}${text}${editor.value.slice(end)}`);
    const caret = start + text.length;
    editor.setSelectionRange(caret, caret);
    editor.dispatchEvent(new Event('input', { bubbles: true }));
  }

  private async rotateSelectedShape(deltaDegrees: number): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('rotate object')) return;

    const selected = this.getSelectedShapeElement();
    if (!selected) return;

    const transform = cloneTransform(this.engine.getShapeTransform(selected));
    const degrees = (((this.engine.ooxmlToDegrees(transform.rot) + deltaDegrees) % 360) + 360) % 360;
    transform.rot = this.engine.degreesToOoxml(degrees);
    await this.commitTransform(transform);
  }

  private async centerSelectedOnPage(axis: 'horizontal' | 'vertical'): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('center object')) return;

    const selected = this.getSelectedShapeElement();
    if (!selected) return;

    const transform = cloneTransform(this.engine.getShapeTransform(selected));
    const { cx, cy } = await this.engine.getSlideSizeEmu();
    if (axis === 'horizontal') {
      transform.x = Math.round((cx - transform.cx) / 2);
    } else {
      transform.y = Math.round((cy - transform.cy) / 2);
    }
    await this.commitTransform(transform);
  }

  private async flipSelectedShape(axis: 'horizontal' | 'vertical'): Promise<void> {
    if (this.selectedShapeIndex === null) return;
    await this.applyShapeMutation(
      this.selectedShapeIndex,
      'Flip object',
      'flip object',
      (slideIndex, shapeIndex) => this.engine!.flipShape(slideIndex, shapeIndex, axis)
    );
  }

  private openImageCropDialog(shapeIndex: number): void {
    if (!this.engine || !this.ensureEditable('crop image')) return;

    const current: ImageCrop = this.engine.getImageCrop(this.currentSlide, shapeIndex)
      ?? { left: 0, top: 0, right: 0, bottom: 0 };
    new ImageCropModal(this.app, current, (crop: ImageCropValues) => {
      void this.applyShapeMutation(
        shapeIndex,
        'Crop image',
        'crop image',
        (slideIndex, index) => this.engine!.setImageCrop(slideIndex, index, crop)
      );
    }).open();
  }

  private async resetSelectedImage(shapeIndex: number): Promise<void> {
    await this.applyShapeMutation(
      shapeIndex,
      'Reset image',
      'reset image',
      (slideIndex, index) => this.engine!.resetImage(slideIndex, index)
    );
  }

  private openReplaceImageVaultPicker(shapeIndex: number): void {
    if (!this.ensureEditable('replace image')) return;
    new VaultImageSuggestModal(this.app, (file) => void this.replaceImageWithVaultFile(shapeIndex, file), this.t).open();
  }

  private async replaceImageWithVaultFile(shapeIndex: number, file: TFile): Promise<void> {
    if (!this.engine) return;
    const bytes = new Uint8Array(await this.app.vault.readBinary(file));
    await this.applyShapeMutation(
      shapeIndex,
      'Replace image',
      'replace image',
      (slideIndex, index) => this.engine!.replaceImage(slideIndex, index, bytes, getImageMimeType(file.extension))
    );
  }

  private async replaceImageWithLocalFile(shapeIndex: number, file: File): Promise<void> {
    if (!this.engine) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    const mimeType = file.type || getImageMimeType(file.name.split('.').pop() ?? 'png');
    await this.applyShapeMutation(
      shapeIndex,
      'Replace image',
      'replace image',
      (slideIndex, index) => this.engine!.replaceImage(slideIndex, index, bytes, mimeType)
    );
  }

  private async setSelectedImageAsBackground(shapeIndex: number): Promise<void> {
    if (!this.engine || !this.ensureEditable('set slide background')) return;

    try {
      const image = await this.engine.getShapeImageData(this.currentSlide, shapeIndex);
      if (!image) {
        pptNotice('powerpoint:notice.selectedObjectNotImage');
        return;
      }

      const history = await this.captureHistoryEntry('Slide background image');
      await this.engine.setSlideBackgroundImage(this.currentSlide, image.bytes, image.mimeType);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) {
        await this.renderThumbnails();
        this.renderInspector();
      }
      debugLog('inspector', 'Changed PowerPoint slide background image', {
        slide: this.currentSlide,
        sourceShapeIndex: shapeIndex,
        mimeType: image.mimeType,
        bytes: image.bytes.byteLength
      });
    } catch (error) {
      errorLog('inspector', 'PowerPoint slide background image change failed', {
        slide: this.currentSlide,
        sourceShapeIndex: shapeIndex,
        error
      });
      pptNotice('powerpoint:notice.couldNotSetSlideBackground', { message: cleanError(error) });
    }
  }

  /**
   * Run an engine mutation against a single shape, threading it through the
   * shared dirty/history/re-render flow so undo/redo and persistence behave
   * like the other object operations.
   */
  private async applyShapeMutation(
    shapeIndex: number,
    historyLabel: string,
    action: string,
    mutate: (slideIndex: number, shapeIndex: number) => Promise<void>
  ): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable(action)) return;

    try {
      const history = await this.captureHistoryEntry(historyLabel);
      await mutate(this.currentSlide, shapeIndex);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderEditedShape(shapeIndex);
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      pptNotice('powerpoint:notice.couldNotAction', { action, message: cleanError(error) });
    }
  }

  private selectShape(shapeIndex: number): void {
    if (!this.engine || !this.svgEl || !isSelectableShapeIndex(shapeIndex)) return;

    this.session.selectShapes([shapeIndex]);
    this.selectedShapeIndex = shapeIndex;
    this.selectedShapeIndices = new Set([shapeIndex]);
    this.removeMultiSelectionBoxes();
    this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      shape.removeClass('native-powerpoint-shape-selected');
    });

    const selected = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!isSVGGElement(selected)) {
      this.selectedShapeIndex = null;
      this.selectedShapeIndices.clear();
      this.selectedTransform = null;
      this.removeSelectionOverlay();
      return;
    }

    selected.addClass('native-powerpoint-shape-selected');
    this.selectedTransform = cloneTransform(this.engine.getShapeTransform(selected));
    this.renderInspector();
    this.updateSelectionOverlay();
    this.updateObjectClipboardAvailability();
    this.updateTextToolbar();
  }

  /** Keep shape context and the outer selection box for inline text edits. */
  private selectShapeForTextEditing(shapeIndex: number): void {
    if (!this.engine || !this.svgEl || !isSelectableShapeIndex(shapeIndex)) return;

    this.session.selectShapes([shapeIndex]);
    this.selectedShapeIndex = shapeIndex;
    this.selectedShapeIndices = new Set([shapeIndex]);
    this.removeMultiSelectionBoxes();
    this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      shape.removeClass('native-powerpoint-shape-selected');
    });

    const selected = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!isSVGGElement(selected)) {
      this.selectedShapeIndex = null;
      this.selectedShapeIndices.clear();
      this.selectedTransform = null;
      this.removeSelectionOverlay();
      return;
    }

    this.selectedTransform = cloneTransform(this.engine.getShapeTransform(selected));
    this.renderInspector();
    this.updateSelectionOverlay();
    this.updateObjectClipboardAvailability();
    this.updateTextToolbar();
  }

  private getTopLevelShapeIndices(): number[] {
    if (!this.svgEl) return [];
    const indices: number[] = [];
    this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      if (shape.parentElement?.closest('g[data-ooxml-shape-idx]')) return;
      const index = getShapeIndex(shape);
      if (isSelectableShapeIndex(index)) indices.push(index);
    });
    return indices;
  }

  private selectAllShapes(): void {
    const indices = this.getTopLevelShapeIndices();
    if (indices.length === 0) return;
    this.applyMultiSelection(indices);
  }






  private applyMultiSelection(indices: number[]): void {
    if (!this.engine || !this.svgEl) return;

    const valid = indices.filter(
      (index) =>
        isSelectableShapeIndex(index) &&
        this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${index}"]`) !== null
    );
    if (valid.length === 0) {
      this.clearSelection();
      return;
    }
    const [first] = valid;
    if (valid.length === 1 && first !== undefined) {
      this.selectShape(first);
      return;
    }

    this.session.selectShapes(valid);
    this.selectedShapeIndex = null;
    this.selectedShapeIndices = new Set(valid);
    this.selectedTransform = null;
    this.applySelectionClasses();
    this.removeSelectionOverlay();
    this.renderInspector();
    this.updateSelectionOverlay();
    this.updateObjectClipboardAvailability();
  }

  private applySelectionClasses(): void {
    this.svgEl?.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      const index = getShapeIndex(shape);
      if (index !== null && this.selectedShapeIndices.has(index)) {
        shape.addClass('native-powerpoint-shape-selected');
      } else {
        shape.removeClass('native-powerpoint-shape-selected');
      }
    });
  }

  private clearSelection(options: { skipTextCommit?: boolean } = {}): void {
    if (!options.skipTextCommit) {
      this.commitActiveTextEditing();
    }
    this.toolbarFormattingSnapshot = null;
    this.textToolbarController.clearFormattingSnapshot();
    this.session.clearSelection();
    this.selectedShapeIndex = null;
    this.selectedShapeIndices.clear();
    this.selectedTransform = null;
    this.svgEl?.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      shape.removeClass('native-powerpoint-shape-selected');
    });
    this.removeSelectionOverlay();
    this.removeMultiSelectionBoxes();
    this.removeMarqueeSelectionPreview();
    this.snapController.clearSnapGuides();
    this.renderInspector();
    this.updateObjectClipboardAvailability();
    this.updateTextToolbar();
  }

  private renderSlideBackgroundControl(container: HTMLElement): void {
    if (!this.engine || this.engine.slideCount === 0) return;

    const section = container.createDiv({ cls: 'native-powerpoint-slide-background' });
    section.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: this.t('powerpoint:inspector.slideBackground') });
    section.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: this.t('powerpoint:inspector.slideBackgroundHint')
    });

    const currentColor = this.engine.getSlideBackgroundColor(this.currentSlide);
    const row = section.createDiv({ cls: 'native-powerpoint-background-row' });
    const colorInput = row.createEl('input', {
      type: 'color',
      cls: 'native-powerpoint-background-color',
      value: currentColor ? `#${currentColor}` : '#ffffff'
    });
    colorInput.disabled = !this.canEdit();

    const applyButton = row.createEl('button', {
      cls: 'native-powerpoint-inspector-button',
      text: this.t('powerpoint:inspector.apply')
    });
    applyButton.disabled = !this.canEdit();
    applyButton.addEventListener('click', () => {
      void this.applySlideBackgroundColor(colorInput.value);
    });
  }

  private async applySlideBackgroundColor(hexColor: string): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('change slide background')) return;

    try {
      const history = await this.captureHistoryEntry('Slide background');
      await this.engine.setSlideBackgroundColor(this.currentSlide, hexColor);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) {
        await this.renderThumbnails();
        this.renderInspector();
      }
      debugLog('inspector', 'Changed PowerPoint slide background', {
        slide: this.currentSlide,
        color: hexColor
      });
    } catch (error) {
      errorLog('inspector', 'PowerPoint slide background change failed', {
        slide: this.currentSlide,
        color: hexColor,
        error
      });
      pptNotice('powerpoint:notice.couldNotChangeSlideBackground', { message: cleanError(error) });
    }
  }

  private renderInspector(): void {
    if (!this.getSettings().showInspector || !this.inspectorEl) return;
    this.inspectorController.render();
  }

  private renderChartDataEditor(chartData: ChartDataGrid): void {
    if (!this.inspectorEl) return;

    const section = this.inspectorEl.createDiv({ cls: 'native-powerpoint-chart-data' });
    section.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: this.t('powerpoint:inspector.chartData') });

    if (!chartData.editable) {
      section.createDiv({
        cls: 'native-powerpoint-inspector-hint',
        text: chartData.reason || this.t('powerpoint:inspector.chartDataReadOnly')
      });
      return;
    }

    section.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: this.t('powerpoint:inspector.chartDataHint')
    });

    const viewport = section.createDiv({ cls: 'native-powerpoint-chart-data-scroll' });
    const table = viewport.createEl('table', { cls: 'native-powerpoint-chart-data-grid' });
    const header = table.createEl('thead').createEl('tr');
    header.createEl('th', { text: chartData.categoryLabel });
    chartData.series.forEach((series) => {
      header.createEl('th', { text: series.name });
      if (series.pointLabels !== null) {
        header.createEl('th', { text: this.t('powerpoint:inspector.seriesLabel', { seriesName: series.name }) });
      }
    });

    const body = table.createEl('tbody');
    const categoryInputs: HTMLInputElement[] = [];
    const valueInputs = chartData.series.map(() => [] as HTMLInputElement[]);
    const pointLabelInputs = chartData.series.map(() => [] as HTMLInputElement[]);

    chartData.categories.forEach((category, rowIndex) => {
      const row = body.createEl('tr');
      categoryInputs.push(this.createChartDataInput(row, category));

      chartData.series.forEach((series, seriesIndex) => {
        valueInputs[seriesIndex]?.push(this.createChartDataInput(row, series.values[rowIndex] ?? '', true));
        if (series.pointLabels !== null) {
          pointLabelInputs[seriesIndex]?.push(
            this.createChartDataInput(row, series.pointLabels[rowIndex] ?? '')
          );
        }
      });
    });

    const apply = section.createEl('button', {
      cls: 'native-powerpoint-inspector-button',
      text: this.t('powerpoint:inspector.applyChartData')
    });
    apply.disabled = !this.canEdit();
    apply.addEventListener('click', () => {
      const update: ChartDataUpdate = {
        categories: categoryInputs.map((input) => input.value),
        series: chartData.series.map((series, index) => ({
          values: valueInputs[index]?.map((input) => input.value) ?? [],
          pointLabels: series.pointLabels === null
            ? null
            : pointLabelInputs[index]?.map((input) => input.value) ?? []
        }))
      };
      void this.applyChartData(update);
    });
  }

  private createChartDataInput(row: HTMLTableRowElement, value: string, numeric = false): HTMLInputElement {
    const input = row.createEl('td').createEl('input', {
      type: 'text',
      value,
      attr: numeric ? { inputmode: 'decimal' } : {}
    });
    input.disabled = !this.canEdit();
    return input;
  }

  private async applyChartData(update: ChartDataUpdate): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('edit chart data')) return;

    const chartShapeIndex = this.selectedShapeIndex;
    try {
      const history = await this.captureHistoryEntry('Edit chart data');
      await this.engine.updateChartData(this.currentSlide, chartShapeIndex, update);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderEditedShape(chartShapeIndex);
      if (rendered) await this.renderThumbnails();
      debugLog('inspector', 'Updated PowerPoint chart data', {
        slide: this.currentSlide,
        shapeIndex: chartShapeIndex,
        categoryCount: update.categories.length,
        seriesCount: update.series.length
      });
    } catch (error) {
      errorLog('inspector', 'PowerPoint chart-data update failed', {
        slide: this.currentSlide,
        shapeIndex: chartShapeIndex,
        categoryCount: update.categories.length,
        seriesCount: update.series.length,
        error
      });
      pptNotice('powerpoint:notice.couldNotUpdateChartData', { message: cleanError(error) });
    }
  }

  private createNumberField(container: HTMLElement, label: string, value: number): HTMLInputElement {
    const wrapper = container.createDiv({ cls: 'native-powerpoint-field' });
    wrapper.createEl('label', { text: label });
    const input = wrapper.createEl('input', {
      type: 'number',
      value: String(Math.round(value * 100) / 100)
    });
    return input;
  }

  private renderViewOnlyWarning(container: HTMLElement): void {
    if (!this.isViewOnly || !this.viewOnlyReason) return;

    container.createDiv({
      cls: 'native-powerpoint-view-only-warning',
      text: this.viewOnlyReason
    });
  }

  private async applyTextValue(
    text: string,
    target: TextEditTarget | null = null,
    slideIndex = this.currentSlide,
    options: { authoritativePreviousText?: string | null } = {},
  ): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('edit text')) return;

    const shapeIndex = target?.shapeIndex ?? this.selectedShapeIndex;
    if (shapeIndex === null) {
      debugLog('text-edit', 'applyTextValue skipped (no shape target)', { slideIndex });
      return;
    }

    const applyStarted = performance.now();
    debugLog('text-edit', 'applyTextValue start', {
      slideIndex,
      currentSlide: this.currentSlide,
      shapeIndex,
      textLength: text.length
    });

    try {
      const { previousText, source: previousSource } = previousTextForInlineApply({
        sessionBaseline: options.authoritativePreviousText,
        targetText: target?.text,
        liveSvgText: this.getSelectedShapeElement()?.textContent?.trim() ?? null,
      });
      if (text === previousText) {
        debugLog('text-edit', 'applyTextValue skipped (unchanged text)', {
          slideIndex,
          shapeIndex,
          previousSource,
          textLength: text.length,
        });
        return;
      }

      const history = await this.captureHistoryEntry('Edit text');
      const scrollPosition = this.captureCanvasScroll();
      if (target?.kind === 'shape-paragraph') {
        await this.engine.updateParagraphText(
          slideIndex,
          target.shapeIndex,
          target.paragraphIndex,
          text
        );
      } else if (target) {
        await this.engine.updateGeneratedText(slideIndex, target.shapeIndex, target, text);
      } else {
        await this.engine.updateShapeText(slideIndex, shapeIndex, text);
      }
      this.recordHistoryEntry(history);
      this.markDirty();

      if (slideIndex !== this.currentSlide) {
        debugLog('text-edit', 'applyTextValue skipped render (slide changed during commit)', {
          slideIndex,
          currentSlide: this.currentSlide
        });
        return;
      }

      const rendered = await this.renderEditedShape(shapeIndex);
      if (rendered) {
        this.restoreCanvasScrollSoon(scrollPosition);
      }
      if (rendered) await this.renderThumbnails();

      const applyMs = Math.round(performance.now() - applyStarted);
      debugLog('text-edit', 'applyTextValue complete', {
        slideIndex,
        shapeIndex,
        ms: applyMs,
        previousSource,
        textLength: text.length,
        previousTextLength: previousText.length,
      });
      if (applyMs > 1000) {
        warnLog('text-edit', 'slow applyTextValue', { slideIndex, shapeIndex, ms: applyMs });
      }
    } catch (error) {
      errorLog('text-edit', 'applyTextValue failed', { slideIndex, shapeIndex, error });
      pptNotice('powerpoint:notice.couldNotUpdateText', { message: cleanError(error) });
    }
  }

  /**
   * Turn bare Enter into a native DrawingML paragraph split. Shift+Enter stays
   * in the textarea's normal path and is intentionally persisted as `<a:br/>`.
   */
  private startInlineParagraphSplit(editor: HTMLTextAreaElement, target: ShapeTextEditTarget): void {
    if (this.paragraphSplitPromise || this.paragraphRemovalPromise || this.rangeDeletionPromise) {
      debugLog('text-edit', 'Ignored repeated paragraph split while one is pending', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
      });
      return;
    }

    const operation = this.splitInlineParagraph(editor, target);
    this.paragraphSplitPromise = operation;
    void operation.finally(() => {
      if (this.paragraphSplitPromise === operation) {
        this.paragraphSplitPromise = null;
      }
    });
  }

  /** Start a structural Backspace operation without flattening `<a:p>` boundaries into a newline. */
  private startInlineEmptyParagraphRemoval(editor: HTMLTextAreaElement, target: ShapeTextEditTarget): void {
    if (this.paragraphSplitPromise || this.paragraphRemovalPromise || this.rangeDeletionPromise) {
      debugLog('text-edit', 'Ignored paragraph-boundary Backspace while a structural text mutation is pending', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
      });
      return;
    }

    const operation = this.removeInlineEmptyPrecedingParagraph(editor, target);
    this.paragraphRemovalPromise = operation;
    void operation.finally(() => {
      if (this.paragraphRemovalPromise === operation) {
        this.paragraphRemovalPromise = null;
      }
    });
  }

  /** Start Backspace's structural merge when the preceding paragraph contains text. */
  private startInlinePrecedingParagraphMerge(editor: HTMLTextAreaElement, target: ShapeTextEditTarget): void {
    if (this.paragraphSplitPromise || this.paragraphRemovalPromise || this.rangeDeletionPromise) {
      debugLog('text-edit', 'Ignored paragraph-boundary merge while a structural text mutation is pending', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
      });
      return;
    }

    const operation = this.mergeInlinePrecedingParagraph(editor, target);
    this.paragraphRemovalPromise = operation;
    void operation.finally(() => {
      if (this.paragraphRemovalPromise === operation) {
        this.paragraphRemovalPromise = null;
      }
    });
  }

  /** Commit a visual multi-paragraph selection through the authoritative OOXML text model. */
  private startInlineRangeDeletion(editor: HTMLTextAreaElement, target: ShapeTextEditTarget): void {
    const selection = this.inlineRangeSelection;
    if (!selection || selection.shapeIndex !== target.shapeIndex || selection.ranges.length === 0) return;
    if (this.paragraphSplitPromise || this.paragraphRemovalPromise || this.rangeDeletionPromise) {
      debugLog('text-edit', 'Ignored repeated inline range deletion while a structural text mutation is pending', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
      });
      return;
    }

    const operation = this.deleteInlineTextRange(editor, target, selection.ranges);
    this.rangeDeletionPromise = operation;
    void operation.finally(() => {
      if (this.rangeDeletionPromise === operation) {
        this.rangeDeletionPromise = null;
      }
    });
  }

  private async deleteInlineTextRange(
    editor: HTMLTextAreaElement,
    target: ShapeTextEditTarget,
    editorRanges: ParagraphTextRange[],
  ): Promise<void> {
    const engine = this.engine;
    if (!engine || this.activeEditor !== editor || !this.ensureEditable('delete text')) return;

    const slideIndex = this.currentSlide;
    const ranges = this.mapRangesToOoxmlOffsets(target.shapeIndex, editorRanges);
    const scrollPosition = this.captureCanvasScroll();
    editor.readOnly = true;
    logPptxAction('text-edit', 'delete-text-ranges', {
      slide: slideIndex,
      shapeIndex: target.shapeIndex,
      editorRangeCount: editorRanges.length,
      rangeCount: ranges.length,
      paragraphIndexes: [...new Set(ranges.map((range) => range.paragraphIndex))],
    });

    try {
      const history = await this.captureHistoryEntry('Delete text');
      const result: TextRangeDeletionResult = await engine.deleteTextRanges(slideIndex, target.shapeIndex, ranges);
      if (!result.changed) {
        debugLog('text-edit', 'Skipped inline range deletion without text', {
          slide: slideIndex,
          shapeIndex: target.shapeIndex,
          rangeCount: ranges.length,
        });
        return;
      }
      this.recordHistoryEntry(history);
      this.markDirty();

      if (slideIndex !== this.currentSlide || this.activeEditor !== editor) {
        debugLog('text-edit', 'Inline range deletion committed after editor or slide changed', {
          slide: slideIndex,
          currentSlide: this.currentSlide,
          shapeIndex: target.shapeIndex,
          paragraphIndex: result.paragraphIndex,
        });
        return;
      }

      target.paragraphIndex = result.paragraphIndex;
      this.activeEditorTextDirty = false;
      const rendered = await this.renderEditedShape(target.shapeIndex);
      if (!rendered) {
        throw new Error('Could not re-render after deleting the selected PowerPoint text.');
      }
      this.restoreCanvasScrollSoon(scrollPosition);
      await this.renderThumbnails();

      if (!this.refreshActiveShapeEditorAfterRender()) {
        throw new Error('Could not restore the inline editor after deleting the selected PowerPoint text.');
      }
      const caretOffset = Math.max(0, Math.min(result.caretOffset, target.text.length));
      this.setInlineEditorValue(editor, target.text);
      editor.setSelectionRange(caretOffset, caretOffset);
      this.clearWholeShapeInlineSelection();
      this.resetInlineEditorScroll(editor);
      this.rememberInlineCaretPlacement(editor, target.element, caretOffset);
      this.refreshInlineEditorGeometry();
      this.updateTextToolbar();
      debugLog('text-edit', 'Deleted inline PowerPoint text selection', {
        slide: slideIndex,
        shapeIndex: target.shapeIndex,
        rangeCount: result.deletedRangeCount,
        paragraphIndex: result.paragraphIndex,
        caretOffset,
        removedParagraphCount: result.removedParagraphCount,
        mergedParagraphs: result.mergedParagraphs,
      });
    } catch (error) {
      errorLog('text-edit', 'Inline range deletion failed', {
        slide: slideIndex,
        shapeIndex: target.shapeIndex,
        rangeCount: ranges.length,
        error: cleanError(error),
      });
      pptNotice('powerpoint:notice.couldNotUpdateText', { message: cleanError(error) });
    } finally {
      if (this.activeEditor === editor) {
        editor.readOnly = false;
      }
    }
  }

  private async removeInlineEmptyPrecedingParagraph(
    editor: HTMLTextAreaElement,
    target: ShapeTextEditTarget,
  ): Promise<void> {
    const engine = this.engine;
    if (!engine || this.activeEditor !== editor || !this.ensureEditable('delete empty paragraph')) return;

    const slideIndex = this.currentSlide;
    const sourceParagraphIndex = target.paragraphIndex;
    const pendingText = editor.value;
    const hasPendingText = this.activeEditorTextDirty;
    const scrollPosition = this.captureCanvasScroll();
    editor.readOnly = true;

    logPptxAction('text-edit', 'remove-empty-preceding-paragraph', {
      slide: slideIndex,
      shapeIndex: target.shapeIndex,
      paragraphIndex: sourceParagraphIndex,
      hasPendingText,
      pendingTextLength: pendingText.length,
    });

    try {
      const history = await this.captureHistoryEntry('Delete empty paragraph');
      const response = await this.session.applyCommand({
        type: 'remove-empty-preceding-paragraph',
        slideIndex,
        shapeIndex: target.shapeIndex,
        paragraphIndex: sourceParagraphIndex,
      });
      const result = response as EmptyPrecedingParagraphRemovalResult;
      if (!result.removed) {
        debugLog('text-edit', 'Inline Backspace kept preceding paragraph', {
          slide: slideIndex,
          shapeIndex: target.shapeIndex,
          paragraphIndex: sourceParagraphIndex,
          reason: result.reason,
        });
        return;
      }
      this.recordHistoryEntry(history);

      if (slideIndex !== this.currentSlide || this.activeEditor !== editor) {
        debugLog('text-edit', 'Empty paragraph removal committed after editor or slide changed', {
          slide: slideIndex,
          currentSlide: this.currentSlide,
          shapeIndex: target.shapeIndex,
          sourceParagraphIndex,
          paragraphIndex: result.paragraphIndex,
        });
        return;
      }

      target.paragraphIndex = result.paragraphIndex;
      const rendered = await this.renderEditedShape(target.shapeIndex);
      if (!rendered) {
        throw new Error('Could not re-render after deleting the empty PowerPoint paragraph.');
      }
      this.restoreCanvasScrollSoon(scrollPosition);
      await this.renderThumbnails();

      if (!this.refreshActiveShapeEditorAfterRender()) {
        throw new Error('Could not restore the inline editor after deleting the empty PowerPoint paragraph.');
      }
      this.setInlineEditorValue(editor, hasPendingText ? pendingText : target.text);
      if (hasPendingText) {
        this.syncShapeParagraphPreview(target, pendingText);
      }
      editor.setSelectionRange(0, 0);
      this.clearWholeShapeInlineSelection();
      this.resetInlineEditorScroll(editor);
      this.rememberInlineCaretPlacement(editor, target.element, 0);
      this.refreshInlineEditorGeometry();
      this.updateTextToolbar();
      debugLog('text-edit', 'Inline Backspace removed preceding empty paragraph', {
        slide: slideIndex,
        shapeIndex: target.shapeIndex,
        sourceParagraphIndex,
        paragraphIndex: result.paragraphIndex,
        beforeParagraphCount: result.beforeParagraphCount,
        afterParagraphCount: result.afterParagraphCount,
        retainedPendingText: hasPendingText,
      });
    } catch (error) {
      errorLog('text-edit', 'Inline Backspace empty paragraph removal failed', {
        slide: slideIndex,
        shapeIndex: target.shapeIndex,
        paragraphIndex: sourceParagraphIndex,
        error: cleanError(error),
      });
      pptNotice('powerpoint:notice.couldNotUpdateText', { message: cleanError(error) });
    } finally {
      if (this.activeEditor === editor) {
        editor.readOnly = false;
      }
    }
  }

  private async mergeInlinePrecedingParagraph(
    editor: HTMLTextAreaElement,
    target: ShapeTextEditTarget,
  ): Promise<void> {
    const engine = this.engine;
    if (!engine || this.activeEditor !== editor || !this.ensureEditable('merge paragraphs')) return;

    const slideIndex = this.currentSlide;
    const sourceParagraphIndex = target.paragraphIndex;
    const pendingText = editor.value;
    const normalizedText = this.paragraphEditorTextFromDom(
      target.shapeIndex,
      sourceParagraphIndex,
      pendingText,
    );
    const hasPendingText = this.activeEditorTextDirty;
    const scrollPosition = this.captureCanvasScroll();
    editor.readOnly = true;

    logPptxAction('text-edit', 'merge-preceding-paragraph', {
      slide: slideIndex,
      shapeIndex: target.shapeIndex,
      paragraphIndex: sourceParagraphIndex,
      hasPendingText,
      pendingTextLength: pendingText.length,
      normalizedTextLength: normalizedText.length,
    });

    try {
      const history = await this.captureHistoryEntry('Merge paragraphs');
      const response = await this.session.applyCommand({
        type: 'merge-preceding-paragraph',
        slideIndex,
        shapeIndex: target.shapeIndex,
        paragraphIndex: sourceParagraphIndex,
        text: hasPendingText ? normalizedText : undefined,
      });
      const result = response as PrecedingParagraphMergeResult;
      if (!result.merged) {
        debugLog('text-edit', 'Inline Backspace kept paragraph boundary', {
          slide: slideIndex,
          shapeIndex: target.shapeIndex,
          paragraphIndex: sourceParagraphIndex,
          reason: result.reason,
        });
        return;
      }
      this.recordHistoryEntry(history);

      if (slideIndex !== this.currentSlide || this.activeEditor !== editor) {
        debugLog('text-edit', 'Paragraph merge committed after editor or slide changed', {
          slide: slideIndex,
          currentSlide: this.currentSlide,
          shapeIndex: target.shapeIndex,
          sourceParagraphIndex,
          paragraphIndex: result.paragraphIndex,
        });
        return;
      }

      target.paragraphIndex = result.paragraphIndex;
      const rendered = await this.renderEditedShape(target.shapeIndex);
      if (!rendered) {
        throw new Error('Could not re-render after merging PowerPoint paragraphs.');
      }
      this.restoreCanvasScrollSoon(scrollPosition);
      await this.renderThumbnails();

      if (!this.refreshActiveShapeEditorAfterRender()) {
        throw new Error('Could not restore the inline editor after merging PowerPoint paragraphs.');
      }
      const caretOffset = Math.max(0, Math.min(result.caretOffset, target.text.length));
      this.setInlineEditorValue(editor, target.text);
      editor.setSelectionRange(caretOffset, caretOffset);
      this.activeEditorTextDirty = false;
      this.clearWholeShapeInlineSelection();
      this.resetInlineEditorScroll(editor);
      this.rememberInlineCaretPlacement(editor, target.element, caretOffset);
      this.refreshInlineEditorGeometry();
      this.updateTextToolbar();
      debugLog('text-edit', 'Inline Backspace merged preceding paragraph', {
        slide: slideIndex,
        shapeIndex: target.shapeIndex,
        sourceParagraphIndex,
        paragraphIndex: result.paragraphIndex,
        caretOffset,
        beforeParagraphCount: result.beforeParagraphCount,
        afterParagraphCount: result.afterParagraphCount,
        retainedPendingText: hasPendingText,
      });
    } catch (error) {
      errorLog('text-edit', 'Inline Backspace paragraph merge failed', {
        slide: slideIndex,
        shapeIndex: target.shapeIndex,
        paragraphIndex: sourceParagraphIndex,
        error: cleanError(error),
      });
      pptNotice('powerpoint:notice.couldNotUpdateText', { message: cleanError(error) });
    } finally {
      if (this.activeEditor === editor) {
        editor.readOnly = false;
      }
    }
  }

  private async splitInlineParagraph(editor: HTMLTextAreaElement, target: ShapeTextEditTarget): Promise<void> {
    const engine = this.engine;
    if (!engine || this.activeEditor !== editor || !this.ensureEditable('split paragraph')) return;

    const rawText = editor.value;
    const rawStart = Math.max(0, Math.min(editor.selectionStart ?? rawText.length, rawText.length));
    const rawEnd = Math.max(rawStart, Math.min(editor.selectionEnd ?? rawStart, rawText.length));
    const textAfterSelection = rawText.slice(0, rawStart) + rawText.slice(rawEnd);
    const normalizedText = this.paragraphEditorTextFromDom(
      target.shapeIndex,
      target.paragraphIndex,
      textAfterSelection,
    );
    const hasPendingText = this.activeEditorTextDirty || rawStart !== rawEnd;
    const ooxmlText = engine.getParagraphRunText(
      this.currentSlide,
      target.shapeIndex,
      target.paragraphIndex,
    ) ?? target.text;
    const splitOffset = hasPendingText
      ? mapEditorOffsetToOoxmlOffset(textAfterSelection, normalizedText, rawStart, false)
      : mapEditorOffsetToOoxmlOffset(rawText, ooxmlText, rawStart, false);
    const slideIndex = this.currentSlide;
    const scrollPosition = this.captureCanvasScroll();

    logPptxAction('text-edit', 'split-paragraph', {
      slide: slideIndex,
      shapeIndex: target.shapeIndex,
      paragraphIndex: target.paragraphIndex,
      selectionStart: rawStart,
      selectionEnd: rawEnd,
      splitOffset,
      textLength: rawText.length,
      normalizedTextLength: normalizedText.length,
      hasPendingText,
    });

    try {
      const history = await this.captureHistoryEntry('Split paragraph');
      const response = await this.session.applyCommand({
        type: 'split-paragraph',
        slideIndex,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
        splitOffset,
        text: hasPendingText ? normalizedText : undefined,
      });
      const result = response as ParagraphSplitResult;
      this.recordHistoryEntry(history);

      if (slideIndex !== this.currentSlide || this.activeEditor !== editor) {
        debugLog('text-edit', 'Paragraph split committed after editor or slide changed', {
          slide: slideIndex,
          currentSlide: this.currentSlide,
          shapeIndex: target.shapeIndex,
          paragraphIndex: target.paragraphIndex,
          insertedParagraphIndex: result.paragraphIndex,
        });
        return;
      }

      const rendered = await this.renderEditedShape(target.shapeIndex);
      if (!rendered) {
        throw new Error('Could not re-render the split PowerPoint paragraph.');
      }
      this.restoreCanvasScrollSoon(scrollPosition);
      await this.renderThumbnails();

      target.paragraphIndex = result.paragraphIndex;
      if (!this.refreshActiveShapeEditorAfterRender()) {
        throw new Error('Could not restore the inline editor for the new paragraph.');
      }
      this.setInlineEditorValue(editor, target.text);
      editor.setSelectionRange(0, 0);
      this.activeEditorTextDirty = false;
      this.clearWholeShapeInlineSelection();
      this.resetInlineEditorScroll(editor);
      this.rememberInlineCaretPlacement(editor, target.element, 0);
      this.refreshInlineEditorGeometry();
      this.logParagraphSplitLayout(slideIndex, target.shapeIndex, result.paragraphIndex, result);
      this.updateTextToolbar();
    } catch (error) {
      errorLog('text-edit', 'PowerPoint paragraph split failed', {
        slide: slideIndex,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
        error: cleanError(error),
      });
      pptNotice('powerpoint:notice.couldNotUpdateText', { message: cleanError(error) });
    }
  }

  /** Durable geometry breadcrumb for paragraph-split regressions. */
  private logParagraphSplitLayout(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    result: ParagraphSplitResult,
  ): void {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    const text = shape?.querySelector('text');
    const frameBox = isSVGGElement(shape) && this.engine
      ? this.getTransformSelectionBox(this.engine.getShapeTransform(shape))
      : null;
    const textBox = isSVGTextElement(text) ? this.getElementBox(text) : null;
    const paragraphTarget = this.buildParagraphEditTarget(shapeIndex, paragraphIndex);
    const paragraphBox = paragraphTarget ? this.getElementBox(paragraphTarget.element) : null;
    const visualLineCount = this.getRunLineContainers(shapeIndex, paragraphIndex).length;
    const overflowRight = frameBox && textBox
      ? textBox.left + textBox.width > frameBox.left + frameBox.width + 0.5
      : null;
    const overflowBottom = frameBox && textBox
      ? textBox.top + textBox.height > frameBox.top + frameBox.height + 0.5
      : null;
    const data = {
      slide: slideIndex,
      shapeIndex,
      paragraphIndex,
      paragraphCount: result.afterParagraphCount,
      visualLineCount,
      frameBox,
      textBox,
      paragraphBox,
      overflowRight,
      overflowBottom,
      removedSoftBreaks: result.removedSoftBreaks,
    };

    if (!frameBox || !textBox || !paragraphBox || overflowRight || overflowBottom) {
      warnLog('text-edit', 'Paragraph split layout needs attention', data);
      return;
    }
    debugLog('text-edit', 'Paragraph split layout verified', data);
  }

  private async applyInspectorTransform(): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null || !this.selectedTransform) return;
    if (!this.ensureEditable('edit layout')) return;

    const transform = cloneTransform(this.selectedTransform);
    transform.x = this.engine.pxToEmu(Number(this.xInput?.value || 0));
    transform.y = this.engine.pxToEmu(Number(this.yInput?.value || 0));
    transform.cx = this.engine.pxToEmu(Math.max(1, Number(this.widthInput?.value || 1)));
    transform.cy = this.engine.pxToEmu(Math.max(1, Number(this.heightInput?.value || 1)));
    transform.rot = this.engine.degreesToOoxml(Number(this.rotationInput?.value || 0));
    await this.commitTransform(transform);
  }

  private handleInlineTextPointerDown(event: PointerEvent, target: Element): void {
    const textTarget = target.closest('text') && target.closest(GENERATED_GRID_SELECTOR)
      ? this.getGeneratedTextEditTarget(target)
      : this.getTextEditTarget(target);

    event.preventDefault();
    event.stopPropagation();
    this.clearBrowserTextSelection();

    if (!textTarget) {
      if (target.closest(GENERATED_GRID_SELECTOR)) {
        this.showGeneratedTextNotice();
      }
      return;
    }
    if (!this.ensureEditable('edit text')) return;

    this.suppressNextTextClick = true;
    this.selectShapeForTextEditing(textTarget.shapeIndex);
    this.startTextEditor(textTarget, event.clientX, event.clientY);

    const editor = this.activeEditor;
    if (!editor || this.activeEditorTarget !== textTarget.element) return;

    const box = this.getElementBox(textTarget.element);
    if (!box) return;

    this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(textTarget.element, event.clientY, box);
    const offset = this.getInlineTextOffsetAtClientPoint(textTarget.element, editor, event.clientX, event.clientY, box);
    this.focusEditorWithoutCanvasScroll(editor);

    const paragraphIndex = this.getParagraphIndexFromInlineElement(textTarget.element);
    const runContainers = paragraphIndex !== null
      ? this.getRunLineContainers(textTarget.shapeIndex, paragraphIndex)
      : [];
    debugLog('text-select', 'inline drag start', {
      slide: this.currentSlide,
      shapeIndex: textTarget.shapeIndex,
      paragraphIndex,
      anchorOffset: offset,
      pointer: { x: event.clientX, y: event.clientY },
      targetTag: textTarget.element.tagName.toLowerCase(),
      targetText: (textTarget.element.textContent ?? '').slice(0, 80),
      editorLength: editor.value.length,
      runLineCount: runContainers.length,
      runLineCharCounts: runContainers.map((container) => this.getRunCharInfo(container).total),
    });
    editor.setSelectionRange(offset, offset);
    this.rememberInlineCaretPlacement(editor, textTarget.element, offset);
    this.resetInlineEditorScroll(editor);
    this.updateInlineCaret(editor, textTarget.element);
    this.beginInlineSelectionDrag(editor, textTarget.element, offset, event);
  }

  private applyInlineMultiClickSelectionAtPoint(
    target: TextEditTarget,
    clientX: number,
    clientY: number,
    clickCount: number
  ): boolean {
    const editor = this.activeEditor;
    if (!editor || this.activeEditorTarget !== target.element) return false;

    const box = this.getElementBox(target.element);
    if (!box) return false;

    this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(target.element, clientY, box);
    const offset = this.getInlineTextOffsetAtClientPoint(target.element, editor, clientX, clientY, box);
    this.focusEditorWithoutCanvasScroll(editor);
    const handled = this.applyInlineMultiClickSelection(editor, target.element, offset, clickCount);
    if (handled) this.resetInlineEditorScroll(editor);
    return handled;
  }

  private applyInlineMultiClickSelection(
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement,
    offset: number,
    clickCount: number
  ): boolean {
    if (clickCount < 2) return false;

    this.stopInlineSelectionDrag();
    if (clickCount >= 3) {
      debugLog('text-select', 'triple-click select all', {
        slide: this.currentSlide,
        shapeIndex: this.activeShapeTextTarget?.shapeIndex ?? this.selectedShapeIndex
      });
      this.selectAllInlineText(editor, element);
      this.updateTextToolbar();
      return true;
    }

    this.clearWholeShapeInlineSelection();
    const range = getInlineWordRange(editor.value, offset);
    editor.setSelectionRange(range.start, range.end);
    const target = this.activeShapeTextTarget;
    if (target) {
      this.inlineRangeSelection = {
        shapeIndex: target.shapeIndex,
        ranges: [{ paragraphIndex: target.paragraphIndex, start: range.start, end: range.end }]
      };
    }
    this.lastInlineCaretPlacement = null;
    this.updateInlineCaret(editor, element);
    debugLog('text-select', 'double-click select word', {
      slide: this.currentSlide,
      shapeIndex: this.activeShapeTextTarget?.shapeIndex ?? this.selectedShapeIndex,
      start: range.start,
      end: range.end
    });
    this.updateTextToolbar();
    return true;
  }

  private beginInlineSelectionDrag(
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement,
    anchorOffset: number,
    event: PointerEvent
  ): void {
    this.stopInlineSelectionDrag();
    this.clearWholeShapeInlineSelection();
    const capturedPointer = this.captureInlineSelectionPointer(event.pointerId);

    // Pointermove fires far faster than we can run the (expensive) SVG glyph
    // measurement, so coalesce to one selection update per animation frame using
    // the most recent pointer position. This keeps dragging smooth and ensures
    // the final position is always processed instead of dropped mid-flood.
    const flushDragFrame = () => {
      const drag = this.inlineSelectionDrag;
      if (!drag) return;
      drag.pendingFrame = null;
      this.clearBrowserTextSelection();
      this.extendInlineSelectionDrag(drag.pendingClientX, drag.pendingClientY);
    };

    const onPointerMove = (moveEvent: PointerEvent) => {
      if (moveEvent.pointerId !== event.pointerId) return;
      moveEvent.preventDefault();
      moveEvent.stopPropagation();
      const drag = this.inlineSelectionDrag;
      if (!drag) return;
      drag.pendingClientX = moveEvent.clientX;
      drag.pendingClientY = moveEvent.clientY;
      if (drag.pendingFrame === null) {
        drag.pendingFrame = window.requestAnimationFrame(flushDragFrame);
      }
    };
    const onPointerUp = (upEvent: PointerEvent) => {
      if (upEvent.pointerId !== event.pointerId) return;
      upEvent.preventDefault();
      upEvent.stopPropagation();
      this.clearBrowserTextSelection();
      this.extendInlineSelectionDrag(upEvent.clientX, upEvent.clientY);
      const drag = this.inlineSelectionDrag;
      debugLog('text-select', 'inline drag ended', {
        slide: this.currentSlide,
        shapeIndex: this.activeShapeTextTarget?.shapeIndex ?? this.selectedShapeIndex,
        paragraphIndex: this.activeShapeTextTarget?.paragraphIndex ?? null,
        pointerId: upEvent.pointerId,
        eventType: upEvent.type,
        pointer: { x: upEvent.clientX, y: upEvent.clientY },
        isSelecting: drag?.isSelecting ?? false,
        capturedPointer,
        selectionStart: editor.selectionStart ?? 0,
        selectionEnd: editor.selectionEnd ?? 0,
      });
      this.stopInlineSelectionDrag();
    };
    const cleanup = () => {
      if (this.inlineSelectionDrag?.pendingFrame !== null && this.inlineSelectionDrag) {
        window.cancelAnimationFrame(this.inlineSelectionDrag.pendingFrame);
      }
      activeDocument.removeEventListener('pointermove', onPointerMove, true);
      activeDocument.removeEventListener('pointerup', onPointerUp, true);
      activeDocument.removeEventListener('pointercancel', onPointerUp, true);
      this.releaseInlineSelectionPointer(event.pointerId);
    };

    this.inlineSelectionDrag = {
      editor,
      element,
      anchorOffset,
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      isSelecting: false,
      pendingFrame: null,
      pendingClientX: event.clientX,
      pendingClientY: event.clientY,
      lastLogKey: null,
      lastLogAt: 0,
      cleanup
    };
    activeDocument.addEventListener('pointermove', onPointerMove, true);
    activeDocument.addEventListener('pointerup', onPointerUp, true);
    activeDocument.addEventListener('pointercancel', onPointerUp, true);
  }

  /** Keep receiving the drag sequence after the pointer leaves a text glyph. */
  private captureInlineSelectionPointer(pointerId: number): boolean {
    const svg = this.svgEl;
    if (!svg || typeof svg.setPointerCapture !== 'function') return false;
    try {
      svg.setPointerCapture(pointerId);
      return true;
    } catch {
      return false;
    }
  }

  /** Release the SVG capture once the selection ends or is cancelled. */
  private releaseInlineSelectionPointer(pointerId: number): void {
    const svg = this.svgEl;
    if (!svg || typeof svg.releasePointerCapture !== 'function') return;
    try {
      if (typeof svg.hasPointerCapture !== 'function' || svg.hasPointerCapture(pointerId)) {
        svg.releasePointerCapture(pointerId);
      }
    } catch {
      // The browser may implicitly release capture before a pointercancel.
    }
  }

  private extendInlineSelectionDrag(clientX: number, clientY: number): void {
    const drag = this.inlineSelectionDrag;
    if (!drag || this.activeEditor !== drag.editor || this.activeEditorTarget !== drag.element) return;

    const box = this.getElementBox(drag.element);
    if (!box) return;

    if (!drag.isSelecting) {
      if (!this.hasInlineSelectionDragMoved(drag, clientX, clientY)) {
        drag.editor.setSelectionRange(drag.anchorOffset, drag.anchorOffset);
        this.rememberInlineCaretPlacement(drag.editor, drag.element, drag.anchorOffset);
        this.updateInlineCaret(drag.editor, drag.element);
        return;
      }
      drag.isSelecting = true;
      this.lastInlineCaretPlacement = null;
      debugLog('text-select', 'inline drag threshold crossed', {
        slide: this.currentSlide,
        anchorOffset: drag.anchorOffset,
        startPointer: { x: drag.startClientX, y: drag.startClientY },
        pointer: { x: clientX, y: clientY },
      });
    }

    // Determine which paragraph in the text box the pointer is currently over so a
    // drag can extend the selection across paragraph boundaries (not just within
    // the paragraph the drag started in).
    const shape = this.getSelectedShapeElement();
    const paragraphs = shape ? this.getShapeTextParagraphs(shape) : [];
    const anchorIndex = paragraphs.indexOf(drag.element);
    const focusParagraph = anchorIndex >= 0
      ? this.getDragFocusParagraph(paragraphs, clientY, drag.element)
      : drag.element;
    const focusIndex = paragraphs.indexOf(focusParagraph);

    // A single OOXML paragraph can wrap into several visual-line containers (and
    // bullet/number markers render as their own container). Those share the same
    // data-ooxml-para-idx. The editor's selection model already spans the whole
    // paragraph with flat run-only offsets, so a drag inside the same paragraph —
    // even across wrapped lines — must stay on the single-paragraph path. Routing
    // it through the cross-paragraph path would treat the flat anchor offset as a
    // per-line local offset and re-add the bullet length, shifting the selection.
    const anchorParagraphIndex = this.getParagraphIndexFromInlineElement(drag.element);
    const focusParagraphIndex = this.getParagraphIndexFromInlineElement(focusParagraph);
    const sameOoxmlParagraph = focusParagraph === drag.element
      || (anchorParagraphIndex !== null && anchorParagraphIndex === focusParagraphIndex);

    if (anchorIndex < 0 || focusIndex < 0 || sameOoxmlParagraph) {
      // Single-paragraph selection: keep using the editor's native selection model.
      this.clearWholeShapeInlineSelection();
      this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(drag.element, clientY, box);
      const focusOffset = this.getInlineTextOffsetAtClientPoint(drag.element, drag.editor, clientX, clientY, box);
      const selectionStart = Math.min(drag.anchorOffset, focusOffset);
      const selectionEnd = Math.max(drag.anchorOffset, focusOffset);
      const direction = focusOffset < drag.anchorOffset ? 'backward' : 'forward';
      this.focusEditorWithoutCanvasScroll(drag.editor);
      drag.editor.setSelectionRange(selectionStart, selectionEnd, direction);
      this.resetInlineEditorScroll(drag.editor);
      this.updateInlineCaret(drag.editor, drag.element);
      const shapeIndex = this.activeShapeTextTarget?.shapeIndex ?? this.selectedShapeIndex ?? -1;
      const paragraphIndex = this.getParagraphIndexFromInlineElement(drag.element) ?? 0;
      const runContainers = this.getRunLineContainers(shapeIndex, paragraphIndex);
      const logKey = `single:${shapeIndex}:${paragraphIndex}:${focusOffset}:${selectionStart}:${selectionEnd}`;
      const now = performance.now();
      if (drag.lastLogKey !== logKey && now - (drag.lastLogAt ?? 0) >= 200) {
        drag.lastLogKey = logKey;
        drag.lastLogAt = now;
        debugLog('text-select', 'single-paragraph drag selection', {
          slide: this.currentSlide,
          shapeIndex,
          paragraphIndex,
          anchorOffset: drag.anchorOffset,
          focusOffset,
          selectionStart,
          selectionEnd,
          direction,
          pointer: { x: clientX, y: clientY },
          selectedText: drag.editor.value.slice(selectionStart, selectionEnd),
          runLineCount: runContainers.length,
          runLineCharCounts: runContainers.map((container) => this.getRunCharInfo(container).total),
          leafLineCharCounts: runContainers.map((container) => this.getLeafCharInfo(container).total)
        });
      }
      return;
    }

    // Cross-paragraph selection.
    const focusBox = this.getElementBox(focusParagraph) ?? box;
    const focusGeometryOffset = this.getInlineTextOffsetAtClientPointForElement(
      focusParagraph,
      clientX,
      clientY,
      focusBox
    );
    const focusRunTotal = this.getRunCharInfo(focusParagraph).total;
    const paneRect = this.canvasPane?.getBoundingClientRect();
    const localClientX = paneRect
      ? clientX - paneRect.left + (this.canvasPane?.scrollLeft ?? 0)
      : clientX;
    const focusRunOffset = this.snapWrappedRunLocalToLineEnd(
      focusParagraph,
      this.geometryIndexToRunOffset(focusParagraph, focusGeometryOffset),
      focusRunTotal,
      localClientX
    );
    // The inline editor uses one flat run-only offset for the whole OOXML
    // paragraph, while `focusParagraph` is one of its rendered visual lines.
    // Preserve that flat coordinate here; otherwise a drag that starts on a
    // later wrapped line is clamped to the first visual line when it crosses
    // into another paragraph.
    const focusOffset = this.getVisualParagraphStartOffset(paragraphs, focusIndex) + focusRunOffset;
    const focusParagraphIndexForLog = this.getParagraphIndexFromInlineElement(focusParagraph);
    const logKey = `cross:${anchorIndex}:${focusIndex}:${focusOffset}:${focusParagraphIndexForLog}`;
    const now = performance.now();
    if (drag.lastLogKey !== logKey && now - (drag.lastLogAt ?? 0) >= 200) {
      drag.lastLogKey = logKey;
      drag.lastLogAt = now;
      debugLog('text-select', 'cross-paragraph drag selection', {
        slide: this.currentSlide,
        shapeIndex: this.activeShapeTextTarget?.shapeIndex ?? this.selectedShapeIndex,
        anchorIndex,
        focusIndex,
        anchorOffset: drag.anchorOffset,
        focusGeometryOffset,
        focusRunOffset,
        focusOffset,
        anchorParagraphIndex,
        focusParagraphIndex: focusParagraphIndexForLog,
        focusHasRuns: this.collectParagraphRuns(focusParagraph).length > 0,
        focusText: (focusParagraph.textContent ?? '').slice(0, 80),
        pointer: { x: clientX, y: clientY },
      });
    }
    this.renderCrossParagraphSelection(paragraphs, anchorIndex, drag.anchorOffset, focusIndex, focusOffset);
  }

  private getDragFocusParagraph(
    paragraphs: (SVGTextElement | SVGTSpanElement)[],
    clientY: number,
    fallback: SVGTextElement | SVGTSpanElement
  ): SVGTextElement | SVGTSpanElement {
    const paneRect = this.canvasPane?.getBoundingClientRect();
    const localY = paneRect ? clientY - paneRect.top + (this.canvasPane?.scrollTop ?? 0) : clientY;

    const candidatesWithRuns = paragraphs.filter((paragraph) => this.collectParagraphRuns(paragraph).length > 0);
    const candidates = candidatesWithRuns.length > 0 ? candidatesWithRuns : paragraphs;

    let best = fallback;
    let bestDistance = Number.POSITIVE_INFINITY;
    for (const paragraph of candidates) {
      const box = this.getElementBox(paragraph);
      if (!box) continue;
      if (localY >= box.top && localY <= box.top + box.height) {
        return paragraph;
      }
      const center = box.top + box.height / 2;
      const distance = Math.abs(localY - center);
      if (distance < bestDistance) {
        bestDistance = distance;
        best = paragraph;
      }
    }
    return best;
  }

  private getParagraphIndexFromInlineElement(element: SVGTextElement | SVGTSpanElement): number | null {
    const paragraph = element.closest('tspan[data-ooxml-para-idx]') ?? element;
    const paragraphIndex = Number((paragraph as Element).getAttribute('data-ooxml-para-idx'));
    return Number.isFinite(paragraphIndex) ? paragraphIndex : null;
  }

  private getVisualParagraphStartOffset(
    paragraphs: (SVGTextElement | SVGTSpanElement)[],
    targetIndex: number
  ): number {
    const target = paragraphs[targetIndex];
    if (!target) return 0;

    const targetParagraphIndex = this.getParagraphIndexFromInlineElement(target);
    if (targetParagraphIndex === null) return 0;

    let offset = 0;
    for (let index = 0; index < targetIndex; index++) {
      const paragraph = paragraphs[index];
      if (!paragraph || this.getParagraphIndexFromInlineElement(paragraph) !== targetParagraphIndex) continue;
      // Skip bullet/number marker containers: they carry no OOXML runs, so their
      // glyphs must not count toward the run-only paragraph offset.
      if (this.collectParagraphRuns(paragraph).length === 0) continue;
      offset += this.getRunCharInfo(paragraph).total;
    }
    return offset;
  }

  private renderCrossParagraphSelection(
    paragraphs: (SVGTextElement | SVGTSpanElement)[],
    anchorIndex: number,
    anchorOffset: number,
    focusIndex: number,
    focusOffset: number
  ): void {
    this.removeInlineSelection();
    this.activeInlineCaret?.addClass('native-powerpoint-inline-caret-hidden');

    const anchorParagraph = paragraphs[anchorIndex];
    const focusParagraph = paragraphs[focusIndex];
    if (!anchorParagraph || !focusParagraph) return;
    const anchorParagraphIndex = this.getParagraphIndexFromInlineElement(anchorParagraph);
    const focusParagraphIndex = this.getParagraphIndexFromInlineElement(focusParagraph);
    if (anchorParagraphIndex === null || focusParagraphIndex === null) return;

    const anchorComesFirst = anchorIndex < focusIndex
      || (anchorIndex === focusIndex && anchorOffset <= focusOffset);
    const startIndex = anchorComesFirst ? anchorIndex : focusIndex;
    const endIndex = anchorComesFirst ? focusIndex : anchorIndex;
    const startParagraphIndex = anchorComesFirst ? anchorParagraphIndex : focusParagraphIndex;
    const endParagraphIndex = anchorComesFirst ? focusParagraphIndex : anchorParagraphIndex;
    const startOffset = anchorComesFirst ? anchorOffset : focusOffset;
    const endOffset = anchorComesFirst ? focusOffset : anchorOffset;

    // A text box paragraph may occupy several sibling tspans after soft-wrap.
    // Keep selection ranges in paragraph-wide run offsets, then hand them to
    // the range renderer, which already maps every range back over those visual
    // lines. The old visual-line loop clamped a flat anchor (for example 187)
    // to the first line's length (for example 59), making the original
    // paragraph's highlight appear to disappear during a cross-paragraph drag.
    const paragraphIndexes: number[] = [];
    for (let index = startIndex; index <= endIndex; index++) {
      const paragraph = paragraphs[index];
      if (!paragraph || this.collectParagraphRuns(paragraph).length === 0) continue;
      const paragraphIndex = this.getParagraphIndexFromInlineElement(paragraph);
      if (paragraphIndex !== null && !paragraphIndexes.includes(paragraphIndex)) {
        paragraphIndexes.push(paragraphIndex);
      }
    }

    const textParts: string[] = [];
    const ranges: ParagraphTextRange[] = [];
    const shapeIndex = this.activeShapeTextTarget?.shapeIndex ?? this.selectedShapeIndex;
    for (const paragraphIndex of paragraphIndexes) {
      const visualLines = paragraphs.filter((paragraph) => (
        this.getParagraphIndexFromInlineElement(paragraph) === paragraphIndex
        && this.collectParagraphRuns(paragraph).length > 0
      ));
      const total = visualLines.reduce((sum, line) => sum + this.getRunCharInfo(line).total, 0);
      const rangeStart = paragraphIndex === startParagraphIndex
        ? Math.max(0, Math.min(total, startOffset))
        : 0;
      const rangeEnd = paragraphIndex === endParagraphIndex
        ? Math.max(0, Math.min(total, endOffset))
        : total;
      if (rangeEnd <= rangeStart) continue;

      ranges.push({ paragraphIndex, start: rangeStart, end: rangeEnd });

      let lineStart = 0;
      for (const line of visualLines) {
        const runText = this.collectParagraphRuns(line).map((run) => run.textContent ?? '').join('');
        const lineEnd = lineStart + runText.length;
        const localStart = Math.max(0, rangeStart - lineStart);
        const localEnd = Math.min(runText.length, rangeEnd - lineStart);
        if (localEnd > localStart) {
          textParts.push(runText.slice(localStart, localEnd));
        }
        lineStart = lineEnd;
      }
    }

    // Reuse the compound-selection copy buffer so Ctrl/Cmd+C copies the full
    // range. Rendering through this same model also keeps later caret updates
    // from repainting only the first visual line.
    this.inlineWholeShapeSelected = false;
    this.inlineWholeShapeSelection = textParts.join('\n');
    this.inlineRangeSelection = shapeIndex !== null && ranges.length > 0
      ? { shapeIndex, ranges }
      : null;
    if (this.inlineRangeSelection) {
      this.renderInlineRangeSelection(this.inlineRangeSelection);
    }
    const logKey = `${shapeIndex}:${anchorIndex}:${focusIndex}:${startIndex}:${endIndex}:${startOffset}:${endOffset}`;
    const now = performance.now();
    if (logKey !== this.lastInlineSelectionRenderedLogKey && now - this.lastInlineSelectionRenderedLogAt >= 250) {
      const skippedLogs = this.skippedInlineSelectionRenderedLogs;
      this.lastInlineSelectionRenderedLogKey = logKey;
      this.lastInlineSelectionRenderedLogAt = now;
      this.skippedInlineSelectionRenderedLogs = 0;
      debugLog('text-select', 'cross-paragraph selection rendered', {
        slide: this.currentSlide,
        shapeIndex,
        anchorIndex,
        focusIndex,
        startIndex,
        endIndex,
        startOffset,
        endOffset,
        ranges,
        renderedRectCount: this.activeInlineSelectionRects.length,
        connectedRectCount: this.activeInlineSelectionRects.filter((rect) => rect.isConnected).length,
        rectParentCount: new Set(this.activeInlineSelectionRects.map((rect) => rect.parentElement)).size,
        selectedText: this.inlineWholeShapeSelection,
        skippedLogs,
      });
    } else {
      this.skippedInlineSelectionRenderedLogs += 1;
    }
    this.updateTextToolbar();
  }

  private hasInlineSelectionDragMoved(drag: InlineSelectionDrag, clientX: number, clientY: number): boolean {
    const dx = clientX - drag.startClientX;
    const dy = clientY - drag.startClientY;
    return Math.hypot(dx, dy) >= 4;
  }

  private stopInlineSelectionDrag(): void {
    this.inlineSelectionDrag?.cleanup();
    this.inlineSelectionDrag = null;
  }

  private clearBrowserTextSelection(): void {
    activeDocument.getSelection()?.removeAllRanges();
  }

  private commitActiveTextEditing(): void {
    if (!this.activeEditor) return;
    void this.finishInlineTextEditing('commitActiveTextEditing');
  }

  private handleOutsideSlidePointerDown = (event: PointerEvent): void => {
    if (!this.activeEditor) return;

    const target = isNode(event.target) ? event.target : null;
    if (isElement(target) && target.closest(PPTX_EDITOR_FORMATTING_SURFACE_SELECTOR)) return;
    if (isElement(target) && target.closest('.native-powerpoint-thumbnail')) {
      debugLog('text-edit', 'deferring outside commit (thumbnail navigation will handle it)');
      return;
    }
    if (target && this.slideSurface?.contains(target)) return;
    if (target && this.activeEditor.contains(target)) return;

    const clearSelectionAfterCommit = this.isCanvasPaneBackgroundTarget(target);
    if (clearSelectionAfterCommit) {
      event.preventDefault();
      event.stopPropagation();
    }
    this.clearBrowserTextSelection();
    this.commitActiveEditorFromOutside(clearSelectionAfterCommit);
  };

  private handleCanvasContextMenu = (event: MouseEvent): void => {
    event.preventDefault();

    const shapeIndex = this.getTopLevelShapeIndexFromEvent(event);
    if (shapeIndex !== null) {
      this.showObjectContextMenu(event, shapeIndex);
      return;
    }

    this.showCanvasContextMenu(event);
  };

  private showCanvasContextMenu(event: MouseEvent): void {
    const menu = this.createNativeMenu();
    // Capture the slide coordinate now. The menu item's click event occurs on
    // the floating menu, not where the user opened it on the canvas.
    const textBoxOrigin = this.getTextBoxInsertOrigin(event);

    menu.addItem((item) => {
      item
        .setTitle(this.tb('paste'))
        .setIcon('clipboard-paste')
        .onClick(() => void this.pasteCopiedShape());
      if (!this.objectClipboard) item.setDisabled(true);
    });

    menu.addItem((item) => {
      const shapeCount = this.getTopLevelShapeIndices().length;
      item
        .setTitle(this.tb('selectAll'))
        .setIcon('box-select')
        .onClick(() => this.selectAllShapes());
      if (shapeCount === 0) item.setDisabled(true);
    });

    menu.addItem((item) => {
      item
        .setTitle(this.t('powerpoint:contextMenu.newTextBox'))
        .setIcon('type')
        .onClick(() => {
          if (!this.ensureEditable('add text box')) return;
          logPptxAction('insert', 'insert-text-box-context-menu', {
            slide: this.currentSlide,
            requestedOrigin: textBoxOrigin ?? null,
          });
          void this.insertController.insertTextBox(true, textBoxOrigin ?? undefined);
        });
    });

    menu.showAtMouseEvent(event);
  }

  /** Convert the original canvas right-click into a slide-space text-box origin. */
  private getTextBoxInsertOrigin(event: MouseEvent): TextBoxInsertOrigin | null {
    if (!this.engine || !this.svgEl) return null;
    const point = this.getSvgPoint(event);
    const slideScale = this.engine.getSlideScale(this.svgEl);
    if (!point || !Number.isFinite(slideScale) || slideScale <= 0) return null;
    return {
      x: Math.round(point.x * slideScale),
      y: Math.round(point.y * slideScale),
    };
  }

  /**
   * Resolve the top-level shape index for a right-click, mirroring the
   * top-level ancestor logic in {@link getTopLevelShapeIndices}: walk up out of
   * any nested group shapes to the outermost `g[data-ooxml-shape-idx]`.
   */
  private getTopLevelShapeIndexFromEvent(event: MouseEvent): number | null {
    const target = event.target;
    if (!isElement(target)) return null;

    let shape = target.closest('g[data-ooxml-shape-idx]');
    if (!shape) return null;

    let ancestor = shape.parentElement?.closest('g[data-ooxml-shape-idx]') ?? null;
    while (ancestor) {
      shape = ancestor;
      ancestor = shape.parentElement?.closest('g[data-ooxml-shape-idx]') ?? null;
    }
    const shapeIndex = getShapeIndex(shape);
    return isSelectableShapeIndex(shapeIndex) ? shapeIndex : null;
  }

  private getObjectKind(shapeIndex: number, shapeEl: SVGGElement): 'image' | 'text' | 'generic' {
    const type = shapeEl.getAttribute('data-ooxml-shape-type');
    if (type === 'table' || type === 'chart' || type === 'group') {
      return 'generic';
    }
    if (this.engine?.isImageShape(this.currentSlide, shapeIndex) || shapeEl.querySelector('image')) {
      return 'image';
    }
    if (shapeEl.querySelector('text')) {
      return 'text';
    }
    return 'generic';
  }

  private canFlipShape(shapeEl: SVGGElement): boolean {
    const type = shapeEl.getAttribute('data-ooxml-shape-type');
    return type !== 'chart' && type !== 'table';
  }

  /**
   * Build the right-click menu for a slide object. The right-clicked shape is
   * selected first (unless it is already part of the current multi-selection)
   * so every action operates on it.
   */
  private showObjectContextMenu(event: MouseEvent, shapeIndex: number): void {
    if (!this.selectedShapeIndices.has(shapeIndex)) {
      this.selectShape(shapeIndex);
    }

    const shapeEl = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    const kind = isSVGGElement(shapeEl) ? this.getObjectKind(shapeIndex, shapeEl) : 'generic';
    const canEdit = this.canEdit();
    const menu = this.createNativeMenu();

    menu.addItem((item) => {
      item.setTitle(this.tb('cut')).setIcon('scissors').onClick(() => void this.cutSelectedShape());
      if (!canEdit) item.setDisabled(true);
    });
    menu.addItem((item) =>
      item.setTitle(this.tb('copy')).setIcon('copy').onClick(() => void this.copySelectedShape())
    );
    menu.addItem((item) => {
      item.setTitle(this.tb('paste')).setIcon('clipboard-paste').onClick(() => void this.pasteCopiedShape());
      if (!canEdit || !this.objectClipboard) item.setDisabled(true);
    });
    menu.addItem((item) => {
      item
        .setTitle(this.tb('pasteWithoutFormatting'))
        .setIcon('clipboard-type')
        .onClick(() => void this.pasteWithoutFormatting());
      if (!canEdit || (!this.activeEditor && !this.objectClipboard)) item.setDisabled(true);
    });
    menu.addItem((item) => {
      item.setTitle(this.tb('delete')).setIcon('trash-2').onClick(() => void this.deleteSelectedShape());
      if (!canEdit) item.setDisabled(true);
    });

    menu.addSeparator();
    this.addOrderSubsection(menu, canEdit);
    this.addRotateSubsection(
      menu,
      canEdit,
      isSVGGElement(shapeEl) ? this.canFlipShape(shapeEl) : false
    );
    this.addCenterSubsection(menu, canEdit);

    if (isSVGGElement(shapeEl)) {
      this.addShapeFillColorMenuItem(menu, shapeIndex, canEdit);
    }

    if (kind === 'image') {
      menu.addSeparator();
      menu.addItem((item) => {
        item.setTitle(this.t('powerpoint:contextMenu.cropImage')).setIcon('crop').onClick(() => this.openImageCropDialog(shapeIndex));
        if (!canEdit) item.setDisabled(true);
      });
      this.addReplaceImageSubsection(menu, canEdit, shapeIndex);
      menu.addItem((item) => {
        item
          .setTitle(this.t('powerpoint:contextMenu.resetImage'))
          .setIcon('rotate-ccw')
          .onClick(() => void this.resetSelectedImage(shapeIndex));
        if (!canEdit) item.setDisabled(true);
      });
      menu.addItem((item) => {
        item
          .setTitle(this.t('powerpoint:contextMenu.setAsBackground'))
          .setIcon('image')
          .onClick(() => void this.setSelectedImageAsBackground(shapeIndex));
        if (!canEdit) item.setDisabled(true);
      });
    }

    menu.showAtMouseEvent(event);
  }

  private addShapeFillColorMenuItem(menu: Menu, shapeIndex: number, canEdit: boolean): void {
    if (!this.engine?.canSetShapeFillColor(this.currentSlide, shapeIndex)) return;

    menu.addSeparator();
    menu.addItem((item) => {
      item
        .setTitle(this.t('powerpoint:contextMenu.shapeFillColor'))
        .setIcon('palette')
        .onClick(() => this.openShapeFillColorPicker(shapeIndex));
      if (!canEdit) item.setDisabled(true);
    });
  }

  private openShapeFillColorPicker(shapeIndex: number): void {
    if (!this.engine?.canSetShapeFillColor(this.currentSlide, shapeIndex)) return;
    const anchor = this.selectionOverlay ?? this.canvasPane;
    if (!anchor) return;

    const currentColor = this.engine.getShapeVisualStyle(this.currentSlide, shapeIndex)?.fill ?? 'FFFFFF';
    logPptxAction('inspector', 'open-shape-color-picker', {
      slide: this.currentSlide,
      shapeIndexes: [shapeIndex],
      color: currentColor,
    });
    this.openColorPopover(anchor, currentColor, false, (color) => {
      if (color) void this.applyShapeFillColor(shapeIndex, color);
    });
  }

  private async applyShapeFillColor(shapeIndex: number, hex: string): Promise<void> {
    if (!this.engine || !this.ensureEditable('change shape fill color')) return;

    logPptxAction('inspector', 'set-shape-fill-color', {
      slide: this.currentSlide,
      shapeIndexes: [shapeIndex],
      color: hex,
    });
    try {
      const history = this.captureSlideXmlHistoryEntry(this.currentSlide, 'Shape fill color');
      await this.session.applyCommand({
        type: 'set-shape-fill-color',
        slideIndex: this.currentSlide,
        shapeIndex,
        hex,
      });
      this.recordHistoryEntry(this.completeSlideXmlHistoryEntry(history));
      const rendered = await this.renderEditedShape(shapeIndex);
      if (rendered) {
        this.scheduleThumbnailRefresh(this.currentSlide);
        this.updateSelectionOverlay();
      }
      debugLog('inspector', 'Changed PowerPoint shape fill color', {
        op: 'set-shape-fill-color-complete',
        slide: this.currentSlide,
        shapeIndexes: [shapeIndex],
        color: hex,
      });
    } catch (error) {
      errorLog('inspector', 'PowerPoint shape fill color change failed', {
        slide: this.currentSlide,
        shapeIndexes: [shapeIndex],
        color: hex,
        error,
      });
      pptNotice('powerpoint:notice.couldNotChangeShapeFillColor', { message: cleanError(error) });
    }
  }

  /**
   * Add a labelled group of actions to a menu. Uses a real Obsidian side
   * submenu when {@link MenuItem.setSubmenu} is available (Obsidian 1.4.5+) and
   * falls back to flat, prefixed items with a section label otherwise.
   */
  private addMenuSubsection(
    menu: Menu,
    title: string,
    icon: string,
    populate: (add: (label: string, icon: string, onClick: () => void, disabled?: boolean) => void) => void
  ): void {
    const holder: { submenu: Menu | null } = { submenu: null };
    menu.addItem((item) => {
      item.setTitle(title).setIcon(icon);
      const withSubmenu = item as unknown as { setSubmenu?: () => Menu | undefined };
      const created = typeof withSubmenu.setSubmenu === 'function' ? withSubmenu.setSubmenu() : undefined;
      if (created) {
        holder.submenu = created;
        (created as unknown as { dom?: HTMLElement }).dom?.addClass('native-powerpoint-light-surface');
      } else {
        item.setIsLabel(true);
      }
    });

    const target = holder.submenu;
    if (target) {
      populate((label, itemIcon, onClick, disabled) =>
        target.addItem((item) => {
          item.setTitle(label).setIcon(itemIcon).onClick(onClick);
          if (disabled) item.setDisabled(true);
        })
      );
    } else {
      populate((label, itemIcon, onClick, disabled) =>
        menu.addItem((item) => {
          item.setTitle(this.t('powerpoint:contextMenu.submenuItem', { section: title, label })).setIcon(itemIcon).onClick(onClick);
          if (disabled) item.setDisabled(true);
        })
      );
    }
  }

  private addOrderSubsection(menu: Menu, canEdit: boolean): void {
    this.addMenuSubsection(menu, this.t('powerpoint:contextMenu.order'), 'layers', (add) => {
      add(this.tb('bringToFront'), 'bring-to-front', () => void this.arrangeController.reorderSelection('front'), !canEdit);
      add(this.tb('bringForward'), 'arrow-up', () => void this.arrangeController.reorderSelection('forward'), !canEdit);
      add(this.tb('sendBackward'), 'arrow-down', () => void this.arrangeController.reorderSelection('backward'), !canEdit);
      add(this.tb('sendToBack'), 'send-to-back', () => void this.arrangeController.reorderSelection('back'), !canEdit);
    });
  }

  private addRotateSubsection(menu: Menu, canEdit: boolean, allowFlip: boolean): void {
    this.addMenuSubsection(menu, this.t('powerpoint:contextMenu.rotate'), 'rotate-cw', (add) => {
      add(this.t('powerpoint:contextMenu.rotateRight90'), 'rotate-cw', () => void this.rotateSelectedShape(90), !canEdit);
      add(this.t('powerpoint:contextMenu.rotateLeft90'), 'rotate-ccw', () => void this.rotateSelectedShape(-90), !canEdit);
      if (allowFlip) {
        add(this.t('powerpoint:contextMenu.flipHorizontal'), 'flip-horizontal', () => void this.flipSelectedShape('horizontal'), !canEdit);
        add(this.t('powerpoint:contextMenu.flipVertical'), 'flip-vertical', () => void this.flipSelectedShape('vertical'), !canEdit);
      }
    });
  }

  private addCenterSubsection(menu: Menu, canEdit: boolean): void {
    this.addMenuSubsection(menu, this.t('powerpoint:contextMenu.centerOnPage'), 'align-center-horizontal', (add) => {
      add(this.t('powerpoint:contextMenu.centerHorizontally'), 'align-center-vertical', () => void this.centerSelectedOnPage('horizontal'), !canEdit);
      add(this.t('powerpoint:contextMenu.centerVertically'), 'align-center-horizontal', () => void this.centerSelectedOnPage('vertical'), !canEdit);
    });
  }

  private addReplaceImageSubsection(menu: Menu, canEdit: boolean, shapeIndex: number): void {
    this.addMenuSubsection(menu, this.t('powerpoint:contextMenu.replaceImage'), 'image-plus', (add) => {
      add(this.t('powerpoint:contextMenu.fromVaultEllipsis'), 'folder', () => this.openReplaceImageVaultPicker(shapeIndex), !canEdit);
      add(this.t('powerpoint:contextMenu.uploadFileEllipsis'), 'upload', () => {
        if (!this.ensureEditable('replace image')) return;
        this.pendingReplaceShapeIndex = shapeIndex;
        this.replaceImageFileInput?.click();
      }, !canEdit);
    });
  }

  private handleCanvasPanePointerDown = (event: PointerEvent): void => {
    this.lastInteractionRegion = 'canvas';
    this.slideFilmstripController.clearSlideSelection();
    if (event.button !== 0) return;
    this.suppressNextClick = false;

    const target = isNode(event.target) ? event.target : null;
    if (!this.isCanvasPaneBackgroundTarget(target)) return;

    event.preventDefault();
    event.stopPropagation();
    this.clearBrowserTextSelection();
    if (this.activeEditor) {
      this.commitActiveEditorFromOutside(true);
      return;
    }

    this.removeInlineSelection();
    this.lastInlineCaretPlacement = null;
    this.beginMarquee(event, event.shiftKey || event.ctrlKey || event.metaKey);
  };

  private isCanvasPaneBackgroundTarget(target: Node | null): boolean {
    if (!target || !this.canvasPane?.contains(target)) return false;
    if (this.slideSurface?.contains(target)) return false;
    if (this.activeEditor?.contains(target)) return false;

    const element = isElement(target) ? target : target.parentElement;
    if (element?.closest('.native-powerpoint-selection-box')) return false;
    // The contextual text toolbar lives inside the canvas pane, so clicks on it
    // must not be treated as background clicks (which would commit/clear the
    // editor and hide the toolbar mid-interaction).
    if (element?.closest(PPTX_EDITOR_FORMATTING_SURFACE_SELECTOR)) return false;

    return true;
  }

  private commitActiveEditorFromOutside(clearSelectionAfterCommit: boolean): void {
    this.clearBrowserTextSelection();
    void this.finishInlineTextEditing('outside-pointerdown').finally(() => {
      this.clearBrowserTextSelection();
      if (clearSelectionAfterCommit) {
        this.clearSelection({ skipTextCommit: true });
      }
    });
  }

  private startTextEditor(target: TextEditTarget | null = null, clientX?: number, clientY?: number): void {
    if (!this.canvasPane || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('edit text')) return;

    const selected = this.getSelectedShapeElement();
    if (!selected) return;
    if (!target) {
      target = this.getTextEditTargetFromSelectedShape();
    }
    if (!target && selected.closest(GENERATED_GRID_SELECTOR)) {
      this.showGeneratedTextNotice();
      return;
    }
    if (!target) return;

    if (this.activeEditor && this.activeEditorTarget === target.element) {
      const currentBox = this.getElementBox(target.element);
      if (currentBox) {
        this.placeInlineCaret(this.activeEditor, target.element, clientX, clientY, currentBox);
        this.focusEditorWithoutCanvasScroll(this.activeEditor);
        this.resetInlineEditorScroll(this.activeEditor);
      }
      return;
    }

    if (this.activeEditor) {
      this.commitActiveTextEditing();
    }
    if (this.activeEditor) {
      this.removeActiveEditor();
    }
    const box = this.getElementBox(target.element);
    if (!box) return;

    const editor = this.canvasPane.createEl('textarea', {
      cls: 'native-powerpoint-inline-editor is-text-run',
      attr: { 'aria-label': this.t('powerpoint:accessibility.editSelectedText') }
    });
    const initialText = target.text;
    const initialRunTexts = target.kind === 'shape-paragraph'
      ? target.runElements.map((run) => run.textContent || '')
      : [];
    editor.value = initialText;
    this.activeEditorTextDirty = false;
		this.activeEditorInitialWordCount = countDocumentWords(initialText);
    let previousInlineEditorValue = initialText;

    const styleElement = target.kind === 'shape-paragraph' && target.runElements[0]
      ? target.runElements[0]
      : target.element;
    const style = window.getComputedStyle(styleElement);
    target.element.classList.add('native-powerpoint-text-editing');
    this.activeEditorTarget = target.element;
    this.activeShapeTextTarget = target.kind === 'shape-paragraph' ? target : null;
    this.activeTextStyleTarget = target.kind === 'shape-paragraph'
      ? this.getPrimaryStyleRunTarget(target)
      : null;
    this.activeInlinePreviewLineCount = target.kind === 'shape-paragraph'
      ? this.getRunLineContainers(target.shapeIndex, target.paragraphIndex).length
      : null;
    this.slideSurface?.addClass('is-inline-text-editing');
    editor.setCssProps({
      color: style.fill,
      fontFamily: style.fontFamily,
      fontSize: `${this.getScreenFontSize(styleElement)}px`,
      fontStyle: style.fontStyle,
      fontWeight: style.fontWeight,
      lineHeight: '1.1',
      textAlign: this.getInlineTextAlignment(style.textAnchor)
    });
    this.positionTextRunEditor(editor, box);
    this.activeEditor = editor;
    this.activeEditorCanonicalText = initialText;
    this.inlineUndoStack = [];
    this.inlineRedoStack = [];
    this.lastInlineHistoryRestoreFailed = false;
    this.historyController.updateAvailability();
    let pendingInlineEditScroll: CanvasScrollPosition | null = null;
    let pendingInlineInputType: string | null = null;
    let pendingInlineBeforeInputSeen = false;
    let pendingInlineDeleteKeySeen = false;
    // Native input normally repaints a changed SVG run, but Chromium can retain
    // stale glyphs when a transparent textarea clears all visible text. Keep a
    // one-event flag so the input handler can replace the owning <text> frame.
    let pendingInlineTextFrameRefresh = false;

    const captureInlineEditScroll = (inputType: string | null): void => {
      // Native textarea edits can make Chromium scroll the nearest overflow
      // ancestor to reveal its invisible 1px caret. The on-slide SVG caret is
      // the visible caret, so preserve the canvas rather than accepting that
      // browser-driven pan.
      pendingInlineEditScroll ??= this.captureCanvasScroll();
      pendingInlineInputType = inputType ?? pendingInlineInputType;
    };

    // Capture the pre-edit state (value + selection) before each native edit so
    // an in-place undo can restore the text *and* re-select whatever was just
    // deleted. Programmatic edits (e.g. handleInlineDeleteKey) snapshot
    // themselves since they don't fire beforeinput.
    editor.addEventListener('beforeinput', (event) => {
      if (this.activeEditor === editor) {
        const inputType = event.inputType || null;
        pendingInlineBeforeInputSeen = true;
        captureInlineEditScroll(inputType);
        this.recordInlineEditSnapshot(editor);
        const selectionStart = Math.min(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
        const selectionEnd = Math.max(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
        pendingInlineTextFrameRefresh = this.shouldRefreshInlineTextFrameForNativeInput(
          inputType,
          editor.value,
          selectionStart,
          selectionEnd,
        );
        if (inputType?.startsWith('delete')) {
          debugLog('text-edit', 'Inline editor deletion beforeinput', {
            slide: this.currentSlide,
            shapeIndex: target.shapeIndex,
            paragraphIndex: target.kind === 'shape-paragraph' ? target.paragraphIndex : null,
            inputType,
            previousTextLength: editor.value.length,
            selectionStart,
            selectionEnd,
            requestedTextFrameRefresh: pendingInlineTextFrameRefresh,
          });
        }
      }
    });

    this.updateSelectionOverlay();
    this.activeInlineCaret = this.createInlineCaret();
    debugLog('text-edit', 'Inline text editor opened', {
      slide: this.currentSlide,
      shapeIndex: target.shapeIndex,
      paragraphIndex: target.kind === 'shape-paragraph' ? target.paragraphIndex : null,
      hasCaret: this.activeInlineCaret !== null
    });
    const updateCaret = () => {
      this.rememberCollapsedInlineCaretPlacement(editor, target.element);
      this.updateInlineCaret(editor, target.element);
    };
    const queueCaretUpdate = () => {
      window.requestAnimationFrame(() => {
        if (this.activeEditor === editor) {
          updateCaret();
        }
      });
    };
    editor.addEventListener('pointerdown', (event) => {
      event.preventDefault();
      this.clearWholeShapeInlineSelection();
      const nextBox = this.getElementBox(target.element) ?? box;
      this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(target.element, event.clientY, nextBox);
      const offset = this.getInlineTextOffsetAtClientPoint(target.element, editor, event.clientX, event.clientY, nextBox);
      this.focusEditorWithoutCanvasScroll(editor);
      editor.setSelectionRange(offset, offset);
      this.rememberInlineCaretPlacement(editor, target.element, offset);
      this.resetInlineEditorScroll(editor);
      updateCaret();
    });
    editor.addEventListener('input', () => {
      if (this.activeEditor === editor) {
        const inputType = pendingInlineInputType;
        const previousText = previousInlineEditorValue;
        // Electron can dispatch input without the matching beforeinput metadata.
        // The value transition is the reliable source of truth: every shrink
        // needs the owning SVG <text> frame to repaint immediately.
        const nativeTextWasDeleted = editor.value.length < previousText.length;
        const nativeInputTextFrameRefresh = this.shouldRefreshInlineTextFrameForNativeInput(
          inputType,
          previousText,
          0,
          0,
          editor.value,
        );
        const fullSelectionTextFrameRefresh = pendingInlineTextFrameRefresh || this.inlineWholeShapeSelected;
        const replacesWholeShape = target.kind === 'shape-paragraph'
          && this.inlineWholeShapeSelected
          && this.beginInlineWholeShapeReplacement(target);
        const replaceTextFrame = fullSelectionTextFrameRefresh || nativeInputTextFrameRefresh;
        const destructiveInput = nativeTextWasDeleted
          || inputType?.startsWith('delete') === true
          || pendingInlineDeleteKeySeen;
        if (destructiveInput) {
          // A 1px textarea can receive browser/native editing that is invisible
          // to the user until a later commit. Keep a narrow breadcrumb for the
          // destructive boundary so a future editor/SVG split is diagnosable.
          debugLog('text-edit', 'Inline editor destructive input', {
            slide: this.currentSlide,
            shapeIndex: target.shapeIndex,
            paragraphIndex: target.kind === 'shape-paragraph' ? target.paragraphIndex : null,
            inputType,
            previousTextLength: previousText.length,
            nextTextLength: editor.value.length,
            selectionStart: editor.selectionStart ?? 0,
            selectionEnd: editor.selectionEnd ?? 0,
            beforeInputSeen: pendingInlineBeforeInputSeen,
            deleteKeySeen: pendingInlineDeleteKeySeen,
            requestedTextFrameRefresh: replaceTextFrame,
            replacesWholeShape,
          });
        }
        this.activeEditorTextDirty = true;
        this.clearWholeShapeInlineSelection();
        const textFrameReplaced = this.syncInlineTextPreviewSafely(target, editor.value, {
          replaceTextFrame,
        });
        if (fullSelectionTextFrameRefresh && target.kind === 'shape-paragraph') {
          const previewLines = this.getRunLineContainers(target.shapeIndex, target.paragraphIndex);
          debugLog('text-edit', 'Inline full-selection preview refreshed', {
            slide: this.currentSlide,
            shapeIndex: target.shapeIndex,
            paragraphIndex: target.paragraphIndex,
            inputType,
            previousTextLength: previousText.length,
            nextTextLength: editor.value.length,
            previewTextLength: this.getParagraphPlainText(previewLines).length,
            previewLineCount: previewLines.length,
            textFrameReplaced,
            targetConnected: target.element.isConnected,
          });
        }
        if (destructiveInput) {
          const previewLines = target.kind === 'shape-paragraph'
            ? this.getRunLineContainers(target.shapeIndex, target.paragraphIndex)
            : [];
          const previewText = target.kind === 'shape-paragraph'
            ? this.getParagraphPlainText(previewLines)
            : target.element.textContent || '';
          debugLog('text-edit', 'Synced native inline deletion preview', {
            slide: this.currentSlide,
            shapeIndex: target.shapeIndex,
            paragraphIndex: target.kind === 'shape-paragraph' ? target.paragraphIndex : null,
            inputType,
            fullSelectionTextFrameRefresh,
            nativeTextWasDeleted,
            beforeInputSeen: pendingInlineBeforeInputSeen,
            deleteKeySeen: pendingInlineDeleteKeySeen,
            previousTextLength: previousText.length,
            nextTextLength: editor.value.length,
            previewTextLength: previewText.length,
            previewLineCount: previewLines.length,
            textFrameReplaced,
            previewTargetConnected: target.element.isConnected,
          });
          if (previewText !== editor.value) {
            warnLog('text-edit', 'Inline deletion preview did not match editor text after sync', {
              slide: this.currentSlide,
              shapeIndex: target.shapeIndex,
              paragraphIndex: target.kind === 'shape-paragraph' ? target.paragraphIndex : null,
              inputType,
              previousTextLength: previousText.length,
              nextTextLength: editor.value.length,
              previewTextLength: previewText.length,
              previewLineCount: previewLines.length,
              previewRunCount: target.kind === 'shape-paragraph' ? target.runElements.length : null,
              previewTargetConnected: target.element.isConnected,
            });
          }
        }
        previousInlineEditorValue = editor.value;
        this.activeEditorCanonicalText = editor.value;
        const nextBox = this.getElementBox(target.element);
        if (nextBox) {
          this.positionTextRunEditor(editor, nextBox);
        }
        updateCaret();
        this.preserveCanvasScrollAfterInlineTextEdit(pendingInlineEditScroll, pendingInlineInputType);
        pendingInlineEditScroll = null;
        pendingInlineInputType = null;
        pendingInlineBeforeInputSeen = false;
        pendingInlineDeleteKeySeen = false;
        pendingInlineTextFrameRefresh = false;
        this.publishPresentationWordCount();
      }
    });
    editor.addEventListener('copy', (event) => {
      if (this.inlineWholeShapeSelection !== null) {
        event.preventDefault();
        event.clipboardData?.setData('text/plain', this.inlineWholeShapeSelection);
      }
    });
    editor.addEventListener('click', updateCaret);
    editor.addEventListener('keyup', updateCaret);
    editor.addEventListener('mouseup', updateCaret);
    editor.addEventListener('select', updateCaret);
    editor.addEventListener('keydown', (event) => {
      if (event.key === 'Delete' || event.key === 'Backspace') {
        pendingInlineDeleteKeySeen = true;
        captureInlineEditScroll(event.key === 'Backspace' ? 'deleteContentBackward' : 'deleteContentForward');
        debugLog('text-edit', 'Inline editor deletion keydown', {
          slide: this.currentSlide,
          shapeIndex: target.shapeIndex,
          paragraphIndex: target.kind === 'shape-paragraph' ? target.paragraphIndex : null,
          key: event.key,
          selectionStart: editor.selectionStart ?? 0,
          selectionEnd: editor.selectionEnd ?? 0,
          textLength: editor.value.length,
        });
      }
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'a') {
        event.preventDefault();
        this.selectAllInlineText(editor, target.element);
        return;
      }
      if (
        target.kind === 'shape-paragraph'
        && event.key === 'Enter'
        && !event.isComposing
        && !event.shiftKey
        && !event.metaKey
        && !event.ctrlKey
        && !event.altKey
      ) {
        // Native Enter means a new PowerPoint paragraph. Leaving this to the
        // textarea would write a raw newline into the SVG preview and later
        // serialize it as a soft break in the same `<a:p>`.
        event.preventDefault();
        this.clearWholeShapeInlineSelection();
        this.startInlineParagraphSplit(editor, target);
        return;
      }
      if (
        event.shiftKey
        && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)
      ) {
        this.clearWholeShapeInlineSelection();
        this.lastInlineCaretPlacement = null;
      } else if (
        this.inlineWholeShapeSelection !== null
        && !this.inlineWholeShapeSelected
        && !event.metaKey
        && !event.ctrlKey
        && !event.altKey
        && (event.key === 'Enter' || event.key.length === 1 || ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key))
      ) {
        this.clearWholeShapeInlineSelection();
      }
      if (this.handleInlineDeleteKey(event, editor, target.element)) {
        previousInlineEditorValue = editor.value;
        pendingInlineEditScroll = null;
        pendingInlineInputType = null;
        pendingInlineBeforeInputSeen = false;
        pendingInlineDeleteKeySeen = false;
        pendingInlineTextFrameRefresh = false;
        return;
      }
      queueCaretUpdate();
    });

    const commitSlideIndex = this.currentSlide;
    const commitTarget = target;
    const commit = async () => {
      if (this.paragraphSplitPromise) {
        await this.paragraphSplitPromise;
      }
      if (this.paragraphRemovalPromise) {
        await this.paragraphRemovalPromise;
      }
      if (this.rangeDeletionPromise) {
        await this.rangeDeletionPromise;
      }
      if (this.activeEditor !== editor) return;
      const text = this.resolveInlineTextForCommit(editor);
      const textWasEdited = this.activeEditorTextDirty;
      const replaceWholeShape = this.inlineWholeShapeReplacement?.shapeIndex === commitTarget.shapeIndex;
      const wholeShapeBaseline = replaceWholeShape
        ? this.inlineWholeShapeReplacement?.baselineText ?? null
        : null;
      const normalizedText = !replaceWholeShape && commitTarget.kind === 'shape-paragraph'
        ? this.paragraphEditorTextFromDom(commitTarget.shapeIndex, commitTarget.paragraphIndex, text)
        : text;
      this.removeActiveEditor(editor);
      if (!textWasEdited) {
        debugLog('text-edit', 'Skipped inline text commit (unchanged)', {
          slideIndex: commitSlideIndex,
          shapeIndex: commitTarget.shapeIndex
        });
        return;
      }
      if (replaceWholeShape) {
        debugLog('text-edit', 'Committing whole-shape inline text replacement', {
          slideIndex: commitSlideIndex,
          shapeIndex: commitTarget.shapeIndex,
          characterCount: normalizedText.length,
          baselineLength: wholeShapeBaseline?.length ?? null,
        });
      }
      await this.applyTextValue(
        normalizedText,
        replaceWholeShape ? null : commitTarget,
        commitSlideIndex,
        replaceWholeShape ? { authoritativePreviousText: wholeShapeBaseline } : {},
      );
    };
    this.activeEditorCommit = commit;

    editor.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
        event.preventDefault();
        void commit();
      } else if (event.key === 'Escape') {
        if (this.restoreInlineWholeShapeReplacementPreview(target.shapeIndex)) {
          // The all-text preview was detached and replaced while editing. Its
          // captured SVG is the only complete cancellation snapshot.
        } else if (target.kind === 'shape-paragraph') {
          target.runElements.forEach((run, index) => {
            run.textContent = initialRunTexts[index] ?? '';
          });
        } else {
          target.element.textContent = initialText;
        }
        this.setInlineEditorValue(editor, initialText);
        this.removeActiveEditor(editor);
      }
    });
    editor.addEventListener('blur', (event) => {
      if (this.isTearingDownEditor || this.isNavigatingSlide) {
        debugLog('text-edit', 'ignored blur during slide navigation/teardown');
        return;
      }

      // A bare modifier press (e.g. ⌘/Ctrl arming the app menu) or the window
      // losing key focus blurs the textarea with no related target. Treat that
      // as transient: keep the editor open and restore focus so an in-progress
      // text selection isn't discarded. Genuine clicks outside the editor are
      // already committed by the document-level pointerdown handler before this
      // fires, so by then activeEditor no longer matches and we fall through to
      // commit (which early-returns). Focus moving to another real element
      // (e.g. Tab) still commits.
      const next = event.relatedTarget;
      const stayEditing = next === null
        || (isElement(next)
          && next.closest(PPTX_EDITOR_FORMATTING_SURFACE_SELECTOR) !== null);
      if (stayEditing && this.activeEditor === editor && editor.isConnected) {
        this.focusEditorWithoutCanvasScroll(editor);
        updateCaret();
        return;
      }
      void commit();
    });
    this.focusEditorWithoutCanvasScroll(editor);
    this.placeInlineCaret(editor, target.element, clientX, clientY, box);
    this.updateTextToolbar();
  }

  private createInlineCaret(): SVGLineElement | null {
    const svg = this.svgEl;
    if (!svg) return null;

    const caret = this.getSvgFactory(svg).createSvg('line');
    caret.classList.add('native-powerpoint-svg-caret');
    caret.setAttribute('aria-hidden', 'true');
    svg.appendChild(caret);
    return caret;
  }

  private getSvgFactory(owner: SVGElement): SvgFactoryWindow {
    const ownerDocument = owner.ownerDocument as Document & { win: SvgFactoryWindow };
    return ownerDocument.win;
  }

  private renderFontFidelity(container: HTMLElement): void {
    if (!this.engine) return;

    const section = container.createDiv({ cls: 'native-powerpoint-font-fidelity' });
    section.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: this.t('powerpoint:inspector.fonts') });

    if (this.fontSubstitutions.length === 0) {
      section.createDiv({
        cls: 'native-powerpoint-inspector-hint',
        text: this.t('powerpoint:inspector.fontsAvailable')
      });
      return;
    }

    section.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: this.t('powerpoint:inspector.fontsSubstituted', { count: this.fontSubstitutions.length })
    });
    const list = section.createDiv({ cls: 'native-powerpoint-font-substitution-list' });
    for (const substitution of this.fontSubstitutions) {
      const item = list.createDiv({ cls: 'native-powerpoint-font-substitution' });
      item.createSpan({ cls: 'native-powerpoint-font-substitution-source', text: substitution.requested });
      item.createSpan({ cls: 'native-powerpoint-font-substitution-arrow', text: this.t('powerpoint:inspector.substitutionArrow') });
      item.createSpan({ cls: 'native-powerpoint-font-substitution-target', text: substitution.substitute });
    }
  }

  private removeActiveEditor(editor = this.activeEditor): void {
    if (editor && this.activeEditor && editor !== this.activeEditor) return;

    debugLog('text-edit', 'removeActiveEditor', {
      slide: this.currentSlide,
      navigating: this.isNavigatingSlide
    });
    this.stopInlineSelectionDrag();
    this.clearWholeShapeInlineSelection();
    this.inlineWholeShapeReplacement = null;
    this.activeEditor?.remove();
    this.activeInlineCaret?.remove();
    this.removeInlineSelection();
    this.activeEditor = null;
		this.activeEditorCanonicalText = null;
		this.activeEditorInitialWordCount = 0;
    this.activeEditorTextDirty = false;
    this.activeEditorCommit = null;
    this.inlineUndoStack = [];
    this.inlineRedoStack = [];
    this.lastInlineHistoryRestoreFailed = false;
    this.activeInlineCaret = null;
    this.lastInlineCaretPlacement = null;
    this.activeInlineCaretRow = null;
    this.activeInlinePreviewLineCount = null;
    this.activeEditorTarget?.classList.remove('native-powerpoint-text-editing');
    this.activeEditorTarget = null;
    this.activeShapeTextTarget = null;
    this.activeTextStyleTarget = null;
    this.slideSurface?.removeClass('is-inline-text-editing');
    if (this.svgEl?.isConnected) {
      this.updateSelectionOverlay();
    } else {
      this.updateTextToolbar();
    }
    this.historyController.updateAvailability();
  }

  // --- Contextual text formatting toolbar (Google Slides–style) ------------

  private shapeHasEditableText(shape: Element): boolean {
    return shape.querySelector('tspan[data-ooxml-run-idx]') !== null
      && shape.closest(GENERATED_GRID_SELECTOR) === null;
  }

  private getTextStyleContext(): TextStyleContext | null {
    if (!this.engine || !this.canEdit() || !this.svgEl) return null;

    if (this.activeTextStyleTarget && this.activeEditor) {
      const target = this.activeTextStyleTarget;
      const anchor = this.getElementBox(this.activeEditorTarget ?? target.element);
      if (!anchor) return null;
      return {
        shapeIndex: target.shapeIndex,
        run: { paragraphIndex: target.paragraphIndex, runIndex: target.runIndex },
        anchor
      };
    }

    const formattingSnapshot = this.textToolbarController.getFormattingSnapshot() ?? this.toolbarFormattingSnapshot;
    if (formattingSnapshot) {
      const anchor = this.getSelectedBox() ?? formattingSnapshot.anchor;
      return {
        shapeIndex: formattingSnapshot.shapeIndex,
        run: formattingSnapshot.run,
        anchor
      };
    }

    // Keep the formatting toolbar visible while a single text shape is selected,
    // even when no inline editor is active (e.g. after the editor is flushed
    // because the user clicked into the toolbar's font-size box).
    if (this.selectedShapeIndex !== null && this.selectedShapeIndices.size <= 1) {
      const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${this.selectedShapeIndex}"]`);
      if (shape && this.shapeHasEditableText(shape)) {
        const anchor = this.getSelectedBox();
        if (anchor) {
          return { shapeIndex: this.selectedShapeIndex, run: null, anchor };
        }
      }
    }

    return null;
  }

  private getFirstRunTarget(shapeIndex: number): RunTarget | null {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    const run = shape?.querySelector('tspan[data-ooxml-run-idx]') ?? null;
    if (!run) return null;

    const paragraph = run.closest('tspan[data-ooxml-para-idx]');
    const paragraphIndex = Number(paragraph?.getAttribute('data-ooxml-para-idx'));
    const runIndex = Number(run.getAttribute('data-ooxml-run-idx'));
    if (!Number.isFinite(paragraphIndex) || !Number.isFinite(runIndex)) return null;
    return { paragraphIndex, runIndex };
  }

  private buildParagraphEditTarget(shapeIndex: number, paragraphIndex: number): ShapeTextEditTarget | null {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!shape) return null;

    // Seed from a run tspan inside the target paragraph so getTextEditTarget
    // resolves the correct paragraph (passing the paragraph tspan alone makes it
    // fall back to the first paragraph in the shape).
    const seed = shape.querySelector(`tspan[data-ooxml-para-idx="${paragraphIndex}"] tspan[data-ooxml-run-idx]`)
      ?? shape.querySelector(`tspan[data-ooxml-para-idx="${paragraphIndex}"]`)
      ?? shape.querySelector('text');
    const target = this.getTextEditTarget(seed);
    return target?.kind === 'shape-paragraph' ? target : null;
  }

  private buildRunTarget(shapeIndex: number, paragraphIndex: number, runIndex: number): ShapeTextEditTarget | null {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!shape) return null;

    const run = Array.from(shape.querySelectorAll('tspan[data-ooxml-run-idx]')).find((candidate) => {
      const paragraph = candidate.closest('tspan[data-ooxml-para-idx]');
      return Number(candidate.getAttribute('data-ooxml-run-idx')) === runIndex
        && Number(paragraph?.getAttribute('data-ooxml-para-idx')) === paragraphIndex;
    });
    if (!isSVGTSpanElement(run)) return null;

    const paragraph = run.closest('tspan[data-ooxml-para-idx]');
    const geometryElement = paragraph && isSVGTSpanElement(paragraph) ? paragraph : run;
    const runElements = paragraph
      ? this.collectParagraphRuns(paragraph)
      : [run];
    const text = runElements.map((candidate) => candidate.textContent || '').join('');

    return {
      kind: 'shape-paragraph',
      shapeIndex,
      paragraphIndex,
      runIndex,
      text,
      element: geometryElement,
      runElements
    };
  }

  private getStoredInlineSelectionRanges(shapeIndex: number): ParagraphTextRange[] | null {
    if (this.inlineRangeSelection?.shapeIndex === shapeIndex && this.inlineRangeSelection.ranges.length > 0) {
      return this.inlineRangeSelection.ranges;
    }

    if (this.inlineWholeShapeSelected) {
      const ranges = this.getShapeTextRanges(shapeIndex);
      return ranges.length > 0 ? ranges : null;
    }

    return null;
  }

  private getActiveInlineSelectionRanges(shapeIndex: number): ParagraphTextRange[] | null {
    return this.getStoredInlineSelectionRanges(shapeIndex)
      ?? (this.textToolbarController.getFormattingSnapshot()?.shapeIndex === shapeIndex
        ? this.textToolbarController.getFormattingSnapshot()?.ranges ?? null
        : this.toolbarFormattingSnapshot?.shapeIndex === shapeIndex
          ? this.toolbarFormattingSnapshot.ranges
          : null);
  }

  private getSelectedRangeFontSizePt(
    shapeIndex: number,
    ranges: ParagraphTextRange[]
  ): number | null {
    const engine = this.engine;
    if (!engine) return null;
    return engine.getRangesFontSizePt(
      this.currentSlide,
      shapeIndex,
      this.mapRangesToOoxmlOffsets(shapeIndex, ranges)
    );
  }

  private captureToolbarFormattingSnapshot(): ToolbarFormattingSnapshot | null {
    const context = this.getTextStyleContext();
    if (!context) return null;

    let ranges = this.getStoredInlineSelectionRanges(context.shapeIndex);
    if (!ranges && this.toolbarFormattingSnapshot?.shapeIndex === context.shapeIndex) {
      ranges = this.toolbarFormattingSnapshot.ranges;
    }
    const editor = this.activeEditor;
    const target = this.activeTextStyleTarget;
    if (!ranges && editor && target && target.shapeIndex === context.shapeIndex) {
      const start = Math.min(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      const end = Math.max(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      ranges = [{ paragraphIndex: target.paragraphIndex, start, end }];
    }

    return { ...context, ranges };
  }

  private flushActiveEditorForToolbarInput(): void {
    const snapshot = this.captureToolbarFormattingSnapshot();
    this.flushActiveEditor();
    this.toolbarFormattingSnapshot = snapshot;
  }

  private updateTextToolbar(): void {
    this.textToolbarController.updateTextToolbar();
		this.publishPresentationWordCount();
  }

  private hideTextToolbar(): void {
    this.textToolbarEl?.removeClass('is-visible');
    this.textToolbarShapeIndex = null;
    this.closeToolbarPopover();
    this.currentRunStyle = null;
    this.setTopFontControl(this.tb('font'), false);
  }

  // The font-family picker lives in the always-visible top toolbar but only acts
  // on the active text selection, so it is disabled (and reset to a neutral
  // label) whenever there is no editable text context.
  private setTopFontControl(label: string, enabled: boolean): void {
    this.topFontLabel?.setText(label);
    const button = this.topFontButton;
    if (!button) return;
    button.disabled = !enabled;
    button.toggleClass('is-disabled', !enabled);
    button.setAttribute('aria-disabled', String(!enabled));
  }

  private positionTextToolbar(anchor: { left: number; top: number; width: number; height: number }): void {
    const toolbar = this.textToolbarEl;
    if (!toolbar || !this.canvasPane) return;

    const toolbarHeight = toolbar.offsetHeight || 40;
    const gap = 8;
    let top = anchor.top - toolbarHeight - gap;
    if (top < this.canvasPane.scrollTop + 4) {
      top = anchor.top + anchor.height + gap;
    }

    const maxLeft = Math.max(0, this.canvasPane.scrollWidth - (toolbar.offsetWidth || 0) - 4);
    const left = Math.min(Math.max(anchor.left, 4), maxLeft);
    toolbar.setCssProps({ left: `${left}px`, top: `${Math.max(0, top)}px` });
  }

  private reflectTextToolbarState(context: TextStyleContext): void {
    const controls = this.textToolbarControls;
    if (!controls || !this.engine) return;

    const runTarget = context.run ?? this.getFirstRunTarget(context.shapeIndex);
    const style = runTarget
      ? this.engine.getRunStyle(this.currentSlide, context.shapeIndex, runTarget.paragraphIndex, runTarget.runIndex)
      : null;
    const selectedRanges = this.getActiveInlineSelectionRanges(context.shapeIndex);
    const reflectedStyle = selectedRanges?.length && style
      ? {
          ...style,
          bold: this.engine.areRangesStyled(this.currentSlide, context.shapeIndex, selectedRanges, 'bold'),
          italic: this.engine.areRangesStyled(this.currentSlide, context.shapeIndex, selectedRanges, 'italic'),
          underline: this.engine.areRangesStyled(this.currentSlide, context.shapeIndex, selectedRanges, 'underline')
        }
      : style;
    this.currentRunStyle = reflectedStyle;

    controls.bold.toggleClass('is-active', Boolean(reflectedStyle?.bold));
    controls.italic.toggleClass('is-active', Boolean(reflectedStyle?.italic));
    controls.underline.toggleClass('is-active', Boolean(reflectedStyle?.underline));
    this.setTopFontControl(reflectedStyle?.fontFamily ?? this.getEffectiveFontFamily(context) ?? this.tb('font'), true);

    if (activeDocument.activeElement !== controls.fontSizeInput) {
      const sizePt = reflectedStyle?.fontSizePt ?? this.getEffectiveFontSizePt(context);
      controls.fontSizeInput.value = sizePt ? String(Math.round(sizePt)) : '';
    }

    if (reflectedStyle?.color) {
      this.textColorValue = reflectedStyle.color;
    }
    if (reflectedStyle?.highlight) {
      this.textHighlightValue = reflectedStyle.highlight;
    }
    controls.textColorBar.style.setProperty('--np-swatch-color', `#${reflectedStyle?.color ?? this.textColorValue}`);
    controls.highlightBar.style.setProperty('--np-swatch-color', reflectedStyle?.highlight ? `#${reflectedStyle.highlight}` : 'transparent');

    const alignment = reflectedStyle?.alignment ?? 'l';
    for (const align of ['l', 'ctr', 'r', 'just'] as ParagraphAlignment[]) {
      controls.alignButtons[align].toggleClass('is-active', alignment === align);
    }
  }

  // EMU per SVG user unit, used to convert rendered font sizes back to points.
  private getSvgEmuPerUnit(): number | null {
    const scale = Number(this.svgEl?.getAttribute('data-ooxml-scale'));
    if (Number.isFinite(scale) && scale > 0) return scale;

    const cx = Number(this.svgEl?.getAttribute('data-ooxml-slide-cx'));
    const width = this.svgEl ? Number.parseFloat(this.svgEl.getAttribute('width') || '') : Number.NaN;
    if (Number.isFinite(cx) && Number.isFinite(width) && width > 0) return cx / width;

    return null;
  }

  /**
   * Detect the effective font size (in points) actually rendered for the
   * relevant runs, so the toolbar can show a value even when the size is
   * inherited from the theme/placeholder rather than authored on each run.
   * Returns null when sizes are mixed or cannot be determined.
   */
  private getRelevantTextRuns(context: TextStyleContext): SVGTSpanElement[] {
    if (!this.svgEl) return [];
    const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${context.shapeIndex}"]`);
    if (!shape) return [];

    const allRuns = Array.from(shape.querySelectorAll('tspan[data-ooxml-run-idx]')).filter(isSVGTSpanElement);
    const targetRun = context.run;
    if (!targetRun) return allRuns;

    return allRuns.filter((run) => {
      const para = run.closest('tspan[data-ooxml-para-idx]');
      return Number(run.getAttribute('data-ooxml-run-idx')) === targetRun.runIndex
        && Number(para?.getAttribute('data-ooxml-para-idx')) === targetRun.paragraphIndex;
    });
  }

  private getEffectiveFontSizePt(context: TextStyleContext): number | null {
    const emuPerUnit = this.getSvgEmuPerUnit();
    if (!emuPerUnit) return null;

    const runs = this.getRelevantTextRuns(context);
    if (runs.length === 0) return null;

    const EMU_PER_POINT = 12700;
    let detected: number | null = null;
    for (const run of runs) {
      if ((run.textContent || '').length === 0) continue;
      const userUnits = Number.parseFloat(window.getComputedStyle(run).fontSize);
      if (!Number.isFinite(userUnits) || userUnits <= 0) continue;
      const rounded = Math.round((userUnits * emuPerUnit) / EMU_PER_POINT);
      if (detected === null) {
        detected = rounded;
      } else if (detected !== rounded) {
        return null;
      }
    }
    return detected;
  }

  /**
   * Detect the effective font family actually rendered for the relevant runs,
   * so the toolbar can show a value when the face is inherited from the theme
   * or placeholder rather than authored on each run. Returns null when families
   * are mixed or cannot be determined.
   */
  private getEffectiveFontFamily(context: TextStyleContext): string | null {
    const runs = this.getRelevantTextRuns(context);
    if (runs.length === 0) return null;

    let detected: string | null = null;
    for (const run of runs) {
      if ((run.textContent || '').length === 0) continue;
      const family = parsePrimaryFontFamily(window.getComputedStyle(run).fontFamily);
      if (!family) continue;
      if (detected === null) {
        detected = family;
      } else if (detected !== family) {
        return null;
      }
    }
    return detected;
  }

  private ensureTextToolbar(): TextToolbarControls | null {
    if (this.textToolbarControls && this.textToolbarEl?.isConnected) {
      return this.textToolbarControls;
    }
    if (!this.canvasPane) return null;

    this.textToolbarEl?.remove();
    const toolbar = this.canvasPane.createDiv({ cls: PPTX_EDITOR_CHROME_CONTEXT_TOOLBAR_CLASS });
    toolbar.setAttribute('role', 'toolbar');
    toolbar.setAttribute('aria-label', pptT('powerpoint:accessibility.textFormatting'));
    toolbar.addEventListener('pointerdown', (event) => event.stopPropagation());

    const styleGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const bold = this.createTextToolbarButton(styleGroup, 'bold', this.tb('bold'), () => this.toggleRunFlag('bold'));
    const italic = this.createTextToolbarButton(styleGroup, 'italic', this.tb('italic'), () => this.toggleRunFlag('italic'));
    const underline = this.createTextToolbarButton(styleGroup, 'underline', this.tb('underline'), () => this.toggleRunFlag('underline'));

    const sizeGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    this.createTextToolbarButton(sizeGroup, 'minus', this.tb('decreaseFontSize'), () => this.stepFontSize(-1));
    const fontSizeInput = sizeGroup.createEl('input', {
      cls: 'native-powerpoint-text-toolbar-size',
      type: 'number',
      attr: {
        'aria-label': this.t('powerpoint:accessibility.fontSize'),
        min: String(TEXT_TOOLBAR_MIN_FONT_SIZE),
        max: String(TEXT_TOOLBAR_MAX_FONT_SIZE)
      }
    });
    fontSizeInput.addEventListener('pointerdown', () => this.flushActiveEditorForToolbarInput(), true);
    fontSizeInput.addEventListener('focus', () => this.flushActiveEditorForToolbarInput());
    fontSizeInput.addEventListener('change', () => this.commitFontSizeInput());
    fontSizeInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        this.commitFontSizeInput();
      }
    });
    this.createTextToolbarButton(sizeGroup, 'plus', this.tb('increaseFontSize'), () => this.stepFontSize(1));

    const colorGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const textColorButton = this.createTextToolbarSwatchButton(colorGroup, 'baseline', this.tb('textColor'));
    const textColorBar = textColorButton.createDiv({ cls: 'native-powerpoint-text-toolbar-swatch-bar' });
    this.bindToolbarButton(textColorButton, () =>
      this.openColorPopover(textColorButton, this.textColorValue, false, (color) => {
        debugLog('text-format', 'setTextColor', { color });
        this.applyRunStyle({ color });
      }));

    const highlightButton = this.createTextToolbarSwatchButton(colorGroup, 'highlighter', this.tb('highlightColor'));
    const highlightBar = highlightButton.createDiv({ cls: 'native-powerpoint-text-toolbar-swatch-bar' });
    this.bindToolbarButton(highlightButton, () =>
      this.openColorPopover(highlightButton, this.textHighlightValue, true, (color) => {
        debugLog('text-format', 'setHighlight', { color });
        this.applyRunStyle({ highlight: color });
      }));

    const alignGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const alignButtons: Record<ParagraphAlignment, HTMLButtonElement> = {
      l: this.createTextToolbarButton(alignGroup, 'align-left', this.tb('alignLeft'), () => this.applyAlignment('l')),
      ctr: this.createTextToolbarButton(alignGroup, 'align-center', this.tb('alignCenter'), () => this.applyAlignment('ctr')),
      r: this.createTextToolbarButton(alignGroup, 'align-right', this.tb('alignRight'), () => this.applyAlignment('r')),
      just: this.createTextToolbarButton(alignGroup, 'align-justify', this.tb('justify'), () => this.applyAlignment('just'))
    };

    this.textToolbarEl = toolbar;
    this.textToolbarControls = {
      bold,
      italic,
      underline,
      fontSizeInput,
      textColorBar,
      highlightBar,
      alignButtons
    };
    return this.textToolbarControls;
  }

  private createTextToolbarButton(
    container: HTMLElement,
    icon: string,
    label: string,
    action: () => void
  ): HTMLButtonElement {
    const button = createToolbarIconButton(container, {
      className: ['native-powerpoint-toolbar-btn', 'native-powerpoint-text-toolbar-btn'],
      icon,
      label
    });
    this.bindToolbarButton(button, action);
    return button;
  }

  private createTextToolbarSwatchButton(container: HTMLElement, icon: string, label: string): HTMLButtonElement {
    return createToolbarIconButton(container, {
      className: ['native-powerpoint-toolbar-btn', 'native-powerpoint-text-toolbar-btn', 'native-powerpoint-text-toolbar-swatch'],
      icon,
      label,
      iconClassName: 'native-powerpoint-text-toolbar-swatch-icon'
    });
  }

  private bindToolbarButton(button: HTMLElement, action: () => void): void {
    button.addEventListener('pointerdown', (event) => {
      event.preventDefault();
      event.stopPropagation();
      action();
    });
  }

  private toggleRunFlag(flag: 'bold' | 'italic' | 'underline'): void {
    const context = this.getTextStyleContext();
    if (context && this.engine) {
      const selectedRanges = this.getActiveInlineSelectionRanges(context.shapeIndex);
      if (selectedRanges?.length) {
        const next = !this.engine.areRangesStyled(this.currentSlide, context.shapeIndex, selectedRanges, flag);
        debugLog('text-format', 'toggleRunFlag', { flag, path: 'inline-ranges', next, shapeIndex: context.shapeIndex });
        this.applyRunStyle({ [flag]: next });
        return;
      }
    }

    const editor = this.activeEditor;
    const target = this.activeTextStyleTarget;
    if (editor && target && this.engine) {
      const start = Math.min(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      const end = Math.max(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      if (start < end) {
        const next = !this.engine.isRangeStyled(
          this.currentSlide,
          target.shapeIndex,
          target.paragraphIndex,
          start,
          end,
          flag
        );
        debugLog('text-format', 'toggleRunFlag', {
          flag,
          path: 'editor-range',
          next,
          shapeIndex: target.shapeIndex,
          paragraphIndex: target.paragraphIndex,
          start,
          end
        });
        this.applyRunStyle({ [flag]: next });
        return;
      }
    }

    const current = this.currentRunStyle?.[flag] ?? false;
    debugLog('text-format', 'toggleRunFlag', { flag, path: 'caret-or-shape', next: !current });
    this.applyRunStyle({ [flag]: !current });
  }

  private stepFontSize(delta: number): void {
    const inputValue = Number(this.textToolbarControls?.fontSizeInput?.value);
    const current = this.currentRunStyle?.fontSizePt
      ?? (Number.isFinite(inputValue) && inputValue > 0 ? inputValue : 18);
    const next = Math.min(TEXT_TOOLBAR_MAX_FONT_SIZE, Math.max(TEXT_TOOLBAR_MIN_FONT_SIZE, Math.round(current) + delta));
    debugLog('text-format', 'stepFontSize', { delta, current, next });
    this.applyRunStyle({ fontSizePt: next });
  }

  private commitFontSizeInput(): void {
    const input = this.textToolbarControls?.fontSizeInput;
    if (!input) return;

    const value = Number(input.value);
    if (!Number.isFinite(value)) return;

    const clamped = Math.min(
      TEXT_TOOLBAR_MAX_FONT_SIZE,
      Math.max(TEXT_TOOLBAR_MIN_FONT_SIZE, Math.round(value))
    );
    debugLog('text-format', 'commitFontSizeInput', { value, clamped });
    this.applyRunStyle({ fontSizePt: clamped });
  }

  private flushActiveEditor(): void {
    const editor = this.activeEditor;
    const target = this.activeTextStyleTarget;
    if (!editor) return;

    const textWasEdited = this.activeEditorTextDirty;
    this.removeActiveEditor(editor);
    if (target && textWasEdited && editor.value !== target.text) {
      void this.applyTextValue(editor.value, target);
    }
  }

  private applyRunStyle(change: RunStyleChange): void {
    const engine = this.engine;
    if (!engine) return;
    const isHighlightClear = change.highlight === null;
    void this.runTextFormatting('Format text', (shapeIndex, run, selection) => {
      if (selection?.length) {
        debugLog('text-select', 'applyRunStyle with ranges', {
          shapeIndex,
          change,
          ranges: selection.map((range) => ({
            paragraphIndex: range.paragraphIndex,
            start: range.start,
            end: range.end
          }))
        });
        if (isHighlightClear) {
          this.pendingHighlightClear = {
            slide: this.currentSlide,
            shapeIndex,
            paragraphs: new Set(selection.map((range) => range.paragraphIndex)),
            ranges: selection.map((range) => ({ ...range }))
          };
        }
        return this.session.applyCommand({
          type: 'set-run-style-ranges',
          slideIndex: this.currentSlide,
          shapeIndex,
          ranges: selection,
          change
        });
      }
      debugLog('text-format', 'applyRunStyle whole run/shape', {
        shapeIndex,
        change,
        run: run ? { paragraphIndex: run.paragraphIndex, runIndex: run.runIndex } : null
      });
      if (isHighlightClear) {
        this.pendingHighlightClear = {
          slide: this.currentSlide,
          shapeIndex,
          paragraphs: new Set(run ? [run.paragraphIndex] : []),
          ranges: null
        };
      }
      return this.session.applyCommand({
        type: 'set-run-style',
        slideIndex: this.currentSlide,
        shapeIndex,
        target: run,
        change
      });
    });
  }

  private applyAlignment(align: ParagraphAlignment): void {
    const engine = this.engine;
    if (!engine) return;
    void this.runTextFormatting('Align text', (shapeIndex, run, selection) => {
      if (selection?.length) {
        debugLog('text-format', 'applyAlignment with ranges', {
          align,
          shapeIndex,
          ranges: selection.map((range) => ({
            paragraphIndex: range.paragraphIndex,
            start: range.start,
            end: range.end
          }))
        });
        return this.session.applyCommand({
          type: 'set-paragraph-alignment-ranges',
          slideIndex: this.currentSlide,
          shapeIndex,
          ranges: selection,
          align
        });
      }
      debugLog('text-format', 'applyAlignment paragraph', {
        align,
        shapeIndex,
        paragraphIndex: run ? run.paragraphIndex : null
      });
      return this.session.applyCommand({
        type: 'set-paragraph-alignment',
        slideIndex: this.currentSlide,
        shapeIndex,
        paragraphIndex: run ? run.paragraphIndex : null,
        align
      });
    });
  }

  private async runTextFormatting(
    label: string,
    apply: (
      shapeIndex: number,
      run: RunTarget | null,
      selection: ParagraphTextRange[] | null
    ) => Promise<unknown>
  ): Promise<void> {
    const engine = this.engine;
    if (!engine || !this.ensureEditable('format text')) {
      debugLog('text-format', 'runTextFormatting aborted', { label, reason: 'not-editable' });
      return;
    }

    const context = this.getTextStyleContext();
    if (!context) {
      debugLog('text-format', 'runTextFormatting aborted', { label, reason: 'no-context' });
      pptNotice('powerpoint:notice.selectTextBoxFirst');
      return;
    }
    const usedToolbarFormattingSnapshot =
      this.toolbarFormattingSnapshot !== null || this.textToolbarController.hasFormattingSnapshot();
    debugLog('text-format', 'runTextFormatting start', {
      label,
      slide: this.currentSlide,
      shapeIndex: context.shapeIndex,
      run: context.run ? { paragraphIndex: context.run.paragraphIndex, runIndex: context.run.runIndex } : null,
      usedToolbarFormattingSnapshot
    });

    // Capture the live inline-editor selection up front. The textarea itself
    // lives in the canvas pane (not the slide SVG), so it survives a slide
    // re-render — we keep it open and refresh its element references in place
    // afterwards instead of tearing it down and reopening, which is what used to
    // drop the selection and make a follow-up Bold apply to the whole shape.
    const editor = this.activeEditor;
    const styleTarget = this.activeTextStyleTarget;
    let pendingText: string | null = null;
    const selectedRanges = this.getActiveInlineSelectionRanges(context.shapeIndex);
    let selectionRanges: ParagraphTextRange[] | null = selectedRanges;
    let savedStart = 0;
    let savedEnd = 0;
    if (editor && styleTarget) {
      savedStart = Math.min(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      savedEnd = Math.max(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      if (!selectionRanges) {
        selectionRanges = [{ paragraphIndex: styleTarget.paragraphIndex, start: savedStart, end: savedEnd }];
      }
      if (this.activeEditorTextDirty) {
        const normalizedEditorText = this.paragraphEditorTextFromDom(
          context.shapeIndex,
          styleTarget.paragraphIndex,
          editor.value
        );
        pendingText = normalizedEditorText !== styleTarget.text ? normalizedEditorText : null;
      } else if (editor.value !== styleTarget.text) {
        debugLog('text-format', 'Skipped inline text sync before formatting', {
          shapeIndex: context.shapeIndex,
          paragraphIndex: styleTarget.paragraphIndex,
          editorTextLength: editor.value.length,
          renderedTextLength: styleTarget.text.length,
          reason: 'editor-unmodified'
        });
      }
    }

    const restoreInlineRangeSelection = selectedRanges !== null;
    const scrollPosition = this.captureCanvasScroll();
    try {
      const history = this.captureSlideXmlHistoryEntry(this.currentSlide, label);
      if (pendingText !== null && styleTarget) {
        await this.session.applyCommand({
          type: 'update-paragraph-text',
          slideIndex: this.currentSlide,
          shapeIndex: context.shapeIndex,
          paragraphIndex: styleTarget.paragraphIndex,
          text: pendingText
        });
        if (this.activeEditor === editor) {
          this.activeEditorTextDirty = false;
        }
      }
      // The engine styles in OOXML run-offset space; the editor/SVG ranges drop
      // wrap-boundary whitespace, so map them across before applying. The UI
      // keeps the editor-space `selectionRanges` for caret/selection restore.
      const engineRanges = selectionRanges
        ? this.mapRangesToOoxmlOffsets(context.shapeIndex, selectionRanges)
        : null;
      await apply(context.shapeIndex, context.run, engineRanges);
      this.recordHistoryEntry(this.completeSlideXmlHistoryEntry(history));

      const rendered = await this.renderEditedShape(context.shapeIndex);
      if (rendered) {
        this.restoreCanvasScrollSoon(scrollPosition);
        this.scheduleThumbnailRefresh(this.currentSlide);
        if (editor && this.activeEditor === editor && selectionRanges) {
          if (this.refreshActiveShapeEditorAfterRender()) {
            const length = editor.value.length;
            if (!restoreInlineRangeSelection) {
              this.clearWholeShapeInlineSelection();
              editor.setSelectionRange(Math.min(savedStart, length), Math.min(savedEnd, length));
            }
            this.refreshInlineEditorGeometry();
          } else {
            this.removeActiveEditor(editor);
          }
        }
        this.updateTextToolbar();
      }
      debugLog('text-format', 'runTextFormatting complete', {
        label,
        slide: this.currentSlide,
        rendered,
        pendingTextCommitted: pendingText !== null,
        selectionRanges: selectionRanges?.map((range) => ({
          paragraphIndex: range.paragraphIndex,
          start: range.start,
          end: range.end
        })) ?? null
      });
    } catch (error) {
      errorLog('text-format', 'runTextFormatting failed', { label, error: cleanError(error) });
      pptNotice('powerpoint:notice.couldNotFormatText', { message: cleanError(error) });
    } finally {
      if (usedToolbarFormattingSnapshot) {
        this.toolbarFormattingSnapshot = null;
        this.textToolbarController.clearFormattingSnapshot();
      }
    }
  }

  /**
   * After a slide re-render, re-point the still-open inline editor at the freshly
   * rendered paragraph/run nodes (the old ones are detached) and rebuild the
   * SVG-side caret. Returns false when the paragraph can no longer be found.
   */
  private refreshActiveShapeEditorAfterRender(): boolean {
    const target = this.activeShapeTextTarget;
    if (!target || !this.activeEditor) return false;

    const fresh = this.buildParagraphEditTarget(target.shapeIndex, target.paragraphIndex);
    if (!fresh) return false;

    this.activeEditorTarget?.classList.remove('native-powerpoint-text-editing');
    target.element = fresh.element;
    target.runElements = fresh.runElements;
    target.runIndex = fresh.runIndex;
    target.text = fresh.text;

    this.activeEditorTarget = fresh.element;
    this.activeEditorTarget.classList.add('native-powerpoint-text-editing');
    this.activeTextStyleTarget = this.getPrimaryStyleRunTarget(target);
    this.slideSurface?.addClass('is-inline-text-editing');

    // Drop the stale caret/selection and re-create the caret line. With a full
    // re-render the old nodes died with the replaced SVG, but the incremental
    // single-shape swap keeps the same root `<svg>`, so the previous caret line
    // must be removed explicitly to avoid leaving an orphaned stray caret.
    this.activeInlineCaret?.remove();
    this.removeInlineSelection();
    this.activeInlineCaret = this.createInlineCaret();
    this.removeSelectionOverlay();
    return true;
  }

  private refreshInlineEditorGeometry(): void {
    const editor = this.activeEditor;
    const target = this.activeShapeTextTarget;
    if (!editor || !target) return;

    const box = this.getElementBox(target.element);
    if (box) {
      this.positionTextRunEditor(editor, box);
    }
    this.focusEditorWithoutCanvasScroll(editor);
    this.updateInlineCaret(editor, target.element);
  }

  private openFontMenu(anchor: HTMLElement): void {
    const fonts = [...TEXT_TOOLBAR_FONTS];
    const context = this.getTextStyleContext();
    const current = this.currentRunStyle?.fontFamily
      ?? (context ? this.getEffectiveFontFamily(context) : null);
    if (current && !fonts.includes(current)) {
      fonts.unshift(current);
    }

    this.openToolbarPopover(anchor, (popover) => {
      popover.addClass('native-powerpoint-font-menu');
      for (const font of fonts) {
        const item = createMenuItem(popover, {
          className: 'native-powerpoint-color-popover-item native-powerpoint-font-menu-item',
          text: font
        });
        item.style.setProperty('--np-font-family', font);
        if (current === font) {
          item.addClass('is-active');
        }
        this.bindToolbarButton(item, () => {
          this.closeToolbarPopover();
          debugLog('text-format', 'setFontFamily', { font });
          this.applyRunStyle({ fontFamily: font });
        });
      }
    });
  }

  private openColorPopover(
    anchor: HTMLElement,
    currentColor: string,
    allowNone: boolean,
    onPick: (color: string | null) => void
  ): void {
    this.openToolbarPopover(anchor, (popover) => {
      popover.addClass('native-powerpoint-color-popover');

      if (allowNone) {
        const noneButton = createMenuItem(popover, {
          className: 'native-powerpoint-color-popover-none',
          text: this.t('powerpoint:color.noColor')
        });
        this.bindToolbarButton(noneButton, () => {
          this.closeToolbarPopover();
          onPick(null);
        });
      }

      const grid = popover.createDiv({ cls: 'native-powerpoint-color-popover-grid' });
      for (const swatch of TEXT_TOOLBAR_SWATCHES) {
        const cell = grid.createEl('button', {
          cls: 'native-powerpoint-color-popover-swatch',
          attr: { 'aria-label': this.t('powerpoint:accessibility.swatchColor', { color: swatch }) }
        });
        cell.style.setProperty('--np-swatch-color', `#${swatch}`);
        const fill = cell.createSpan({ cls: 'native-powerpoint-color-popover-swatch-fill' });
        fill.style.setProperty('--np-swatch-color', `#${swatch}`);
        if (swatch.toUpperCase() === currentColor.toUpperCase()) {
          cell.addClass('is-active');
        }
        this.bindToolbarButton(cell, () => {
          this.closeToolbarPopover();
          onPick(swatch);
        });
      }

      const customRow = popover.createDiv({ cls: 'native-powerpoint-color-popover-custom' });
      customRow.createSpan({ text: this.t('powerpoint:color.custom') });
      const customInput = customRow.createEl('input', {
        type: 'color',
        attr: { 'aria-label': this.t('powerpoint:color.custom'), value: `#${currentColor}` }
      });
      customInput.value = `#${currentColor}`;
      customInput.addEventListener('pointerdown', () => this.flushActiveEditorForToolbarInput(), true);
      customInput.addEventListener('focus', () => this.flushActiveEditorForToolbarInput());
      customInput.addEventListener('change', () => {
        const picked = customInput.value.replace(/^#/, '').toUpperCase();
        this.closeToolbarPopover();
        onPick(picked);
      });
    });
  }

  private openToolbarPopover(anchor: HTMLElement, build: (popover: HTMLElement) => void): void {
    this.closeToolbarPopover();

    const popover = createPopoverShell(activeDocument.body, {
      className: 'native-powerpoint-toolbar-popover native-powerpoint-light-surface',
      stopPointerDown: true
    });
    build(popover);

    positionPopoverBelow(popover, anchor);

    this.activeToolbarPopover = popover;
    this.toolbarPopoverCleanup = bindPopoverDismissHandlers({
      popover,
      anchor,
      onDismiss: () => this.closeToolbarPopover(),
      closeOnEscape: false
    });
  }

  private closeToolbarPopover(): void {
    this.toolbarPopoverCleanup?.();
    this.toolbarPopoverCleanup = null;
    this.activeToolbarPopover?.remove();
    this.activeToolbarPopover = null;
  }

  private captureCanvasScroll(): CanvasScrollPosition | null {
    if (!this.canvasPane) return null;
    return {
      left: this.canvasPane.scrollLeft,
      top: this.canvasPane.scrollTop
    };
  }

  private restoreCanvasScroll(position: CanvasScrollPosition | null): void {
    if (!position || !this.canvasPane) return;
    this.canvasPane.scrollLeft = position.left;
    this.canvasPane.scrollTop = position.top;
  }

  private restoreCanvasScrollSoon(position: CanvasScrollPosition | null): void {
    this.restoreCanvasScroll(position);
    if (!position) return;

    window.requestAnimationFrame(() => this.restoreCanvasScroll(position));
    window.setTimeout(() => this.restoreCanvasScroll(position), 0);
  }

  private focusEditorWithoutCanvasScroll(editor: HTMLTextAreaElement): void {
    const scrollPosition = this.captureCanvasScroll();
    editor.focus({ preventScroll: true });
    this.restoreCanvasScrollSoon(scrollPosition);
  }

  /**
   * The canvas is the user's viewport. Native edits occur in an invisible,
   * 1px textarea layered over the slide, so browser caret reveal must not pan
   * that viewport away from the text the user is editing.
   */
  private preserveCanvasScrollAfterInlineTextEdit(
    position: CanvasScrollPosition | null,
    inputType: string | null
  ): void {
    if (!position || !this.canvasPane) return;

    const observed = this.captureCanvasScroll();
    const moved = observed !== null
      && (observed.left !== position.left || observed.top !== position.top);
    this.restoreCanvasScrollSoon(position);

    if (observed && moved) {
      debugLog('text-edit', 'Restored canvas position after inline text edit', {
        slide: this.currentSlide,
        inputType,
        expectedScrollLeft: position.left,
        expectedScrollTop: position.top,
        observedScrollLeft: observed.left,
        observedScrollTop: observed.top
      });
    }
  }

  private selectEditorWithoutCanvasScroll(editor: HTMLTextAreaElement): void {
    const scrollPosition = this.captureCanvasScroll();
    editor.select();
    this.restoreCanvasScrollSoon(scrollPosition);
  }

  private getScreenFontSize(element: SVGTextElement | SVGTSpanElement): number {
    return this.inlineGeometry.getScreenFontSize(element);
  }

  private getInlineTextAlignment(textAnchor: string): string {
    if (textAnchor === 'middle') return 'center';
    if (textAnchor === 'end') return 'right';
    return 'left';
  }

  private positionTextRunEditor(
    editor: HTMLTextAreaElement,
    box: { left: number; top: number; width: number; height: number }
  ): void {
    editor.setCssProps({
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: '1px',
      height: '1px'
    });
  }

  private placeInlineCaret(
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement,
    clientX: number | undefined,
    clientY: number | undefined,
    box: { left: number; top: number; width: number; height: number }
  ): void {
    this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(element, clientY, box);
    const text = editor.value;
    if (clientX === undefined || box.width <= 0 || text.length === 0) {
      editor.setSelectionRange(text.length, text.length);
      this.rememberInlineCaretPlacement(editor, element, text.length);
      this.updateInlineCaret(editor, element);
      return;
    }

    const offset = this.getInlineTextOffsetAtClientPoint(element, editor, clientX, clientY, box);
    editor.setSelectionRange(offset, offset);
    this.rememberInlineCaretPlacement(editor, element, offset);
    this.resetInlineEditorScroll(editor);
    this.updateInlineCaret(editor, element);
  }

  private rememberInlineCaretPlacement(
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement,
    offset: number
  ): void {
    this.lastInlineCaretPlacement = {
      editor,
      element,
      offset: Math.max(0, Math.min(offset, editor.value.length)),
      timestamp: Date.now()
    };
  }

  /**
   * Decide whether a native textarea mutation needs the larger SVG repaint
   * boundary. Chromium's input metadata is not reliable across all Electron
   * routes, so a shrink in the textarea value is the authoritative deletion
   * signal and always gets the larger repaint boundary.
   */
  private shouldRefreshInlineTextFrameForNativeInput(
    inputType: string | null,
    previousText: string,
    selectionStart: number,
    selectionEnd: number,
    nextText: string | null = null,
  ): boolean {
    const normalizedInputType = inputType || '';
    const isDeletion = normalizedInputType.startsWith('delete');
    const mutatesText = isDeletion || normalizedInputType.startsWith('insert');
    const textWasDeleted = nextText !== null && nextText.length < previousText.length;
    if (textWasDeleted) return true;
    const selectionCoversPreviousText = previousText.length > 0
      && Math.min(selectionStart, selectionEnd) === 0
      && Math.max(selectionStart, selectionEnd) === previousText.length;
    if (mutatesText && selectionCoversPreviousText) return true;

    // A collapsed Backspace/Delete over the last character does not have a
    // full pre-input selection, but it leaves the same empty-glyph condition.
    return isDeletion && previousText.length > 0 && nextText === '';
  }

  /**
   * Keep deletion's textarea, SVG preview, and later OOXML commit in one
   * transaction. A preview exception must not leave the invisible textarea
   * ahead of the visible SVG until blur.
   */
  private syncInlineTextPreviewSafely(
    target: TextEditTarget,
    text: string,
    options: { replaceTextFrame?: boolean } = {},
  ): boolean {
    try {
      return this.syncShapeParagraphPreview(target, text, options);
    } catch (error) {
      errorLog('text-edit', 'Inline text preview sync failed', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.kind === 'shape-paragraph' ? target.paragraphIndex : null,
        textLength: text.length,
        replaceTextFrame: options.replaceTextFrame === true,
        error: cleanError(error),
      });
      return this.recoverInlineTextPreviewAfterSyncFailure(target, text);
    }
  }

  /**
   * Last-resort immediate preview for a failed incremental sync. It deliberately
   * reduces the current paragraph to one styled visual line; the authoritative
   * OOXML commit re-renders the proper wrapping and run formatting afterward.
   */
  private recoverInlineTextPreviewAfterSyncFailure(target: TextEditTarget, text: string): boolean {
    if (target.kind !== 'shape-paragraph') {
      target.element.textContent = text;
      return false;
    }

    try {
      const lines = this.getRunLineContainers(target.shapeIndex, target.paragraphIndex);
      const firstLine = lines[0];
      const firstRun = firstLine ? this.collectParagraphRuns(firstLine)[0] : target.runElements[0];
      if (!firstLine || !firstRun) return false;

      const replacementLine = firstLine.cloneNode(false) as SVGTSpanElement;
      const replacementRun = firstRun.cloneNode(true) as SVGTSpanElement;
      replacementRun.textContent = text;
      replacementLine.appendChild(replacementRun);
      firstLine.replaceWith(replacementLine);
      for (const line of lines.slice(1)) line.remove();

      const previousElement = target.element;
      target.element = replacementLine;
      target.runElements = [replacementRun];
      if (this.activeShapeTextTarget === target) {
        if (this.activeEditorTarget === previousElement) {
          previousElement.classList.remove('native-powerpoint-text-editing');
        }
        this.activeEditorTarget = replacementLine;
        replacementLine.classList.add('native-powerpoint-text-editing');
        this.activeTextStyleTarget = this.getPrimaryStyleRunTarget(target);
      }

      const textFrameReplaced = this.replaceLiveShapeTextFrame(target);
      debugLog('text-edit', 'Recovered inline text preview after sync failure', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
        textLength: text.length,
        textFrameReplaced,
      });
      return textFrameReplaced;
    } catch (recoveryError) {
      errorLog('text-edit', 'Inline text preview recovery failed', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
        textLength: text.length,
        error: cleanError(recoveryError),
      });
      return false;
    }
  }

  private handleInlineDeleteKey(
    event: KeyboardEvent,
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement
  ): boolean {
    if (event.key !== 'Delete' && event.key !== 'Backspace') return false;
    if (event.metaKey || event.ctrlKey || event.altKey) return false;

    const text = editor.value;
    const selectionStart = Math.max(0, Math.min(editor.selectionStart ?? text.length, text.length));
    const selectionEnd = Math.max(0, Math.min(editor.selectionEnd ?? text.length, text.length));
    const target = this.activeShapeTextTarget;
    const rangeSelection = target && this.inlineRangeSelection?.shapeIndex === target.shapeIndex
      ? this.inlineRangeSelection
      : null;
    const selectedCurrentParagraphOnly = target !== null
      && rangeSelection?.ranges.length === 1
      && rangeSelection.ranges[0]?.paragraphIndex === target.paragraphIndex;
    const replacesWholeShape = target !== null
      && rangeSelection !== null
      && this.isWholeShapeInlineRangeSelection(target.shapeIndex, rangeSelection.ranges);
    if (
      target
      && rangeSelection
      && rangeSelection.ranges.some((range) => range.end > range.start)
      && !selectedCurrentParagraphOnly
      && !replacesWholeShape
    ) {
      event.preventDefault();
      event.stopPropagation();
      debugLog('text-edit', 'Inline Delete/Backspace requested selected text range removal', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
        key: event.key,
        rangeCount: rangeSelection.ranges.length,
      });
      this.startInlineRangeDeletion(editor, target);
      return true;
    }
    if (replacesWholeShape && target) {
      if (!this.beginInlineWholeShapeReplacement(target)) {
        this.startInlineRangeDeletion(editor, target);
        return true;
      }
      this.clearWholeShapeInlineSelection();
      editor.setSelectionRange(0, editor.value.length);
    }
    // A selection wholly inside the active paragraph is already represented by
    // the textarea range. Keep it on the synchronous deletion path so Ctrl/Cmd+A
    // and multi-click deletion repaint before the asynchronous OOXML commit.
    if (selectedCurrentParagraphOnly) {
      this.clearWholeShapeInlineSelection();
    }
    if (
      event.key === 'Backspace'
      && selectionStart === 0
      && selectionEnd === 0
      && target?.element === element
      && this.activeEditor === editor
    ) {
      if (target.paragraphIndex === 0) {
        debugLog('text-edit', 'Inline Backspace skipped paragraph removal', {
          slide: this.currentSlide,
          shapeIndex: target.shapeIndex,
          paragraphIndex: target.paragraphIndex,
          reason: 'no-previous-paragraph',
        });
        return false;
      }
      event.preventDefault();
      event.stopPropagation();
      const hasEmptyPrecedingParagraph = this.engine?.hasEmptyPrecedingParagraph(
        this.currentSlide,
        target.shapeIndex,
        target.paragraphIndex,
      ) === true;
      debugLog('text-edit', 'Inline Backspace requested preceding paragraph operation', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
        textLength: text.length,
        hasEmptyPrecedingParagraph,
      });
      if (hasEmptyPrecedingParagraph) {
        this.startInlineEmptyParagraphRemoval(editor, target);
      } else {
        this.startInlinePrecedingParagraphMerge(editor, target);
      }
      return true;
    }

    if (!text || event.isComposing) return false;

    // The editor textarea is intentionally invisible. Letting Chromium apply
    // a native deletion there can leave it empty while its SVG preview still
    // shows the old paragraph; a later blur then silently saves the empty
    // string. Apply every ordinary deletion here so textarea, preview, and
    // commit state advance as one transaction.
    const hasSelection = selectionEnd > selectionStart;
    const deleteStart = hasSelection
      ? selectionStart
      : event.key === 'Backspace' ? selectionStart - 1 : selectionStart;
    const deleteEnd = hasSelection
      ? selectionEnd
      : event.key === 'Backspace' ? selectionStart : selectionStart + 1;
    if (deleteStart < 0 || deleteEnd > text.length || deleteStart >= deleteEnd) return false;

    event.preventDefault();
    event.stopPropagation();

    const scrollPosition = this.captureCanvasScroll();
    this.recordInlineEditSnapshot(editor);
    const nextText = text.slice(0, deleteStart) + text.slice(deleteEnd);
    const deletedWholeParagraph = hasSelection && deleteStart === 0 && deleteEnd === text.length;
    // Any destructive edit can hit Chromium's stale-glyph path. Repaint the
    // owning <text> boundary rather than hoping a nested run replacement wins.
    const replaceTextFrame = true;
    this.setInlineEditorValue(editor, nextText);
    this.activeEditorTextDirty = true;
    let textFrameReplaced = false;
    if (this.activeShapeTextTarget) {
      textFrameReplaced = this.syncInlineTextPreviewSafely(this.activeShapeTextTarget, nextText, {
        replaceTextFrame,
      });
    } else if (isSVGTextElement(element)) {
      element.textContent = nextText;
    }
    const liveElement = this.activeEditorTarget ?? element;
    editor.setSelectionRange(deleteStart, deleteStart);
    this.rememberInlineCaretPlacement(editor, liveElement, deleteStart);
    this.resetInlineEditorScroll(editor);
    const nextBox = this.getElementBox(liveElement);
    if (nextBox) {
      this.positionTextRunEditor(editor, nextBox);
    }
    this.updateInlineCaret(editor, liveElement);
    this.preserveCanvasScrollAfterInlineTextEdit(
      scrollPosition,
      event.key === 'Backspace' ? 'deleteContentBackward' : 'deleteContentForward'
    );
    debugLog('text-edit', 'Applied inline text deletion to live preview', {
      slide: this.currentSlide,
      shapeIndex: target?.shapeIndex ?? null,
      paragraphIndex: target?.paragraphIndex ?? null,
      key: event.key,
      deleteStart,
      deleteEnd,
      previousTextLength: text.length,
      nextTextLength: nextText.length,
      deletedWholeParagraph,
      replaceTextFrame,
      textFrameReplaced,
    });
    return true;
  }

  private snapshotInlineEdit(editor: HTMLTextAreaElement): InlineEditSnapshot {
    return {
      value: editor.value,
      selectionStart: editor.selectionStart ?? editor.value.length,
      selectionEnd: editor.selectionEnd ?? editor.value.length
    };
  }

  /** Keep commit-time text independent from Electron's invisible textarea repaint quirks. */
  private setInlineEditorValue(editor: HTMLTextAreaElement, text: string): void {
    editor.value = text;
    if (this.activeEditor === editor) {
      this.activeEditorCanonicalText = text;
    }
  }

  private resolveInlineTextForCommit(editor: HTMLTextAreaElement): string {
    const editorText = editor.value;
    const canonicalText = this.activeEditor === editor ? this.activeEditorCanonicalText : null;
    if (canonicalText === null || canonicalText === editorText) return editorText;

    warnLog('text-edit', 'Inline editor text diverged before commit', {
      slide: this.currentSlide,
      editorTextLength: editorText.length,
      canonicalTextLength: canonicalText.length,
    });
    return canonicalText;
  }

  private hasInlineHistory(action: 'undo' | 'redo'): boolean {
    return this.activeEditor !== null
      && (action === 'undo' ? this.inlineUndoStack.length > 0 : this.inlineRedoStack.length > 0);
  }

  /** Push the current editor state onto the in-place undo stack. */
  private recordInlineEditSnapshot(editor: HTMLTextAreaElement): void {
    this.inlineUndoStack.push(this.snapshotInlineEdit(editor));
    if (this.inlineUndoStack.length > INLINE_EDIT_HISTORY_LIMIT) {
      this.inlineUndoStack.shift();
    }
    this.inlineRedoStack = [];
    this.historyController.updateAvailability();
  }

  /**
   * Reapply a captured snapshot to the live editor: restore text, the selection
   * range, and the on-slide preview/caret/selection overlays.
   */
  private applyInlineSnapshot(editor: HTMLTextAreaElement, snapshot: InlineEditSnapshot): void {
    this.setInlineEditorValue(editor, snapshot.value);
    const length = snapshot.value.length;
    const start = Math.max(0, Math.min(snapshot.selectionStart, length));
    const end = Math.max(start, Math.min(snapshot.selectionEnd, length));

    this.clearWholeShapeInlineSelection();
    const target = this.activeShapeTextTarget;
    if (target) {
      this.syncInlineTextPreviewSafely(target, snapshot.value, { replaceTextFrame: true });
    } else if (this.activeEditorTarget && isSVGTextElement(this.activeEditorTarget)) {
      this.activeEditorTarget.textContent = snapshot.value;
    }

    editor.setSelectionRange(start, end);

    const element = this.activeEditorTarget;
    if (element) {
      const box = this.getElementBox(element);
      if (box) {
        this.positionTextRunEditor(editor, box);
      }
      this.rememberInlineCaretPlacement(editor, element, end);
      this.updateInlineCaret(editor, element);
      this.updateInlineSelection(editor, element);
    }
  }

  /**
   * In-place undo for the active text editor. Returns false when there is
   * nothing left to undo within the current edit session, so the caller can
   * fall back to document-level history.
   */
  private undoInlineEdit(editor: HTMLTextAreaElement): boolean {
    return this.restoreInlineHistorySnapshot('undo', this.inlineUndoStack, this.inlineRedoStack, editor);
  }

  private redoInlineEdit(editor: HTMLTextAreaElement): boolean {
    return this.restoreInlineHistorySnapshot('redo', this.inlineRedoStack, this.inlineUndoStack, editor);
  }

  /**
   * Apply an inline history snapshot before moving it between stacks. A failed
   * SVG preview restore must leave the snapshot retryable instead of silently
   * consuming the user's only undo step.
   */
  private restoreInlineHistorySnapshot(
    action: 'undo' | 'redo',
    source: InlineEditSnapshot[],
    destination: InlineEditSnapshot[],
    editor: HTMLTextAreaElement,
  ): boolean {
    this.lastInlineHistoryRestoreFailed = false;
    const snapshot = source[source.length - 1];
    if (!snapshot) return false;

    const current = this.snapshotInlineEdit(editor);
    debugLog('text-edit', `inline ${action} started`, {
      sourceDepth: source.length,
      destinationDepth: destination.length,
      targetTextLength: snapshot.value.length,
    });
    try {
      this.applyInlineSnapshot(editor, snapshot);
    } catch (error) {
      this.lastInlineHistoryRestoreFailed = true;
      errorLog('text-edit', `inline ${action} failed`, {
        sourceDepth: source.length,
        destinationDepth: destination.length,
        error: cleanError(error),
      });
      this.historyController.updateAvailability();
      return false;
    }

    source.pop();
    destination.push(current);
    if (destination.length > INLINE_EDIT_HISTORY_LIMIT) {
      destination.shift();
    }
    this.historyController.updateAvailability();
    debugLog('text-edit', `inline ${action}`, {
      selectionStart: snapshot.selectionStart,
      selectionEnd: snapshot.selectionEnd,
      sourceDepth: source.length,
      destinationDepth: destination.length,
    });
    return true;
  }

  /** Route every user Undo/Redo surface through inline history first. */
  private async requestHistoryAction(
    action: 'undo' | 'redo',
    source: 'keyboard' | 'toolbar' | 'menu',
  ): Promise<void> {
    const editor = this.activeEditor;
    const inlineAvailable = this.hasInlineHistory(action);
    const documentAvailable = action === 'undo'
      ? this.historyController.canUndo
      : this.historyController.canRedo;
    const route = inlineAvailable
      ? 'inline'
      : editor
        ? 'commit-then-document'
        : 'document';
    logPptxAction('history', action, {
      source,
      route,
      activeEditor: editor !== null,
      inlineUndoDepth: this.inlineUndoStack.length,
      inlineRedoDepth: this.inlineRedoStack.length,
      documentAvailable,
    });

    if (editor) {
      const appliedInline = action === 'undo'
        ? this.undoInlineEdit(editor)
        : this.redoInlineEdit(editor);
      if (appliedInline || this.lastInlineHistoryRestoreFailed) return;
      await this.finishInlineTextEditing(`inline-${action}-fallback`);
    }

    const dispatched = action === 'undo' ? this.session.undo() : this.session.redo();
    debugLog('history', 'PowerPoint document history dispatch', {
      action,
      source,
      dispatched,
      documentAvailable,
    });
  }

  /** Ctrl+Z while editing text first unwinds the in-place editor history. */
  private async handleInlineUndo(): Promise<void> {
    await this.requestHistoryAction('undo', 'keyboard');
  }

  private async handleInlineRedo(): Promise<void> {
    await this.requestHistoryAction('redo', 'keyboard');
  }

  private rememberCollapsedInlineCaretPlacement(
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement
  ): void {
    const textLength = editor.value.length;
    const selectionStart = Math.max(0, Math.min(editor.selectionStart ?? textLength, textLength));
    const selectionEnd = Math.max(0, Math.min(editor.selectionEnd ?? textLength, textLength));
    if (selectionStart === selectionEnd) {
      this.rememberInlineCaretPlacement(editor, element, selectionEnd);
    }
  }

  private updateInlineCaret(editor: HTMLTextAreaElement, element: SVGTextElement | SVGTSpanElement): void {
    if (!this.activeInlineCaret) return;
    if (this.isNavigatingSlide || this.isTearingDownEditor) return;

    const box = this.getElementBox(element);
    if (!box) {
      this.activeInlineCaret.addClass('native-powerpoint-inline-caret-hidden');
      return;
    }

    this.refreshActiveInlineCaretRow(element, box);
    this.updateInlineSelection(editor, element);

    if (this.inlineRangeSelection) {
      this.activeInlineCaret.addClass('native-powerpoint-inline-caret-hidden');
      return;
    }

    const selectionStart = Math.max(0, Math.min(editor.selectionStart ?? editor.value.length, editor.value.length));
    const selectionEnd = Math.max(0, Math.min(editor.selectionEnd ?? editor.value.length, editor.value.length));
    if (selectionStart !== selectionEnd) {
      this.activeInlineCaret.addClass('native-powerpoint-inline-caret-hidden');
      return;
    }

    const offset = selectionEnd;
    const geometry = this.getSvgInlineCaretGeometry(element, editor, offset, box);
    if (!geometry) {
      this.activeInlineCaret.addClass('native-powerpoint-inline-caret-hidden');
      return;
    }

    this.activeInlineCaret.removeClass('native-powerpoint-inline-caret-hidden');
    this.activeInlineCaret.setAttribute('x1', this.formatSvgNumber(geometry.x1));
    this.activeInlineCaret.setAttribute('y1', this.formatSvgNumber(geometry.y1));
    this.activeInlineCaret.setAttribute('x2', this.formatSvgNumber(geometry.x2));
    this.activeInlineCaret.setAttribute('y2', this.formatSvgNumber(geometry.y2));
    this.activeInlineCaret.setAttribute('stroke-width', this.formatSvgNumber(geometry.strokeWidth));
  }

  private updateInlineSelection(editor: HTMLTextAreaElement, element: SVGTextElement | SVGTSpanElement): void {
    this.removeInlineSelection();

    if (this.inlineWholeShapeSelected) {
      this.renderWholeShapeInlineSelection();
      return;
    }

    if (this.inlineRangeSelection) {
      this.renderInlineRangeSelection(this.inlineRangeSelection);
      return;
    }

    const selectionStart = Math.max(0, Math.min(editor.selectionStart ?? editor.value.length, editor.value.length));
    const selectionEnd = Math.max(0, Math.min(editor.selectionEnd ?? editor.value.length, editor.value.length));
    const start = Math.min(selectionStart, selectionEnd);
    const end = Math.max(selectionStart, selectionEnd);
    if (start === end) return;

    const target = this.activeShapeTextTarget;
    if (target && this.needsFlatParagraphMapping(target.shapeIndex, target.paragraphIndex, element)) {
      this.renderFlatParagraphSelection(target.shapeIndex, target.paragraphIndex, start, end);
      return;
    }

    this.renderInlineSelectionRects(element, start, end);
  }

  private renderInlineSelectionRects(
    element: SVGTextElement | SVGTSpanElement,
    start: number,
    end: number
  ): void {
    const boxes = this.getSvgInlineSelectionBoxes(element, start, end);
    const textElement = element.closest('text');
    const parent = textElement?.parentNode;
    if (!isSVGTextElement(textElement) || !parent) return;

    for (const box of boxes) {
      const rect = this.getSvgFactory(textElement).createSvg('rect');
      rect.classList.add('native-powerpoint-svg-selection');
      rect.setAttribute('x', this.formatSvgNumber(box.x));
      rect.setAttribute('y', this.formatSvgNumber(box.y));
      rect.setAttribute('width', this.formatSvgNumber(box.width));
      rect.setAttribute('height', this.formatSvgNumber(box.height));
      parent.insertBefore(rect, textElement);
      this.activeInlineSelectionRects.push(rect);
    }
  }

  private getShapeTextParagraphs(shape: Element): (SVGTextElement | SVGTSpanElement)[] {
    const result: (SVGTextElement | SVGTSpanElement)[] = [];
    for (const text of Array.from(shape.querySelectorAll('text'))) {
      if (text.closest(GENERATED_GRID_SELECTOR)) continue;
      const paragraphs = Array.from(text.querySelectorAll('tspan[data-ooxml-para-idx]')).filter(isSVGTSpanElement);
      if (paragraphs.length > 0) {
        result.push(...paragraphs);
      } else if (isSVGTextElement(text) && (text.textContent ?? '').length > 0) {
        result.push(text);
      }
    }
    return result;
  }

  private getShapeTextRanges(shapeIndex: number): ParagraphTextRange[] {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!shape) return [];

    const ranges: ParagraphTextRange[] = [];
    const offsets = new Map<number, number>();
    for (const paragraph of this.getShapeTextParagraphs(shape)) {
      const paragraphIndex = this.getParagraphIndexFromInlineElement(paragraph);
      if (paragraphIndex === null) continue;
      // Bullet/number marker containers carry no runs; excluding them keeps the
      // accumulated offsets in the run-only space the OOXML engine expects.
      if (this.collectParagraphRuns(paragraph).length === 0) continue;

      const total = this.getRunCharInfo(paragraph).total;
      if (total > 0) {
        const start = offsets.get(paragraphIndex) ?? 0;
        const end = start + total;
        ranges.push({ paragraphIndex, start, end });
        offsets.set(paragraphIndex, end);
      }
    }
    return ranges;
  }

  /**
   * Stamp every run tspan with its OOXML char range exactly once per render.
   *
   * The renderer drops the whitespace PowerPoint swallows at soft-wrap
   * boundaries and splits a single run across visual lines, so the SVG run text
   * is a subsequence of the OOXML run text. Rather than re-infer that alignment
   * inside selection, caret, highlight, find, and formatting independently (the
   * historic source of offset-drift bugs), we infer it here, in one place, and
   * write the result onto the DOM as `data-ooxml-char-start`/`-char-end`. Every
   * consumer then reads the stamp instead of re-deriving it. If the alignment
   * fails to reconcile (editor text not a clean subsequence, or coverage does
   * not reach the OOXML end) we log loudly so drift surfaces immediately.
   */
  private annotateSlideTextOffsets(): void {
    const svg = this.svgEl;
    const engine = this.engine;
    if (!svg || !engine) return;

    const slideIndex = this.currentSlide;
    const stamped = stampSlideTextOffsets(svg, (shapeIndex, paragraphIndex) =>
      engine.getParagraphRunText(slideIndex, shapeIndex, paragraphIndex),
    );
    for (const shape of stamped) {
      for (const paragraph of shape.paragraphs) {
        if (paragraph.reconciled) continue;
        warnLog('text-format', 'text offset annotation did not reconcile', {
          slide: slideIndex,
          shapeIndex: shape.shapeIndex,
          paragraphIndex: paragraph.paragraphIndex,
          spanCount: paragraph.spanCount,
          editorLength: paragraph.editorLength,
          ooxmlLength: paragraph.ooxmlLength,
        });
      }
    }
    for (const shapeGroup of Array.from(svg.querySelectorAll('g[data-ooxml-shape-idx]'))) {
      this.warnGlyphCountDivergence(shapeGroup);
    }
  }

  /**
   * Stamp `data-ooxml-char-start/end` onto every run tspan inside a single shape
   * group. Extracted from {@link annotateSlideTextOffsets} so the incremental
   * (single-shape) re-render path can re-annotate just the shape it swapped in.
   */
  private annotateShapeTextOffsets(shapeGroup: Element): void {
    const engine = this.engine;
    if (!engine) return;

    const shapeIndex = Number(shapeGroup.getAttribute('data-ooxml-shape-idx'));
    if (!Number.isFinite(shapeIndex)) return;

    const slideIndex = this.currentSlide;
    try {
      const stamped = annotateShapeGroupTextOffsets(shapeGroup, (paragraphIndex) =>
        engine.getParagraphRunText(slideIndex, shapeIndex, paragraphIndex),
      );
      if (!stamped) return;
      for (const paragraph of stamped.paragraphs) {
        if (paragraph.reconciled) continue;
        warnLog('text-format', 'text offset annotation did not reconcile', {
          slide: slideIndex,
          shapeIndex,
          paragraphIndex: paragraph.paragraphIndex,
          spanCount: paragraph.spanCount,
          editorLength: paragraph.editorLength,
          ooxmlLength: paragraph.ooxmlLength,
        });
      }
      this.warnGlyphCountDivergence(shapeGroup);
    } catch (error) {
      warnLog('text-format', 'text offset annotation failed', {
        slide: slideIndex,
        shapeIndex,
        error: cleanError(error),
      });
    }
  }

  /** Warn when rendered glyph counts diverge from run string lengths (shaping drift). */
  private warnGlyphCountDivergence(shapeGroup: Element): void {
    const shapeIndex = Number(shapeGroup.getAttribute('data-ooxml-shape-idx'));
    if (!Number.isFinite(shapeIndex)) return;

    for (const [paragraphIndex, runSpans] of collectRunSpansByParagraph(shapeGroup)) {
      const stringChars = runSpans.reduce((sum, span) => sum + (span.textContent || '').length, 0);
      const glyphChars = runSpans.reduce(
        (sum, span) => sum + this.inlineGeometry.getGlyphCount(span as SVGTextContentElement),
        0,
      );
      if (stringChars === glyphChars) continue;
      warnLog('text-select', 'caret/find glyph count diverges from text stamps', {
        slide: this.currentSlide,
        shapeIndex,
        paragraphIndex,
        spanCount: runSpans.length,
        stringChars,
        glyphChars,
      });
    }
  }

  /**
   * Map an inline-editor (SVG run-offset) range to OOXML offsets by reading the
   * per-render `data-ooxml-char-*` stamps. Returns null when the paragraph is not
   * annotated, so the caller can fall back to live inference. Within a tspan the
   * editor and OOXML text agree (drops only happen at wrap boundaries), so the
   * mapping is linear; START lands on the first real glyph and END absorbs any
   * trailing dropped whitespace — matching the engine's clear/format semantics.
   */
  private mapEditorRangeViaAnnotations(
    shapeIndex: number,
    paragraphIndex: number,
    editorStart: number,
    editorEnd: number
  ): { start: number; end: number } | null {
    const svg = this.svgEl;
    if (!svg) return null;

    const runSpans = Array.from(
      svg.querySelectorAll(
        `g[data-ooxml-shape-idx="${shapeIndex}"] tspan[data-ooxml-para-idx="${paragraphIndex}"] tspan[data-ooxml-run-idx]`
      )
    ).filter(isSVGTSpanElement);
    if (runSpans.length === 0) return null;

    const tiles: RunTspanOffset[] = [];
    let cursor = 0;
    for (const span of runSpans) {
      const rawStart = span.getAttribute('data-ooxml-char-start');
      const rawEnd = span.getAttribute('data-ooxml-char-end');
      if (rawStart === null || rawEnd === null) return null;
      const length = (span.textContent || '').length;
      tiles.push({ editorStart: cursor, editorEnd: cursor + length, charStart: Number(rawStart), charEnd: Number(rawEnd) });
      cursor += length;
    }

    return mapEditorRangeToOoxml(tiles, editorStart, editorEnd);
  }

  /**
   * Convert paragraph ranges expressed in the inline editor's SVG-run offset
   * space into the authoritative OOXML run offset space the engine styles in.
   *
   * Prefers the per-render `data-ooxml-char-*` stamps (single validated source
   * of truth); falls back to live subsequence inference only when a paragraph is
   * not annotated. The editor text omits whitespace swallowed at soft-wrap
   * boundaries, so without this mapping a clear/format range lands short of the
   * real OOXML runs and leaves trailing characters (and wrap-spaces) untouched.
   */
  private mapRangesToOoxmlOffsets(
    shapeIndex: number,
    ranges: ParagraphTextRange[]
  ): ParagraphTextRange[] {
    const engine = this.engine;
    if (!engine) return ranges;

    return ranges.map((range) => {
      const viaAnnotations = this.mapEditorRangeViaAnnotations(
        shapeIndex,
        range.paragraphIndex,
        range.start,
        range.end
      );
      if (viaAnnotations) {
        if (viaAnnotations.start === range.start && viaAnnotations.end === range.end) return range;
        debugLog('text-select', 'mapped editor range to OOXML offsets (annotation)', {
          shapeIndex,
          paragraphIndex: range.paragraphIndex,
          editorStart: range.start,
          editorEnd: range.end,
          ooxmlStart: viaAnnotations.start,
          ooxmlEnd: viaAnnotations.end
        });
        return { paragraphIndex: range.paragraphIndex, start: viaAnnotations.start, end: viaAnnotations.end };
      }

      const ooxmlText = engine.getParagraphRunText(this.currentSlide, shapeIndex, range.paragraphIndex);
      const editorText = this.getParagraphPlainText(
        this.getRunLineContainers(shapeIndex, range.paragraphIndex)
      );
      if (ooxmlText === null) return range;
      if (editorText === ooxmlText) return range;

      const mappedStart = mapEditorOffsetToOoxmlOffset(editorText, ooxmlText, range.start, false);
      const mappedEnd = mapEditorOffsetToOoxmlOffset(editorText, ooxmlText, range.end, true);
      if (mappedStart === range.start && mappedEnd === range.end) return range;

      debugLog('text-select', 'mapped editor range to OOXML offsets (fallback)', {
        shapeIndex,
        paragraphIndex: range.paragraphIndex,
        editorStart: range.start,
        editorEnd: range.end,
        ooxmlStart: mappedStart,
        ooxmlEnd: mappedEnd,
        editorLen: editorText.length,
        ooxmlLen: ooxmlText.length
      });
      return { paragraphIndex: range.paragraphIndex, start: mappedStart, end: mappedEnd };
    });
  }

  private renderInlineRangeSelection(selection: InlineRangeSelection): void {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${selection.shapeIndex}"]`);
    if (!shape) return;

    const paragraphs = this.getShapeTextParagraphs(shape);
    const offsets = new Map<number, number>();
    for (const paragraph of paragraphs) {
      const paragraphIndex = this.getParagraphIndexFromInlineElement(paragraph);
      if (paragraphIndex === null) continue;
      if (this.collectParagraphRuns(paragraph).length === 0) continue;

      const runTotal = this.getRunCharInfo(paragraph).total;
      const visualStart = offsets.get(paragraphIndex) ?? 0;
      const visualEnd = visualStart + runTotal;
      offsets.set(paragraphIndex, visualEnd);

      for (const range of selection.ranges) {
        if (range.paragraphIndex !== paragraphIndex) continue;
        const runStart = Math.max(0, range.start - visualStart);
        const runEnd = Math.min(runTotal, range.end - visualStart);
        if (runEnd > runStart) {
          const geometryStart = this.runOffsetToGeometryIndex(paragraph, runStart);
          const geometryEnd = this.runOffsetToGeometryIndex(paragraph, runEnd);
          this.renderInlineSelectionRects(paragraph, geometryStart, geometryEnd);
        }
      }
    }
  }

  private renderWholeShapeInlineSelection(): void {
    const shape = this.getSelectedShapeElement();
    if (!shape) return;
    for (const paragraph of this.getShapeTextParagraphs(shape)) {
      const total = this.getLeafCharInfo(paragraph).total;
      if (total <= 0) continue;
      this.renderInlineSelectionRects(paragraph, 0, total);
    }
  }

  /** True only when the visual selection covers every editable character in its text shape. */
  private isWholeShapeInlineRangeSelection(
    shapeIndex: number,
    ranges: readonly ParagraphTextRange[],
  ): boolean {
    const shapeRanges = this.getShapeTextRanges(shapeIndex);
    return shapeRanges.length > 0
      && shapeRanges.length === ranges.length
      && shapeRanges.every((shapeRange, index) => {
        const range = ranges[index];
        return range?.paragraphIndex === shapeRange.paragraphIndex
          && range.start === shapeRange.start
          && range.end === shapeRange.end;
      });
  }

  /**
   * Keep a full-text replacement entirely in the inline session until commit.
   * Rendering an all-empty OOXML shape removes its `<text>` node, which would
   * otherwise orphan the textarea before the user can type the replacement.
   */
  private beginInlineWholeShapeReplacement(target: ShapeTextEditTarget): boolean {
    if (this.inlineWholeShapeReplacement?.shapeIndex === target.shapeIndex) return true;

    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${target.shapeIndex}"]`);
    const textFrame = target.element.closest('text');
    if (!shape || !isSVGTextElement(textFrame)) {
      warnLog('text-edit', 'Could not prepare whole-shape inline replacement preview', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
        hasShape: Boolean(shape),
        hasTextFrame: isSVGTextElement(textFrame),
      });
      return false;
    }

    this.inlineWholeShapeReplacement = {
      shapeIndex: target.shapeIndex,
      textFrame: textFrame.cloneNode(true) as SVGTextElement,
      // Capture before clearing runs for the temporary all-text preview.
      baselineText: textFrame.textContent ?? '',
    };
    const runs = Array.from(shape.querySelectorAll('tspan[data-ooxml-run-idx]')).filter(isSVGTSpanElement);
    for (const run of runs) run.textContent = '';
    debugLog('text-edit', 'Prepared whole-shape inline text replacement', {
      slide: this.currentSlide,
      shapeIndex: target.shapeIndex,
      paragraphIndex: target.paragraphIndex,
      runCount: runs.length,
      baselineLength: this.inlineWholeShapeReplacement.baselineText.length,
    });
    return true;
  }

  /** Restore the SVG snapshot used by an in-progress whole-shape replacement. */
  private restoreInlineWholeShapeReplacementPreview(shapeIndex: number): boolean {
    const replacement = this.inlineWholeShapeReplacement;
    if (!replacement || replacement.shapeIndex !== shapeIndex) return false;

    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!shape) return false;

    const restoredFrame = replacement.textFrame.cloneNode(true) as SVGTextElement;
    const currentFrame = shape.querySelector('text');
    if (isSVGTextElement(currentFrame)) {
      currentFrame.replaceWith(restoredFrame);
    } else {
      shape.appendChild(restoredFrame);
    }
    debugLog('text-edit', 'Restored whole-shape inline text replacement preview', {
      slide: this.currentSlide,
      shapeIndex,
    });
    return true;
  }

  private selectAllInlineText(editor: HTMLTextAreaElement, element: SVGTextElement | SVGTSpanElement): void {
    const shape = this.getSelectedShapeElement();
    if (!shape) return;

    const paragraphs = this.getShapeTextParagraphs(shape);
    const combined = paragraphs.map((paragraph) => paragraph.textContent ?? '').join('\n');
    this.inlineWholeShapeSelection = combined;
    this.inlineWholeShapeSelected = true;
    const shapeIndex = this.activeShapeTextTarget?.shapeIndex ?? this.selectedShapeIndex;
    this.inlineRangeSelection = shapeIndex !== null
      ? { shapeIndex, ranges: this.getShapeTextRanges(shapeIndex) }
      : null;
    this.lastInlineCaretPlacement = null;
    editor.setSelectionRange(0, editor.value.length);
    this.updateInlineCaret(editor, element);
  }

  private clearWholeShapeInlineSelection(): void {
    this.inlineWholeShapeSelection = null;
    this.inlineWholeShapeSelected = false;
    this.inlineRangeSelection = null;
    this.toolbarFormattingSnapshot = null;
  }

  private removeInlineSelection(): void {
    for (const rect of this.activeInlineSelectionRects) {
      rect.remove();
    }
    this.activeInlineSelectionRects = [];
  }

  private getSvgInlineCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    editor: HTMLTextAreaElement,
    offset: number,
    box: { left: number; top: number; width: number; height: number }
  ): SvgInlineCaretGeometry | null {
    const screenGeometry = this.getInlineCaretGeometry(element, editor, offset, box);
    const top = screenGeometry.top;
    const height = screenGeometry.height;
    const start = this.localPointToSvgRoot(screenGeometry.left, top);
    const end = this.localPointToSvgRoot(screenGeometry.left, top + height);
    if (!start || !end) return null;

    return {
      x1: start.x,
      y1: start.y,
      x2: end.x,
      y2: end.y,
      strokeWidth: this.getSvgInlineCaretStrokeWidth(element)
    };
  }

  /**
   * Repaint run-level text highlights as overlay rects. The SVG renderer drops
   * <a:highlight>, so after every slide render we query the engine for runs that
   * carry a highlight color and draw a colored rect behind each run's glyphs
   * (mirroring the find-match highlighting). Soft-wrapped runs expose one tspan
   * per visual line, so each is boxed independently.
   */
  private applyRunHighlights(): void {
    for (const rect of this.runHighlightRects) {
      rect.remove();
    }
    this.runHighlightRects = [];

    const svg = this.svgEl;
    if (!svg || !this.engine) return;

    const highlights = this.engine.getSlideRunHighlights(this.currentSlide);
    debugLog('text-format', 'applyRunHighlights', {
      slide: this.currentSlide,
      count: highlights.length,
      runs: highlights.map((entry) => {
        const svgOffsets = this.getParagraphRunOffsets(entry.shapeIndex, entry.paragraphIndex).get(entry.runIndex);
        return {
          shapeIndex: entry.shapeIndex,
          paragraphIndex: entry.paragraphIndex,
          runIndex: entry.runIndex,
          color: entry.color,
          ooxmlStart: entry.start,
          ooxmlEnd: entry.end,
          svgStart: svgOffsets?.start ?? null,
          svgEnd: svgOffsets?.end ?? null,
          text: svg
            .querySelector(
              `g[data-ooxml-shape-idx="${entry.shapeIndex}"] tspan[data-ooxml-para-idx="${entry.paragraphIndex}"] tspan[data-ooxml-run-idx="${entry.runIndex}"]`
            )
            ?.textContent ?? null
        };
      })
    });

    this.detectResidualHighlightAfterClear(highlights);

    if (!highlights.length) return;

    for (const { shapeIndex, paragraphIndex, runIndex, color } of highlights) {
      const runSpans = Array.from(
        svg.querySelectorAll(
          `g[data-ooxml-shape-idx="${shapeIndex}"] tspan[data-ooxml-para-idx="${paragraphIndex}"] tspan[data-ooxml-run-idx="${runIndex}"]`
        )
      ).filter(isSVGTSpanElement);

      for (const span of runSpans) {
        const total = this.getLeafCharInfo(span).total;
        if (total <= 0) continue;

        const textElement = span.closest('text');
        const parent = textElement?.parentNode;
        if (!isSVGTextElement(textElement) || !parent) continue;

        for (const box of this.getSvgInlineSelectionBoxes(span, 0, total)) {
          const rect = this.getSvgFactory(textElement).createSvg('rect');
          rect.classList.add('native-powerpoint-run-highlight');
          rect.setAttribute('x', this.formatSvgNumber(box.x));
          rect.setAttribute('y', this.formatSvgNumber(box.y));
          rect.setAttribute('width', this.formatSvgNumber(box.width));
          rect.setAttribute('height', this.formatSvgNumber(box.height));
          rect.setAttribute('rx', '1');
          rect.setAttribute('fill', `#${color}`);
          parent.insertBefore(rect, textElement);
          this.runHighlightRects.push(rect);
        }
      }
    }
  }

  /**
   * After a "No color" highlight clear re-renders, check whether any highlight
   * survived in the cleared shape/paragraphs and log it explicitly. This catches
   * the common case where the user's selection stopped short of the end of the
   * line, so a trailing (or leading) chunk stayed highlighted outside the
   * cleared range. Offsets are run-only (paragraph-relative), matching the range
   * that was cleared, so the log makes the residual vs cleared gap obvious.
   */
  private detectResidualHighlightAfterClear(highlights: RunHighlightInfo[]): void {
    const request = this.pendingHighlightClear;
    if (!request) return;
    if (request.slide !== this.currentSlide) {
      this.pendingHighlightClear = null;
      return;
    }
    this.pendingHighlightClear = null;

    const residual = highlights
      .filter((entry) => entry.shapeIndex === request.shapeIndex && request.paragraphs.has(entry.paragraphIndex))
      .map((entry) => {
        const svgOffsets = this.getParagraphRunOffsets(entry.shapeIndex, entry.paragraphIndex).get(entry.runIndex);
        return {
          paragraphIndex: entry.paragraphIndex,
          runIndex: entry.runIndex,
          color: entry.color,
          ooxmlStart: entry.start,
          ooxmlEnd: entry.end,
          svgStart: svgOffsets?.start ?? null,
          svgEnd: svgOffsets?.end ?? null,
          text: this.svgEl
            ?.querySelector(
              `g[data-ooxml-shape-idx="${entry.shapeIndex}"] tspan[data-ooxml-para-idx="${entry.paragraphIndex}"] tspan[data-ooxml-run-idx="${entry.runIndex}"]`
            )
            ?.textContent ?? null
        };
      });

    if (residual.length) {
      warnLog('text-format', 'residual highlight remains after No color clear', {
        slide: this.currentSlide,
        shapeIndex: request.shapeIndex,
        clearedRanges: request.ranges,
        residual
      });
    } else {
      debugLog('text-format', 'No color clear fully removed highlight', {
        slide: this.currentSlide,
        shapeIndex: request.shapeIndex,
        clearedRanges: request.ranges
      });
    }
  }

  /**
   * Map each OOXML run index in a rendered paragraph to its run-only,
   * paragraph-relative `[start, end)` character offsets. Soft-wrapped runs span
   * several line tspans that share a run index; their counts accumulate. Bullet
   * markers live in run-less tspans and are excluded, so offsets match the
   * engine's run-only space (the same space toolbar selections use).
   */
  private getParagraphRunOffsets(
    shapeIndex: number,
    paragraphIndex: number
  ): Map<number, { start: number; end: number }> {
    const offsets = new Map<number, { start: number; end: number }>();
    const svg = this.svgEl;
    if (!svg) return offsets;

    const runSpans = Array.from(
      svg.querySelectorAll(
        `g[data-ooxml-shape-idx="${shapeIndex}"] tspan[data-ooxml-para-idx="${paragraphIndex}"] tspan[data-ooxml-run-idx]`
      )
    ).filter(isSVGTSpanElement);

    let cursor = 0;
    for (const span of runSpans) {
      const runIndex = Number(span.getAttribute('data-ooxml-run-idx'));
      const count = this.getLeafCharInfo(span).total;
      const existing = offsets.get(runIndex);
      if (existing) {
        existing.end = cursor + count;
      } else {
        offsets.set(runIndex, { start: cursor, end: cursor + count });
      }
      cursor += count;
    }
    return offsets;
  }

  private getSvgInlineSelectionBoxes(
    element: SVGTextElement | SVGTSpanElement,
    start: number,
    end: number
  ): SvgInlineSelectionBox[] {
    const rootMatrix = this.svgEl?.getScreenCTM();
    if (!rootMatrix) return [];

    let rootInverse: DOMMatrix;
    try {
      rootInverse = rootMatrix.inverse();
    } catch {
      return [];
    }

    const { entries, total } = this.getLeafCharInfo(element);
    if (total <= 0) return [];

    const normalizedStart = Math.max(0, Math.min(total, start));
    const normalizedEnd = Math.max(normalizedStart, Math.min(total, end));
    if (normalizedEnd <= normalizedStart) return [];

    // Geometry batching: a leaf run tspan renders on a single visual row, and
    // Chrome returns a uniform y/height for every glyph in it, so a contiguous
    // sub-range's union box is exactly bounded by its first and last glyph
    // extents. Measure those two (plus one getScreenCTM) per span instead of
    // calling getExtentOfChar/getScreenCTM per character — O(spans) DOM-geometry
    // calls instead of O(chars), which matters on dense slides where highlights
    // and selections cover long runs and re-measure on every render.
    const rows: SvgInlineSelectionBox[] = [];
    for (const entry of entries) {
      const spanStart = entry.start;
      const spanEnd = entry.start + entry.count;
      const from = Math.max(normalizedStart, spanStart);
      const to = Math.min(normalizedEnd, spanEnd);
      if (to <= from) continue;

      const elementMatrix = entry.span.getScreenCTM();
      if (!elementMatrix) continue;

      for (const box of this.measureSpanRangeBoxes(entry.span, from - spanStart, to - spanStart, elementMatrix, rootInverse)) {
        this.mergeSelectionRowBox(rows, box);
      }
    }

    const padding = this.getSvgInlineSelectionPadding(element);
    return rows.map((box) => ({
      x: box.x - padding,
      y: box.y - padding * 0.5,
      width: box.width + padding * 2,
      height: box.height + padding
    }));
  }

  /**
   * Row boxes covering a leaf span's `[localStart, localEnd)` string sub-range.
   *
   * Fast path (the overwhelming common case): the first and last glyph extents
   * share a row, so the union is exactly `[minX..maxRight] × row height` — two
   * `getExtentOfChar` calls regardless of range length. Falls back to a per-glyph
   * sweep only if a span unexpectedly spans multiple rows (or an endpoint can't
   * be measured), preserving the original behavior in that rare case.
   */
  private measureSpanRangeBoxes(
    span: SVGTextContentElement,
    localStart: number,
    localEnd: number,
    elementMatrix: DOMMatrix,
    rootInverse: DOMMatrix
  ): SvgInlineSelectionBox[] {
    // `localStart`/`localEnd` are string offsets; getExtentOfChar wants glyph
    // indices, so clamp into the rendered glyph range first.
    const firstGlyph = this.inlineGeometry.clampGlyphIndex(span, localStart);
    const lastGlyph = this.inlineGeometry.clampGlyphIndex(span, localEnd - 1);
    if (lastGlyph < firstGlyph) return [];

    const measure = (glyph: number): SvgInlineSelectionBox | null => {
      try {
        const box = this.transformSvgRectToSvgRoot(span.getExtentOfChar(glyph), elementMatrix, rootInverse);
        return box && box.width >= 0 && box.height > 0 ? box : null;
      } catch {
        return null;
      }
    };

    const firstBox = measure(firstGlyph);
    const lastBox = lastGlyph === firstGlyph ? firstBox : measure(lastGlyph);
    if (firstBox && lastBox) {
      const sameRow = Math.abs(
        (firstBox.y + firstBox.height / 2) - (lastBox.y + lastBox.height / 2)
      ) < Math.max(2, firstBox.height * 0.55);
      if (sameRow) {
        const x = Math.min(firstBox.x, lastBox.x);
        const y = Math.min(firstBox.y, lastBox.y);
        const right = Math.max(firstBox.x + firstBox.width, lastBox.x + lastBox.width);
        const bottom = Math.max(firstBox.y + firstBox.height, lastBox.y + lastBox.height);
        return [{ x, y, width: right - x, height: bottom - y }];
      }
    }

    // Multi-row span (or unmeasurable endpoint): fall back to a per-glyph sweep.
    const rows: SvgInlineSelectionBox[] = [];
    for (let glyph = firstGlyph; glyph <= lastGlyph; glyph++) {
      const box = measure(glyph);
      if (box) this.mergeSelectionRowBox(rows, box);
    }
    return rows;
  }

  /** Union `box` into the row whose vertical center it shares, else start a new row. */
  private mergeSelectionRowBox(rows: SvgInlineSelectionBox[], box: SvgInlineSelectionBox): void {
    const centerY = box.y + box.height / 2;
    const row = rows.find((candidate) => (
      Math.abs(centerY - (candidate.y + candidate.height / 2)) < Math.max(2, box.height * 0.55)
    ));
    if (!row) {
      rows.push({ ...box });
      return;
    }
    const left = Math.min(row.x, box.x);
    const top = Math.min(row.y, box.y);
    const right = Math.max(row.x + row.width, box.x + box.width);
    const bottom = Math.max(row.y + row.height, box.y + box.height);
    row.x = left;
    row.y = top;
    row.width = right - left;
    row.height = bottom - top;
  }

  /**
   * When the pointer is on or past the last glyph of a wrapped visual line, snap
   * the run-local offset to the line end so selections can reach trailing text
   * like "app." (geometry hit-testing often stops one character short).
   */
  private snapWrappedRunLocalToLineEnd(
    container: SVGTextElement | SVGTSpanElement,
    runLocal: number,
    runTotal: number,
    localClientX: number
  ): number {
    return this.inlineGeometry.snapWrappedRunLocalToLineEnd(container, runLocal, runTotal, localClientX);
  }

  private getInlineTextOffsetAtClientPoint(
    element: SVGTextElement | SVGTSpanElement,
    editor: HTMLTextAreaElement,
    clientX: number,
    clientY: number | undefined,
    box: { left: number; top: number; width: number; height: number }
  ): number {
    const text = editor.value;
    if (text.length === 0) return 0;

    const target = this.activeShapeTextTarget;
    if (target && this.needsFlatParagraphMapping(target.shapeIndex, target.paragraphIndex, element)) {
      const runContainers = this.getRunLineContainers(target.shapeIndex, target.paragraphIndex);
      const paneRect = this.canvasPane?.getBoundingClientRect();
      const localClientX = paneRect
        ? clientX - paneRect.left + (this.canvasPane?.scrollLeft ?? 0)
        : clientX;
      const localClientY = paneRect && clientY !== undefined
        ? clientY - paneRect.top + (this.canvasPane?.scrollTop ?? 0)
        : null;

      let flatOffset = 0;
      let fallbackOffset = 0;
      let fallbackDistance = Number.POSITIVE_INFINITY;

      for (const container of runContainers) {
        const containerBox = this.getElementBox(container);
        const geometryTotal = this.getLeafCharInfo(container).total;
        const runTotal = this.getRunCharInfo(container).total;

        if (containerBox) {
          const geometryLocal = Math.max(
            0,
            Math.min(
              geometryTotal,
              this.getInlineTextOffsetAtClientPointForElement(container, clientX, clientY, containerBox)
            )
          );
          let runLocal = this.geometryIndexToRunOffset(container, geometryLocal);
          runLocal = this.snapWrappedRunLocalToLineEnd(container, runLocal, runTotal, localClientX);

          if (
            localClientY !== null
            && localClientY >= containerBox.top
            && localClientY <= containerBox.top + containerBox.height
          ) {
            return Math.max(0, Math.min(text.length, flatOffset + runLocal));
          }

          const centerY = containerBox.top + containerBox.height / 2;
          const distance = localClientY !== null
            ? Math.abs(localClientY - centerY)
            : Math.abs(localClientX - (containerBox.left + containerBox.width / 2));
          if (distance < fallbackDistance) {
            fallbackDistance = distance;
            fallbackOffset = flatOffset + runLocal;
          }
        }

        flatOffset += runTotal;
      }

      if (fallbackDistance < Number.POSITIVE_INFINITY) {
        return Math.max(0, Math.min(text.length, fallbackOffset));
      }
    }

    const paneRect = this.canvasPane?.getBoundingClientRect();
    const localClientX = paneRect
      ? clientX - paneRect.left + (this.canvasPane?.scrollLeft ?? 0)
      : clientX;
    const localClientY = paneRect && clientY !== undefined
      ? clientY - paneRect.top + (this.canvasPane?.scrollTop ?? 0)
      : box.top + box.height / 2;

    const geometryTotal = this.getLeafCharInfo(element).total;
    const geometryOffset = this.getInlineTextOffsetFromSvgGeometry(
      element,
      localClientX,
      localClientY,
      geometryTotal
    );
    if (geometryOffset !== null) {
      return this.geometryIndexToRunOffset(element, geometryOffset);
    }

    return this.getMeasuredInlineTextOffset(editor, localClientX, box);
  }

  private getInlineTextOffsetAtClientPointForElement(
    element: SVGTextElement | SVGTSpanElement,
    clientX: number,
    clientY: number | undefined,
    box: { left: number; top: number; width: number; height: number }
  ): number {
    return this.inlineGeometry.getInlineTextOffsetAtClientPointForElement(element, clientX, clientY, box);
  }

  private getInlineTextOffsetFromSvgGeometry(
    element: SVGTextElement | SVGTSpanElement,
    localClientX: number,
    localClientY: number,
    textLength: number
  ): number | null {
    return this.inlineGeometry.getInlineTextOffsetFromSvgGeometry(element, localClientX, localClientY, textLength);
  }

  private getInlineCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    editor: HTMLTextAreaElement,
    offset: number,
    box: { left: number; top: number; width: number; height: number }
  ): { left: number; top: number; height: number } {
    const target = this.activeShapeTextTarget;
    const resolved = target
      ? this.resolveParagraphFlatOffset(target.shapeIndex, target.paragraphIndex, element, offset)
      : { element, localOffset: offset };
    const geometryElement = resolved.element;
    const localOffset = resolved.localOffset;
    const geometryBox = this.getElementBox(geometryElement) ?? box;

    const fallbackHeight = this.getInlineCaretHeight(geometryElement, geometryBox);
    const svgGeometry = this.getSvgTextCaretGeometry(geometryElement, localOffset, fallbackHeight);
    if (svgGeometry) {
      return { left: svgGeometry.left, top: svgGeometry.top, height: svgGeometry.height };
    }

    const row = this.activeInlineCaretRow ?? this.getDefaultInlineCaretRow(geometryElement, geometryBox, fallbackHeight);

    const text = editor.value;
    if (text.length === 0) return { left: box.left, ...row };

    const style = window.getComputedStyle(editor);
    const canvas = activeDocument.createEl('canvas');
    const context = canvas.getContext('2d');
    if (!context) {
      return { left: box.left + box.width * (offset / text.length), ...row };
    }

    context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    const fullWidth = context.measureText(text).width;
    if (fullWidth <= 0) return { left: geometryBox.left, ...row };

    return {
      left: geometryBox.left + context.measureText(text.slice(0, offset)).width * (geometryBox.width / fullWidth),
      ...row
    };
  }

  private getInlineCaretRowFromClientY(
    element: SVGTextElement | SVGTSpanElement,
    clientY: number | undefined,
    box: { left: number; top: number; width: number; height: number }
  ): InlineCaretRow {
    const height = this.getInlineCaretHeight(element, box);
    const paneRect = this.canvasPane?.getBoundingClientRect();
    if (clientY === undefined || !paneRect) {
      return this.getDefaultInlineCaretRow(element, box, height);
    }

    const localY = clientY - paneRect.top + (this.canvasPane?.scrollTop ?? 0);
    if (box.height <= height * 1.8) {
      return this.getDefaultInlineCaretRow(element, box, height);
    }

    const centerRatio = box.height > 0 ? (localY - box.top) / box.height : 0.5;
    return this.getInlineCaretRowFromRatio(element, box, centerRatio, height);
  }

  private getInlineCaretHeight(element: SVGTextElement | SVGTSpanElement, box: { width: number; height: number }): number {
    const lineCount = this.estimateInlineTextRowCount(element);
    const lineBoxHeight = box.height / Math.max(1, lineCount);
    const screenFontSize = Math.min(this.getScreenFontSize(element), lineBoxHeight);
    const baseHeight = Math.min(lineBoxHeight, screenFontSize || lineBoxHeight);
    return Math.max(6, baseHeight * 0.88);
  }

  private getSvgInlineCaretStrokeWidth(element: SVGTextElement | SVGTSpanElement): number {
    const style = window.getComputedStyle(element);
    const fontSize = Number.parseFloat(style.fontSize) || 16;
    return Math.max(1.25, Math.min(4, fontSize / 14));
  }

  private getSvgInlineSelectionPadding(element: SVGTextElement | SVGTSpanElement): number {
    const style = window.getComputedStyle(element);
    const fontSize = Number.parseFloat(style.fontSize) || 16;
    return Math.max(0.75, Math.min(3, fontSize / 18));
  }

  private refreshActiveInlineCaretRow(
    element: SVGTextElement | SVGTSpanElement,
    box: { left: number; top: number; width: number; height: number }
  ): void {
    if (!this.activeInlineCaretRow) return;

    this.activeInlineCaretRow = this.getInlineCaretRowFromRatio(
      element,
      box,
      this.activeInlineCaretRow.centerRatio
    );
  }

  private estimateInlineTextRowCount(element: SVGTextElement | SVGTSpanElement): number {
    const text = element.textContent || '';
    if (!text) return 1;

    const paneRect = this.canvasPane?.getBoundingClientRect();
    if (!paneRect) return 1;

    try {
      const rows: number[] = [];
      for (const { span } of this.getLeafCharInfo(element).entries) {
        const matrix = span.getScreenCTM();
        if (!matrix) continue;
        // Probe per-glyph (not per-string-char): getStartPositionOfChar throws
        // past the rendered glyph range when shaping collapses chars to glyphs.
        const glyphCount = this.inlineGeometry.getGlyphCount(span);
        for (let index = 0; index < glyphCount; index++) {
          const position = span.getStartPositionOfChar(index);
          const point = new DOMPoint(position.x, position.y).matrixTransform(matrix);
          const localY = point.y - paneRect.top + (this.canvasPane?.scrollTop ?? 0);
          if (!rows.some((row) => Math.abs(row - localY) < 4)) {
            rows.push(localY);
          }
        }
      }
      return Math.max(1, rows.length);
    } catch {
      return 1;
    }
  }

  private getDefaultInlineCaretRow(
    _element: SVGTextElement | SVGTSpanElement,
    box: { left: number; top: number; width: number; height: number },
    height: number
  ): InlineCaretRow {
    return this.getInlineCaretRowFromRatio(_element, box, 0.5, height);
  }

  private getInlineCaretRowFromRatio(
    element: SVGTextElement | SVGTSpanElement,
    box: { left: number; top: number; width: number; height: number },
    centerRatio: number,
    height = this.getInlineCaretHeight(element, box)
  ): InlineCaretRow {
    const ratio = Math.max(0, Math.min(1, centerRatio));
    const minCenter = box.top + height / 2;
    const maxCenter = Math.max(minCenter, box.top + box.height - height / 2);
    const center = Math.max(minCenter, Math.min(maxCenter, box.top + box.height * ratio));
    return {
      top: center - height / 2,
      height,
      centerRatio: ratio
    };
  }

  private getMeasuredInlineTextOffset(
    editor: HTMLTextAreaElement,
    localClientX: number,
    box: { left: number; top: number; width: number; height: number }
  ): number {
    const text = editor.value;
    const clickOffset = Math.max(0, Math.min(box.width, localClientX - box.left));
    const style = window.getComputedStyle(editor);
    const canvas = activeDocument.createEl('canvas');
    const context = canvas.getContext('2d');
    if (!context) {
      return Math.round(text.length * (clickOffset / box.width));
    }

    context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    const measuredWidth = context.measureText(text).width;
    const widthScale = measuredWidth > 0 ? box.width / measuredWidth : 1;
    let previousWidth = 0;
    for (let offset = 1; offset <= text.length; offset++) {
      const width = context.measureText(text.slice(0, offset)).width * widthScale;
      if (clickOffset <= (previousWidth + width) / 2) {
        return offset - 1;
      }
      previousWidth = width;
    }

    return text.length;
  }

  private resetInlineEditorScroll(editor: HTMLTextAreaElement): void {
    editor.scrollLeft = 0;
    editor.scrollTop = 0;
  }

  private getSvgTextCaretLeft(element: SVGTextElement | SVGTSpanElement, offset: number): number | null {
    return this.getSvgTextCaretGeometry(element, offset)?.left ?? null;
  }

  private getLeafTextSpans(element: SVGTextElement | SVGTSpanElement): SVGTextContentElement[] {
    return this.inlineGeometry.getLeafTextSpans(element);
  }

  private getParagraphLeafText(element: SVGTextElement | SVGTSpanElement): string {
    return this.getLeafCharInfo(element).entries
      .map((entry) => entry.span.textContent || '')
      .join('');
  }

  private isRunTextSpan(span: SVGTextContentElement): boolean {
    return this.inlineGeometry.isRunTextSpan(span);
  }

  /** Character counts for OOXML runs only — matches the inline editor / engine offsets. */
  private getRunCharInfo(
    element: SVGTextElement | SVGTSpanElement
  ): { entries: { span: SVGTextContentElement; count: number; start: number }[]; total: number } {
    return this.inlineGeometry.getRunCharInfo(element);
  }

  /** Map a geometry (all-leaf) character index to a run-only offset within `element`. */
  private geometryIndexToRunOffset(element: SVGTextElement | SVGTSpanElement, geometryIndex: number): number {
    return this.inlineGeometry.geometryIndexToRunOffset(element, geometryIndex);
  }

  /** Map a run-only offset to a geometry (all-leaf) character index within `element`. */
  private runOffsetToGeometryIndex(element: SVGTextElement | SVGTSpanElement, runOffset: number): number {
    return this.inlineGeometry.runOffsetToGeometryIndex(element, runOffset);
  }

  private getLeafCharInfo(
    element: SVGTextElement | SVGTSpanElement
  ): { entries: { span: SVGTextContentElement; count: number; start: number }[]; total: number } {
    return this.inlineGeometry.getLeafCharInfo(element);
  }

  private getSvgTextCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    offset: number,
    preferredHeight = this.getScreenFontSize(element) * 1.08
  ): { left: number; top: number; height: number } | null {
    return this.inlineGeometry.getSvgTextCaretGeometry(element, offset, preferredHeight);
  }

  private localPointToSvgRoot(left: number, top: number): DOMPoint | null {
    if (!this.canvasPane || !this.svgEl) return null;

    const matrix = this.svgEl.getScreenCTM();
    const paneRect = this.canvasPane.getBoundingClientRect();
    if (!matrix || !paneRect) return null;

    const screenPoint = new DOMPoint(
      paneRect.left + left - this.canvasPane.scrollLeft,
      paneRect.top + top - this.canvasPane.scrollTop
    );

    try {
      return screenPoint.matrixTransform(matrix.inverse());
    } catch {
      return null;
    }
  }

  private transformSvgRectToSvgRoot(
    rect: SvgRectLike,
    elementMatrix: DOMMatrix,
    rootInverse: DOMMatrix
  ): SvgInlineSelectionBox | null {
    const points = [
      new DOMPoint(rect.x, rect.y),
      new DOMPoint(rect.x + rect.width, rect.y),
      new DOMPoint(rect.x, rect.y + rect.height),
      new DOMPoint(rect.x + rect.width, rect.y + rect.height)
    ].map((point) => point.matrixTransform(elementMatrix).matrixTransform(rootInverse));
    if (points.some((point) => !Number.isFinite(point.x) || !Number.isFinite(point.y))) {
      return null;
    }

    const xs = points.map((point) => point.x);
    const ys = points.map((point) => point.y);
    const x = Math.min(...xs);
    const y = Math.min(...ys);
    return {
      x,
      y,
      width: Math.max(...xs) - x,
      height: Math.max(...ys) - y
    };
  }

  private formatSvgNumber(value: number): string {
    return `${Math.round(value * 1000) / 1000}`;
  }

  private getSelectedShapeElement(): SVGGElement | null {
    if (!this.svgEl || this.selectedShapeIndex === null) return null;
    const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${this.selectedShapeIndex}"]`);
    return isSVGGElement(shape) ? shape : null;
  }

  private getTextEditTargetFromSelectedShape(): ShapeTextEditTarget | null {
    const shape = this.getSelectedShapeElement();
    if (!shape) return null;
    const textElement = shape.querySelector('text');
    const target = this.getTextEditTarget(textElement);
    return target?.kind === 'shape-paragraph' ? target : null;
  }

  private getPrimaryStyleRunTarget(target: ShapeTextEditTarget): ShapeTextEditTarget {
    const run = target.runElements[0];
    if (!run) return target;
    return { ...target, element: run };
  }

  private syncShapeParagraphPreview(
    target: TextEditTarget,
    text: string,
    options: { replaceTextFrame?: boolean } = {},
  ): boolean {
    if (target.kind === 'shape-paragraph') {
      // A formatting action or incremental SVG replacement can leave the edit
      // target pointing at detached tspans. The textarea is deliberately
      // transparent, so always rebind before writing: otherwise its caret can
      // advance while the on-slide text stays unchanged until commit re-renders.
      this.refreshShapeParagraphPreviewTarget(target);

      if (!this.reflowShapeParagraphPreview(target, text)) {
        const firstRun = target.runElements[0];
        if (firstRun) {
          const previousRunTexts = target.runElements.map((run) => run.textContent || '');
          const previewRunTexts = redistributeTextAcrossVisualRuns(previousRunTexts, text);
          target.runElements = target.runElements.map((run, index) =>
            this.replaceLiveParagraphRunText(run, previewRunTexts[index] ?? '')
          );
          this.activeTextStyleTarget = this.getPrimaryStyleRunTarget(target);
        } else if (isSVGTextElement(target.element)) {
          target.element.textContent = text;
        }
      }

      this.reconcileShapeParagraphPreview(target, text);
      return options.replaceTextFrame
        ? this.replaceLiveShapeTextFrame(target)
        : false;
    }

    target.element.textContent = text;
    return false;
  }

  /**
   * Chromium occasionally retains stale SVG glyphs after a transparent-textarea
   * deletion clears a selection or the final character. Replacing the owning
   * <text> node is the narrow repaint boundary; paragraph/run identities are
   * immediately rebound below, so the editor and caret continue using live DOM.
   */
  private replaceLiveShapeTextFrame(target: ShapeTextEditTarget): boolean {
    const targetConnected = target.element.isConnected;
    const directTextElement = target.element.closest('text');
    let textElement: SVGTextElement | null = (
      isSVGTextElement(directTextElement) && directTextElement.isConnected
    ) ? directTextElement : null;
    let ownerSource: 'target' | 'shape' | null = textElement ? 'target' : null;

    // Preview reflow can replace tspans before this method reaches the frame.
    // Recover via the live shape instead of silently returning false with a
    // detached edit target—the exact split that leaves the caret moving while
    // old SVG glyphs remain painted.
    if (!textElement) {
      const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${target.shapeIndex}"]`);
      const paragraph = shape?.querySelector(
        `tspan[data-ooxml-para-idx="${target.paragraphIndex}"]`
      ) ?? null;
      const shapeTextElement = paragraph?.closest('text') ?? shape?.querySelector('text') ?? null;
      if (isSVGTextElement(shapeTextElement) && shapeTextElement.isConnected) {
        textElement = shapeTextElement;
        ownerSource = 'shape';
      }
    }

    if (!textElement) {
      debugLog('text-edit', 'Skipped inline SVG text-frame refresh without a connected owner', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
        targetConnected,
      });
      return false;
    }

    const replacement = textElement.cloneNode(true) as SVGTextElement;
    textElement.replaceWith(replacement);
    this.refreshShapeParagraphPreviewTarget(target);
    debugLog('text-edit', 'Refreshed live SVG text frame after inline edit', {
      slide: this.currentSlide,
      shapeIndex: target.shapeIndex,
      paragraphIndex: target.paragraphIndex,
      ownerSource,
      targetConnected,
      reboundTargetConnected: target.element.isConnected,
    });
    return true;
  }

  /**
   * Refresh an active paragraph's SVG references without changing its OOXML
   * identity. Incremental shape renders replace only the visual nodes, so the
   * editor must not keep writing through an old tspan reference.
   */
  private refreshShapeParagraphPreviewTarget(target: ShapeTextEditTarget): void {
    const lineContainers = this.getRunLineContainers(target.shapeIndex, target.paragraphIndex);
    const runElements = lineContainers.flatMap((line) => this.collectParagraphRuns(line));
    const nextElement = lineContainers[0];
    if (!nextElement || runElements.length === 0) return;

    const elementChanged = target.element !== nextElement;
    const runsChanged = target.runElements.length !== runElements.length
      || target.runElements.some((run, index) => run !== runElements[index]);
    if (!elementChanged && !runsChanged) return;

    const previousElement = target.element;
    target.element = nextElement;
    target.runElements = runElements;

    if (this.activeShapeTextTarget === target) {
      if (this.activeEditorTarget === previousElement) {
        previousElement.classList.remove('native-powerpoint-text-editing');
      }
      this.activeEditorTarget = nextElement;
      this.activeEditorTarget.classList.add('native-powerpoint-text-editing');
      this.activeTextStyleTarget = this.getPrimaryStyleRunTarget(target);
    }
  }

  /**
   * Keep the visible SVG and transparent textarea in lockstep even if a custom
   * preview reflow could not preserve every run. Normal keystrokes stay quiet;
   * this logs only a detected mismatch and its repair result.
   */
  private reconcileShapeParagraphPreview(target: ShapeTextEditTarget, text: string): void {
    const lineContainers = this.getRunLineContainers(target.shapeIndex, target.paragraphIndex);
    const visibleText = this.getParagraphPlainText(lineContainers);
    if (visibleText === text) return;

    const runElements = lineContainers.flatMap((line) => this.collectParagraphRuns(line));
    if (runElements.length > 0) {
      const previewRunTexts = redistributeTextAcrossVisualRuns(
        runElements.map((run) => run.textContent || ''),
        text
      );
      target.runElements = runElements.map((run, index) =>
        this.replaceLiveParagraphRunText(run, previewRunTexts[index] ?? '')
      );
      this.activeTextStyleTarget = this.getPrimaryStyleRunTarget(target);
    } else if (isSVGTextElement(target.element)) {
      target.element.textContent = text;
    }

    const reconciledText = this.getParagraphPlainText(
      this.getRunLineContainers(target.shapeIndex, target.paragraphIndex)
    );

    debugLog('text-edit', 'Live inline text preview mismatch reconciled', {
      slide: this.currentSlide,
      shapeIndex: target.shapeIndex,
      paragraphIndex: target.paragraphIndex,
      editorTextLength: text.length,
      visibleTextLength: visibleText.length,
      reconciledTextLength: reconciledText.length,
      resolved: reconciledText === text,
      runCount: runElements.length,
      lineCount: lineContainers.length,
      targetConnected: target.element.isConnected,
    });
  }

  /**
   * Replace changed connected run nodes so Chromium invalidates the SVG glyph
   * cache immediately. Updating `textContent` alone can leave the old glyphs
   * painted until an unrelated focus change forces a later redraw.
   */
  private replaceLiveParagraphRunText(run: SVGTSpanElement, text: string): SVGTSpanElement {
    if (run.textContent === text) return run;
    if (!run.isConnected || !run.parentNode) {
      run.textContent = text;
      return run;
    }

    const replacement = run.cloneNode(true) as SVGTSpanElement;
    replacement.textContent = text;
    run.replaceWith(replacement);
    return replacement;
  }

  /**
   * Rebuild the temporary SVG line tspans while typing so text wraps before the
   * OOXML commit. The final engine render still owns the authoritative layout;
   * this only keeps the on-canvas edit preview inside its text frame.
   */
  private reflowShapeParagraphPreview(target: ShapeTextEditTarget, text: string): boolean {
    // Shift+Enter is an explicit DrawingML soft break. The committed renderer
    // already handles it correctly, while this character-based preview only
    // creates automatic soft wraps.
    if (text.includes('\n')) return false;

    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${target.shapeIndex}"]`);
    const textElement = target.element.closest('text');
    const frame = shape?.querySelector(':scope > rect') ?? null;
    const lineContainers = this.getRunLineContainers(target.shapeIndex, target.paragraphIndex);
    const firstLine = lineContainers[0];
    const firstRun = target.runElements[0];
    if (!shape || !isSVGTextElement(textElement) || !firstLine || !firstRun || !frame) return false;

    const frameBox = this.getElementBox(frame);
    const firstLineBox = this.getElementBox(firstLine);
    const baseY = Number(firstLine.getAttribute('y'));
    const lineStep = this.getLivePreviewLineStep(lineContainers, firstRun);
    if (!frameBox || !firstLineBox || !Number.isFinite(baseY) || lineStep === null) return false;

    const inset = Math.max(0, firstLineBox.left - frameBox.left);
    const maxWidth = frameBox.width - inset * 2;
    if (!Number.isFinite(maxWidth) || maxWidth < 4) return false;

    const measure = this.createInlinePreviewTextMeasurer(firstRun);
    if (!measure) return false;
    const lines = wrapTextForPreview(text, maxWidth, measure);
    if (lines.join('') !== text) return false;

    // A single original line needs no DOM reconstruction until its first wrap.
    // Once there are several visual lines, rebuild on every edit so a word can
    // move back up when text is deleted as well as down when it is added.
    if (lines.length === 1 && lineContainers.length === 1) return false;

    const parent = firstLine.parentNode;
    if (!parent) return false;

    const previousRunTexts = target.runElements.map((run) => run.textContent || '');
    const nextRunTexts = redistributeTextAcrossVisualRuns(previousRunTexts, text);
    const boundaries = [0];
    for (const runText of nextRunTexts) {
      boundaries.push((boundaries[boundaries.length - 1] ?? 0) + runText.length);
    }

    let offset = 0;
    const nextLines = lines.map((lineText, lineIndex) => {
      const line = firstLine.cloneNode(false) as SVGTSpanElement;
      line.setAttribute('y', `${baseY + lineStep * lineIndex}`);
      line.removeAttribute('dy');
      const lineStart = offset;
      const lineEnd = lineStart + lineText.length;
      offset = lineEnd;

      for (let runIndex = 0; runIndex < target.runElements.length; runIndex++) {
        const runStart = boundaries[runIndex] ?? lineEnd;
        const runEnd = boundaries[runIndex + 1] ?? runStart;
        const start = Math.max(lineStart, runStart);
        const end = Math.min(lineEnd, runEnd);
        if (end <= start) continue;
        const source = target.runElements[runIndex];
        if (!source) continue;
        const run = source.cloneNode(true) as SVGTSpanElement;
        run.textContent = text.slice(start, end);
        line.appendChild(run);
      }

      // Preserve an editable caret target for an empty final visual line.
      if (line.children.length === 0) {
        const emptyRun = firstRun.cloneNode(true) as SVGTSpanElement;
        emptyRun.textContent = '';
        line.appendChild(emptyRun);
      }
      return line;
    });

    const anchor = lineContainers[lineContainers.length - 1]?.nextSibling ?? null;
    // Insert each SVG line directly. Obsidian's createFragment helper can be
    // backed by the owning XML document in a pop-out window, which rejects a
    // second root node before this code reaches the replacement step.
    for (const line of nextLines) parent.insertBefore(line, anchor);
    for (const line of lineContainers) line.remove();

    const nextTarget = nextLines[0];
    if (!nextTarget) return false;
    const previousEditorTarget = this.activeEditorTarget;
    target.element = nextTarget;
    target.runElements = nextLines.flatMap((line) => this.collectParagraphRuns(line));
    if (this.activeShapeTextTarget === target) {
      if (previousEditorTarget && previousEditorTarget !== nextTarget) {
        previousEditorTarget.classList.remove('native-powerpoint-text-editing');
      }
      this.activeEditorTarget = nextTarget;
      this.activeEditorTarget.classList.add('native-powerpoint-text-editing');
      this.activeTextStyleTarget = this.getPrimaryStyleRunTarget(target);
    }

    const previousLineCount = this.activeInlinePreviewLineCount ?? lineContainers.length;
    this.activeInlinePreviewLineCount = nextLines.length;
    if (previousLineCount !== nextLines.length) {
      debugLog('text-edit', 'Live inline text preview reflowed', {
        slide: this.currentSlide,
        shapeIndex: target.shapeIndex,
        paragraphIndex: target.paragraphIndex,
        previousLineCount,
        nextLineCount: nextLines.length,
        textLength: text.length,
        frameWidth: Math.round(maxWidth)
      });
    }
    return true;
  }

  /** Use the first run's effective on-screen font to decide where a word wraps. */
  private createInlinePreviewTextMeasurer(run: SVGTSpanElement): ((value: string) => number) | null {
    // Prefer Window.createEl (detached canvas). Document.createEl on an SVG/XML
    // pop-out can append to a Document that already has a root and throw
    // "Only one element on document allowed", aborting preview sync.
    const scopedWindow = (run.ownerDocument?.defaultView ?? window) as Window & {
      createEl?: (tag: 'canvas') => HTMLCanvasElement;
    };
    if (typeof scopedWindow.createEl !== 'function') return null;
    const canvas = scopedWindow.createEl('canvas');
    const context = canvas.getContext('2d');
    if (!context) return null;
    const style = scopedWindow.getComputedStyle(run);
    context.font = `${style.fontStyle} ${style.fontWeight} ${this.getScreenFontSize(run)}px ${style.fontFamily}`;
    return (value: string) => context.measureText(value).width;
  }

  /** SVG line coordinates are unscaled, so derive their step in SVG units. */
  private getLivePreviewLineStep(
    lineContainers: readonly SVGTSpanElement[],
    firstRun: SVGTSpanElement
  ): number | null {
    const firstY = Number(lineContainers[0]?.getAttribute('y'));
    const secondY = Number(lineContainers[1]?.getAttribute('y'));
    if (Number.isFinite(firstY) && Number.isFinite(secondY) && secondY > firstY) {
      return secondY - firstY;
    }

    const fontSize = Number(firstRun.getAttribute('font-size'));
    return Number.isFinite(fontSize) && fontSize > 0 ? fontSize * 1.2 : null;
  }

  private collectParagraphRuns(paraContainer: Element): SVGTSpanElement[] {
    const runs = paraContainer.matches('tspan[data-ooxml-para-idx]')
      ? Array.from(paraContainer.querySelectorAll(':scope > tspan[data-ooxml-run-idx]'))
      : Array.from(paraContainer.querySelectorAll('tspan[data-ooxml-run-idx]'));
    return runs.filter(isSVGTSpanElement);
  }

  private collectParagraphLineContainers(textEl: SVGTextElement, paragraphIndex: number): SVGTSpanElement[] {
    const direct = Array.from(textEl.children).filter(
      (child): child is SVGTSpanElement =>
        isSVGTSpanElement(child) && child.getAttribute('data-ooxml-para-idx') === String(paragraphIndex)
    );
    if (direct.length > 0) return direct;

    const nested = textEl.querySelector(`tspan[data-ooxml-para-idx="${paragraphIndex}"]`);
    return nested && isSVGTSpanElement(nested) ? [nested] : [];
  }

  private getParagraphLineContainers(shapeIndex: number, paragraphIndex: number): SVGTSpanElement[] {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!shape) return [];
    const textEl = shape.querySelector('text');
    if (!isSVGTextElement(textEl)) return [];
    return this.collectParagraphLineContainers(textEl, paragraphIndex);
  }

  /**
   * Visual line containers that actually hold OOXML runs. Bullet/number markers
   * are rendered as their own `data-ooxml-para-idx` tspan with no
   * `data-ooxml-run-idx` child, so they must be excluded: the inline editor's
   * text and the OOXML run offsets are run-only and do not count the bullet
   * glyph. Counting bullet glyphs here shifts every selection offset by the
   * bullet length, which underlines/highlights the wrong characters.
   */
  private getRunLineContainers(shapeIndex: number, paragraphIndex: number): SVGTSpanElement[] {
    return this.getParagraphLineContainers(shapeIndex, paragraphIndex)
      .filter((container) => this.collectParagraphRuns(container).length > 0);
  }

  /**
   * True when paragraph offsets need the run-line-container mapping rather than
   * the editor's plain `element`. That happens when the paragraph soft-wraps
   * across multiple run lines, or when a bullet/leading container makes the
   * editor's geometry element differ from the first run container.
   */
  private needsFlatParagraphMapping(
    shapeIndex: number,
    paragraphIndex: number,
    element: SVGTextElement | SVGTSpanElement
  ): boolean {
    const runContainers = this.getRunLineContainers(shapeIndex, paragraphIndex);
    if (runContainers.length === 0) return false;
    if (runContainers.length > 1) return true;
    const container = runContainers[0];
    if (!container) return false;
    if (container !== element) return true;
    return this.getRunCharInfo(container).total !== this.getLeafCharInfo(container).total;
  }

  private getParagraphPlainText(lineContainers: SVGTSpanElement[]): string {
    return lineContainers
      .map((container) => this.collectParagraphRuns(container).map((run) => run.textContent || '').join(''))
      .join('');
  }

  private paragraphEditorTextFromDom(
    shapeIndex: number,
    paragraphIndex: number,
    editorText: string
  ): string {
    const runContainers = this.getRunLineContainers(shapeIndex, paragraphIndex);
    if (runContainers.length === 0) return editorText;

    const domFlat = this.getParagraphPlainText(runContainers);
    if (editorText === domFlat) return editorText;

    // Strip any legacy soft-wrap newlines that an older editor state may have
    // inserted between visual lines of the same OOXML paragraph.
    const parts = editorText.split('\n');
    if (parts.length === runContainers.length && parts.join('') === domFlat) {
      return domFlat;
    }

    return editorText;
  }

  private resolveParagraphFlatOffset(
    shapeIndex: number,
    paragraphIndex: number,
    fallbackElement: SVGTextElement | SVGTSpanElement,
    flatOffset: number
  ): { element: SVGTextElement | SVGTSpanElement; localOffset: number } {
    const runContainers = this.getRunLineContainers(shapeIndex, paragraphIndex);
    if (runContainers.length === 0) {
      return { element: fallbackElement, localOffset: flatOffset };
    }

    const charCounts = runContainers.map((container) => this.getRunCharInfo(container).total);
    const { lineIndex, localOffset: runLocal } = mapFlatOffsetToRunLine(charCounts, flatOffset);
    const container = runContainers[lineIndex] ?? fallbackElement;
    const geometryLocal = this.runOffsetToGeometryIndex(container, runLocal);
    return { element: container, localOffset: geometryLocal };
  }

  private renderFlatParagraphSelection(
    shapeIndex: number,
    paragraphIndex: number,
    flatStart: number,
    flatEnd: number
  ): void {
    const runContainers = this.getRunLineContainers(shapeIndex, paragraphIndex);
    if (runContainers.length === 0) return;

    const charCounts = runContainers.map((container) => this.getRunCharInfo(container).total);
    for (const segment of mapFlatRangeToRunLineSegments(charCounts, flatStart, flatEnd)) {
      const container = runContainers[segment.lineIndex];
      if (!container) continue;
      const geometryStart = this.runOffsetToGeometryIndex(container, segment.localStart);
      const geometryEnd = this.runOffsetToGeometryIndex(container, segment.localEnd);
      this.renderInlineSelectionRects(container, geometryStart, geometryEnd);
    }
  }

  private getTextEditTarget(element: Element | null): TextEditTarget | null {
    const textEl = element?.closest('text');
    if (!isSVGTextElement(textEl) || textEl.closest(GENERATED_GRID_SELECTOR)) return null;

    const shape = textEl.closest('g[data-ooxml-shape-idx]');
    const shapeIndex = getShapeIndex(shape);
    if (shapeIndex === null) return null;

    const clickedRun = element?.closest('tspan[data-ooxml-run-idx]');
    let paraContainer: Element | null = clickedRun?.closest('tspan[data-ooxml-para-idx]') ?? null;
    if (!paraContainer) {
      paraContainer = textEl.querySelector('tspan[data-ooxml-para-idx]') ?? textEl;
    }

    const seedRuns = this.collectParagraphRuns(paraContainer);
    if (seedRuns.length === 0) {
      const text = textEl.textContent || '';
      if (!text) return null;
      return {
        kind: 'shape-paragraph',
        shapeIndex,
        paragraphIndex: 0,
        runIndex: 0,
        text,
        element: textEl,
        runElements: []
      };
    }

    const firstRun = seedRuns[0];
    if (!firstRun) return null;

    const paragraph = firstRun.closest('tspan[data-ooxml-para-idx]');
    const paragraphIndex = Number(paragraph?.getAttribute('data-ooxml-para-idx') ?? 0);
    const resolvedParagraphIndex = Number.isFinite(paragraphIndex) ? paragraphIndex : 0;
    const lineContainers = this.collectParagraphLineContainers(textEl, resolvedParagraphIndex);
    // Prefer the first container that actually holds runs so geometry helpers
    // (caret height, editor box, row count) never anchor to a bullet/number
    // marker container, whose char offsets don't match the run-only edit model.
    const firstRunContainer = lineContainers.find((container) => this.collectParagraphRuns(container).length > 0);
    const geometryElement = firstRunContainer
      ?? lineContainers[0]
      ?? (paragraph && isSVGTSpanElement(paragraph) ? paragraph : textEl);
    const runElements = lineContainers.flatMap((container) => this.collectParagraphRuns(container));
    const firstRunIndex = Number(firstRun.getAttribute('data-ooxml-run-idx') ?? 0);
    const text = this.getParagraphPlainText(lineContainers.length > 0 ? lineContainers : [paraContainer].filter(isSVGTSpanElement));

    return {
      kind: 'shape-paragraph',
      shapeIndex,
      paragraphIndex: resolvedParagraphIndex,
      runIndex: Number.isFinite(firstRunIndex) ? firstRunIndex : 0,
      text,
      element: geometryElement,
      runElements
    };
  }

  private getGeneratedTextEditTarget(element: Element | null): GeneratedTextEditTarget | null {
    if (!this.engine) return null;

    const textElement = element?.closest('text');
    const shape = textElement?.closest('g[data-ooxml-shape-idx]');
    const shapeIndex = getShapeIndex(shape ?? null);
    const kind = textElement?.getAttribute('data-native-powerpoint-generated-kind') as GeneratedTextKind | null;
    const labelIndex = Number(textElement?.getAttribute('data-native-powerpoint-label-index'));
    const occurrence = Number(textElement?.getAttribute('data-native-powerpoint-label-occurrence'));
    if (
      !isSVGTextElement(textElement)
      || shapeIndex === null
      || (kind !== 'chart' && kind !== 'table')
      || !Number.isFinite(labelIndex)
      || !Number.isFinite(occurrence)
    ) {
      return null;
    }

    const target: GeneratedTextEditTarget = {
      kind,
      shapeIndex,
      labelIndex,
      occurrence,
      previousText: textElement.textContent || '',
      text: textElement.textContent || '',
      element: textElement
    };
    return this.engine.canUpdateGeneratedText(this.currentSlide, shapeIndex, target) ? target : null;
  }

  private markGeneratedTextEditability(svg: SVGSVGElement): void {
    svg.querySelectorAll('text[data-native-powerpoint-generated-kind]').forEach((text) => {
      if (this.getGeneratedTextEditTarget(text)) {
        text.classList.add('native-powerpoint-editable-text');
      } else {
        text.classList.add('native-powerpoint-generated-readonly');
      }
    });
  }

  private showGeneratedTextNotice(): void {
    if (this.hasShownGeneratedTextNotice) return;

    this.hasShownGeneratedTextNotice = true;
    pptNotice('powerpoint:notice.generatedChartLabel');
  }

  private updateSelectionOverlay(): void {
    this.updateMultiSelectionBoxes();
    const isMultiSelection = this.selectedShapeIndices.size > 1;
    if (!this.canvasPane || (!isMultiSelection && this.selectedShapeIndex === null)) {
      this.removeSelectionOverlay();
      this.updateTextToolbar();
      return;
    }

    if (!this.selectionOverlay) {
      this.selectionOverlay = this.canvasPane.createDiv({ cls: 'native-powerpoint-selection-box' });
      if (this.canEdit()) {
        for (const side of ['n', 'e', 's', 'w'] as const) {
          const moveEl = this.selectionOverlay.createDiv({
            cls: `native-powerpoint-move-border native-powerpoint-move-border-${side}`
          });
          moveEl.addEventListener('pointerdown', (event) => {
            event.preventDefault();
            event.stopPropagation();
            this.startCurrentSelectionDrag(event, 'move');
          });
        }

        for (const handle of ['n', 'e', 's', 'w', 'nw', 'ne', 'sw', 'se'] as HandleName[]) {
          const handleEl = this.selectionOverlay.createDiv({
            cls: `native-powerpoint-resize-handle native-powerpoint-resize-${handle}`
          });
          handleEl.addEventListener('pointerdown', (event) => {
            event.preventDefault();
            event.stopPropagation();
            this.startCurrentSelectionDrag(event, 'resize', handle);
          });
        }

        const rotateStem = this.selectionOverlay.createDiv({ cls: 'native-powerpoint-rotate-stem' });
        rotateStem.setAttribute('aria-hidden', 'true');
        const rotateHandle = this.selectionOverlay.createDiv({ cls: 'native-powerpoint-rotate-handle' });
        rotateHandle.setAttribute('aria-label', pptT('powerpoint:accessibility.rotateObject'));
        rotateHandle.setAttribute('data-tooltip', pptT('powerpoint:accessibility.rotateObject'));
        rotateHandle.addEventListener('pointerdown', (event) => {
          event.preventDefault();
          event.stopPropagation();
          this.startCurrentSelectionDrag(event, 'rotate');
        });
      }
    }

    this.selectionOverlay.classList.toggle('native-powerpoint-multi-selection-outline', isMultiSelection);
    const laidOut = isMultiSelection
      ? this.applyMultiSelectionOverlayLayout()
      : this.applySelectionOverlayLayout();
    if (!laidOut) {
      this.removeSelectionOverlay();
    }

    this.updateTextToolbar();
  }

  /** Send the shared outline controls to either one shape or the selected group. */
  private startCurrentSelectionDrag(
    event: PointerEvent,
    mode: DragState['mode'],
    handle?: HandleName,
  ): void {
    if (this.selectedShapeIndices.size > 1) {
      this.startGroupDrag(event, mode, handle);
      return;
    }

    if (mode === 'rotate') {
      this.startRotateDrag(event);
      return;
    }
    this.startDrag(event, mode, handle);
  }

  private removeSelectionOverlay(): void {
    this.selectionOverlay?.remove();
    this.selectionOverlay = null;
  }

  private updateMultiSelectionBoxes(): void {
    this.removeMultiSelectionBoxes();
  }

  private removeMultiSelectionBoxes(): void {
    for (const box of this.multiSelectionBoxes) {
      box.remove();
    }
    this.multiSelectionBoxes = [];
  }

  private collectShapesInClientRect(left: number, top: number, right: number, bottom: number): number[] {
    if (!this.svgEl) return [];

    const indices: number[] = [];
    this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      if (!isSVGGElement(shape)) return;
      if (shape.parentElement?.closest('g[data-ooxml-shape-idx]')) return;

      const index = getShapeIndex(shape);
      if (!isSelectableShapeIndex(index)) return;

      const rect = shape.getBoundingClientRect();
      if (rect.width === 0 && rect.height === 0) return;

      const intersects =
        rect.left <= right && rect.right >= left && rect.top <= bottom && rect.bottom >= top;
      if (intersects) indices.push(index);
    });
    return indices;
  }

  private previewSelectionClasses(indices: Set<number>): void {
    this.svgEl?.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      const index = getShapeIndex(shape);
      if (index !== null && indices.has(index)) {
        shape.addClass('native-powerpoint-shape-selected');
      } else {
        shape.removeClass('native-powerpoint-shape-selected');
      }
    });
  }

  private beginMarquee(event: PointerEvent, additive: boolean): void {
    if (!this.canvasPane) return;

    this.cancelMarquee();
    this.marquee = {
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      additive,
      base: [...this.selectedShapeIndices],
      moved: false
    };
    logPptxAction('selection', 'marquee-select', {
      slide: this.currentSlide,
      shapeIndexes: [...this.selectedShapeIndices],
      additive,
    });
  }

  private updateMarquee(event: PointerEvent): void {
    if (!this.marquee || event.pointerId !== this.marquee.pointerId || !this.canvasPane) return;

    const deltaX = event.clientX - this.marquee.startClientX;
    const deltaY = event.clientY - this.marquee.startClientY;
    if (!this.marquee.moved && Math.hypot(deltaX, deltaY) < 4) return;
    if (!this.marquee.moved) {
      this.removeSelectionOverlay();
      this.removeMultiSelectionBoxes();
    }
    this.marquee.moved = true;

    if (!this.marqueeEl) {
      this.marqueeEl = this.canvasPane.createDiv({ cls: 'native-powerpoint-marquee-box' });
    }

    const paneRect = this.canvasPane.getBoundingClientRect();
    const left = Math.min(event.clientX, this.marquee.startClientX);
    const top = Math.min(event.clientY, this.marquee.startClientY);
    const width = Math.abs(deltaX);
    const height = Math.abs(deltaY);
    this.marqueeEl.setCssProps({
      left: `${left - paneRect.left + this.canvasPane.scrollLeft}px`,
      top: `${top - paneRect.top + this.canvasPane.scrollTop}px`,
      width: `${width}px`,
      height: `${height}px`
    });

    const hits = this.collectShapesInClientRect(left, top, left + width, top + height);
    const preview = new Set<number>(this.marquee.additive ? this.marquee.base : []);
    hits.forEach((index) => preview.add(index));
    this.previewSelectionClasses(preview);
    this.updateMarqueeSelectionPreview(preview);
  }

  private finishMarquee(event: PointerEvent): void {
    if (!this.marquee || event.pointerId !== this.marquee.pointerId) return;

    const marquee = this.marquee;
    this.marquee = null;
    this.marqueeEl?.remove();
    this.marqueeEl = null;
    this.removeMarqueeSelectionPreview();

    if (!marquee.moved) {
      if (marquee.additive) {
        this.suppressNextClick = true;
        this.applyMultiSelection(marquee.base);
      } else {
        this.clearSelection();
      }
      return;
    }

    this.suppressNextClick = true;
    const left = Math.min(event.clientX, marquee.startClientX);
    const top = Math.min(event.clientY, marquee.startClientY);
    const right = Math.max(event.clientX, marquee.startClientX);
    const bottom = Math.max(event.clientY, marquee.startClientY);
    const hits = this.collectShapesInClientRect(left, top, right, bottom);
    const finalSet = new Set<number>(marquee.additive ? marquee.base : []);
    hits.forEach((index) => finalSet.add(index));
    debugLog('selection', 'PowerPoint marquee selection finished', {
      slide: this.currentSlide,
      additive: marquee.additive,
      shapeIndexes: [...finalSet],
      moved: marquee.moved,
    });
    this.applyMultiSelection([...finalSet]);
  }

  private cancelMarquee(): void {
    this.marquee = null;
    this.marqueeEl?.remove();
    this.marqueeEl = null;
    this.removeMarqueeSelectionPreview();
  }

  private removeMarqueeSelectionPreview(): void {
    this.marqueeSelectionPreview?.remove();
    this.marqueeSelectionPreview = null;
  }

  /** Render a passive outline around the live marquee result; controls wait for pointer-up. */
  private updateMarqueeSelectionPreview(indices: Iterable<number>): void {
    if (!this.canvasPane) return;
    const box = this.getSelectionBoxForShapeIndices(indices);
    if (!box) {
      this.removeMarqueeSelectionPreview();
      return;
    }

    if (!this.marqueeSelectionPreview) {
      this.marqueeSelectionPreview = this.canvasPane.createDiv({
        cls: 'native-powerpoint-marquee-selection-preview'
      });
    }
    this.marqueeSelectionPreview.setCssProps({
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: `${box.width}px`,
      height: `${box.height}px`,
    });
  }

  private getSelectionBoxForShapeIndices(
    indices: Iterable<number>,
  ): { left: number; top: number; width: number; height: number } | null {
    if (!this.svgEl) return null;

    let left = Infinity;
    let top = Infinity;
    let right = -Infinity;
    let bottom = -Infinity;
    for (const index of indices) {
      const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
      if (!isSVGGElement(shape)) continue;
      const box = this.getElementBox(this.getShapeSelectionElement(shape));
      if (!box) continue;
      left = Math.min(left, box.left);
      top = Math.min(top, box.top);
      right = Math.max(right, box.left + box.width);
      bottom = Math.max(bottom, box.top + box.height);
    }

    return Number.isFinite(left)
      ? { left, top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) }
      : null;
  }

  private getMultiSelectionBox(): { left: number; top: number; width: number; height: number } | null {
    return this.getSelectionBoxForShapeIndices(this.selectedShapeIndices);
  }

  private getGroupTransformBounds(transforms: Iterable<ShapeTransform>): ShapeTransform | null {
    let left = Infinity;
    let top = Infinity;
    let right = -Infinity;
    let bottom = -Infinity;
    for (const transform of transforms) {
      left = Math.min(left, transform.x);
      top = Math.min(top, transform.y);
      right = Math.max(right, transform.x + transform.cx);
      bottom = Math.max(bottom, transform.y + transform.cy);
    }

    return Number.isFinite(left)
      ? { x: left, y: top, cx: Math.max(1, right - left), cy: Math.max(1, bottom - top), rot: 0 }
      : null;
  }

  private applyMultiSelectionOverlayLayout(): boolean {
    if (!this.selectionOverlay) return false;
    const box = this.getMultiSelectionBox();
    if (!box) return false;

    this.selectionOverlay.style.removeProperty('transform');
    this.selectionOverlay.style.removeProperty('transform-origin');
    this.selectionOverlay.setCssProps({
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: `${box.width}px`,
      height: `${box.height}px`,
    });
    return true;
  }

  private startGroupDrag(
    event: PointerEvent,
    mode: DragState['mode'] = 'move',
    handle?: HandleName,
  ): void {
    if (!this.engine || !this.svgEl) return;
    const action = mode === 'resize' ? 'resize objects' : mode === 'rotate' ? 'rotate objects' : 'move objects';
    if (!this.ensureEditable(action)) return;

    const startPoint = this.getSvgPoint(event);
    const startBox = this.getMultiSelectionBox();
    if (!startPoint || !startBox) return;

    const start = new Map<number, ShapeTransform>();
    let hasTextChildren = false;
    let hasPictureChildren = false;
    let hasRotatedChildren = false;
    for (const index of this.selectedShapeIndices) {
      const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
      if (!isSVGGElement(shape)) continue;
      const transform = cloneTransform(this.engine.getShapeTransform(shape));
      start.set(index, transform);
      hasTextChildren ||= shape.querySelector('text') !== null;
      hasPictureChildren ||= this.isPictureShape(shape);
      hasRotatedChildren ||= this.shapeHasRotation(transform);
    }
    const startBounds = this.getGroupTransformBounds(start.values());
    if (!startBounds) return;

    const overlayRect = this.selectionOverlay?.getBoundingClientRect();
    const centerClientX = overlayRect ? overlayRect.left + overlayRect.width / 2 : undefined;
    const centerClientY = overlayRect ? overlayRect.top + overlayRect.height / 2 : undefined;
    this.snapController.beginDrag(new Set(this.selectedShapeIndices));
    this.groupDrag = {
      mode,
      handle,
      pointerId: event.pointerId,
      startPoint,
      startClientX: event.clientX,
      startClientY: event.clientY,
      startBox,
      startBounds,
      latestBounds: cloneTransform(startBounds),
      centerClientX,
      centerClientY,
      startAngle: mode === 'rotate' && centerClientX !== undefined && centerClientY !== undefined
        ? Math.atan2(event.clientY - centerClientY, event.clientX - centerClientX)
        : undefined,
      lastAngle: mode === 'rotate' && centerClientX !== undefined && centerClientY !== undefined
        ? Math.atan2(event.clientY - centerClientY, event.clientX - centerClientX)
        : undefined,
      accumulatedRotationDegrees: 0,
      previewOriginalTransforms: new Map(),
      previewOriginalText: new Map(),
      start,
      latest: new Map(start),
      moved: false
    };
    logPptxAction('selection', 'group-transform', {
      slide: this.currentSlide,
      shapeIndexes: [...start.keys()],
      mode,
      handle,
      startBounds,
      hasTextChildren,
      hasPictureChildren,
      hasRotatedChildren,
    });
  }

  private updateGroupDrag(event: PointerEvent): void {
    if (!this.groupDrag || event.pointerId !== this.groupDrag.pointerId || !this.engine || !this.svgEl) {
      return;
    }

    const deltaClientX = event.clientX - this.groupDrag.startClientX;
    const deltaClientY = event.clientY - this.groupDrag.startClientY;
    if (!this.groupDrag.moved && Math.hypot(deltaClientX, deltaClientY) < 3) return;
    this.groupDrag.moved = true;

    if (this.groupDrag.mode === 'rotate') {
      this.updateGroupRotateDrag(event);
      return;
    }

    const point = this.getSvgPoint(event);
    if (!point) return;
    const scale = this.engine.getSlideScale(this.svgEl);
    const dx = (point.x - this.groupDrag.startPoint.x) * scale;
    const dy = (point.y - this.groupDrag.startPoint.y) * scale;

    if (this.groupDrag.mode === 'resize') {
      const nextBounds = this.getGroupResizeBounds(this.groupDrag, dx, dy);
      this.setResizeCrossedHandleState(
        this.groupDrag.crossedHorizontal ?? false,
        this.groupDrag.crossedVertical ?? false,
      );
      const scaleX = nextBounds.cx / this.groupDrag.startBounds.cx;
      const scaleY = nextBounds.cy / this.groupDrag.startBounds.cy;
      this.groupDrag.latest = this.scaleGroupTransforms(
        this.groupDrag.start,
        this.groupDrag.startBounds,
        nextBounds,
        scaleX,
        scaleY,
        this.groupDrag.crossedHorizontal ?? false,
        this.groupDrag.crossedVertical ?? false,
      );
      this.groupDrag.latestBounds = nextBounds;
      this.applyGroupOverlayBox(this.getGroupResizeOverlayBox(this.groupDrag, nextBounds));
      this.applyGroupResizePreview(this.groupDrag, nextBounds);
      return;
    }

    const nextBounds = cloneTransform(this.groupDrag.startBounds);
    nextBounds.x += dx;
    nextBounds.y += dy;
    const snap = this.snapController.computeSnap(
      nextBounds,
      new Set(this.selectedShapeIndices)
    );
    nextBounds.x += snap.dx;
    nextBounds.y += snap.dy;
    this.snapController.updateSnapGuides(snap.guideX, snap.guideY);
    this.groupDrag.latest = new Map(
      [...this.groupDrag.start.entries()].map(([index, transform]) => [
        index,
        { ...cloneTransform(transform), x: transform.x + nextBounds.x - this.groupDrag!.startBounds.x, y: transform.y + nextBounds.y - this.groupDrag!.startBounds.y }
      ])
    );
    this.groupDrag.latestBounds = nextBounds;
    const paneScale = this.getPaneEmuScale();
    this.applyGroupOverlayBox({
      left: this.groupDrag.startBox.left + (nextBounds.x - this.groupDrag.startBounds.x) * (paneScale?.x ?? 0),
      top: this.groupDrag.startBox.top + (nextBounds.y - this.groupDrag.startBounds.y) * (paneScale?.y ?? 0),
      width: this.groupDrag.startBox.width,
      height: this.groupDrag.startBox.height,
    }, paneScale === null ? { x: deltaClientX, y: deltaClientY } : undefined);
    this.applyGroupMovePreview(this.groupDrag, nextBounds);
  }

  private getGroupResizeBounds(groupDrag: GroupDragState, dx: number, dy: number): ShapeTransform {
    if (!this.engine) return cloneTransform(groupDrag.startBounds);
    const minSize = this.engine.pxToEmu(12);
    const minScaleX = Math.max(...[...groupDrag.start.values()].map((transform) => minSize / transform.cx));
    const minScaleY = Math.max(...[...groupDrag.start.values()].map((transform) => minSize / transform.cy));
    const minWidth = Math.max(minSize, groupDrag.startBounds.cx * minScaleX);
    const minHeight = Math.max(minSize, groupDrag.startBounds.cy * minScaleY);
    const next = cloneTransform(groupDrag.startBounds);

    if (groupDrag.handle?.includes('w')) {
      const width = groupDrag.startBounds.cx - dx;
      if (width >= minWidth) {
        next.x = groupDrag.startBounds.x + dx;
        next.cx = width;
        groupDrag.crossedHorizontal = false;
      } else if (width <= -minWidth) {
        // Keep OOXML extents positive and mirror the group across its fixed east edge.
        next.x = groupDrag.startBounds.x + groupDrag.startBounds.cx;
        next.cx = -width;
        groupDrag.crossedHorizontal = true;
      } else {
        next.cx = minWidth;
        next.x = groupDrag.startBounds.x + groupDrag.startBounds.cx - minWidth;
        groupDrag.crossedHorizontal = false;
      }
    }
    if (groupDrag.handle?.includes('e')) {
      const width = groupDrag.startBounds.cx + dx;
      if (width >= minWidth) {
        next.cx = width;
        groupDrag.crossedHorizontal = false;
      } else if (width <= -minWidth) {
        next.x = groupDrag.startBounds.x + width;
        next.cx = -width;
        groupDrag.crossedHorizontal = true;
      } else {
        next.cx = minWidth;
        groupDrag.crossedHorizontal = false;
      }
    }
    if (groupDrag.handle?.includes('n')) {
      const height = groupDrag.startBounds.cy - dy;
      if (height >= minHeight) {
        next.y = groupDrag.startBounds.y + dy;
        next.cy = height;
        groupDrag.crossedVertical = false;
      } else if (height <= -minHeight) {
        // Keep OOXML extents positive and mirror the group across its fixed south edge.
        next.y = groupDrag.startBounds.y + groupDrag.startBounds.cy;
        next.cy = -height;
        groupDrag.crossedVertical = true;
      } else {
        next.cy = minHeight;
        next.y = groupDrag.startBounds.y + groupDrag.startBounds.cy - minHeight;
        groupDrag.crossedVertical = false;
      }
    }
    if (groupDrag.handle?.includes('s')) {
      const height = groupDrag.startBounds.cy + dy;
      if (height >= minHeight) {
        next.cy = height;
        groupDrag.crossedVertical = false;
      } else if (height <= -minHeight) {
        next.y = groupDrag.startBounds.y + height;
        next.cy = -height;
        groupDrag.crossedVertical = true;
      } else {
        next.cy = minHeight;
        groupDrag.crossedVertical = false;
      }
    }
    return next;
  }

  private scaleGroupTransforms(
    start: Map<number, ShapeTransform>,
    startBounds: ShapeTransform,
    nextBounds: ShapeTransform,
    scaleX: number,
    scaleY: number,
    crossedHorizontal = false,
    crossedVertical = false,
  ): Map<number, ShapeTransform> {
    return new Map(
      [...start.entries()].map(([index, transform]) => [
        index,
        {
          ...cloneTransform(transform),
          x: crossedHorizontal
            ? nextBounds.x + (startBounds.cx - (transform.x - startBounds.x + transform.cx)) * scaleX
            : nextBounds.x + (transform.x - startBounds.x) * scaleX,
          y: crossedVertical
            ? nextBounds.y + (startBounds.cy - (transform.y - startBounds.y + transform.cy)) * scaleY
            : nextBounds.y + (transform.y - startBounds.y) * scaleY,
          cx: Math.max(1, transform.cx * scaleX),
          cy: Math.max(1, transform.cy * scaleY),
        }
      ])
    );
  }

  private getGroupResizeOverlayBox(
    groupDrag: GroupDragState,
    nextBounds: ShapeTransform,
  ): { left: number; top: number; width: number; height: number } {
    const scaleX = nextBounds.cx / groupDrag.startBounds.cx;
    const scaleY = nextBounds.cy / groupDrag.startBounds.cy;
    const box = {
      ...groupDrag.startBox,
      width: groupDrag.startBox.width * scaleX,
      height: groupDrag.startBox.height * scaleY,
    };
    if (groupDrag.handle?.includes('w')) {
      box.left = groupDrag.crossedHorizontal
        ? groupDrag.startBox.left + groupDrag.startBox.width
        : groupDrag.startBox.left + groupDrag.startBox.width - box.width;
    } else if (groupDrag.handle?.includes('e') && groupDrag.crossedHorizontal) {
      box.left = groupDrag.startBox.left - box.width;
    }
    if (groupDrag.handle?.includes('n')) {
      box.top = groupDrag.crossedVertical
        ? groupDrag.startBox.top + groupDrag.startBox.height
        : groupDrag.startBox.top + groupDrag.startBox.height - box.height;
    } else if (groupDrag.handle?.includes('s') && groupDrag.crossedVertical) {
      box.top = groupDrag.startBox.top - box.height;
    }
    return box;
  }

  private applyGroupOverlayBox(
    box: { left: number; top: number; width: number; height: number } | null,
    fallbackDelta?: { x: number; y: number },
  ): void {
    if (!this.selectionOverlay || !box || !this.groupDrag) return;
    this.selectionOverlay.style.removeProperty('transform');
    this.selectionOverlay.style.removeProperty('transform-origin');
    this.selectionOverlay.setCssProps({
      left: `${box.left + (fallbackDelta?.x ?? 0)}px`,
      top: `${box.top + (fallbackDelta?.y ?? 0)}px`,
      width: `${box.width}px`,
      height: `${box.height}px`,
    });
  }

  /**
   * Live-translate every selected shape during a multi-object move so the
   * shapes follow the shared group outline (not just the selection chrome).
   */
  private applyGroupMovePreview(
    groupDrag: GroupDragState,
    nextBounds: ShapeTransform,
  ): void {
    if (!this.svgEl || !this.engine) return;
    const slideScale = this.engine.getSlideScale(this.svgEl);
    if (!slideScale) return;

    const dxUser = (nextBounds.x - groupDrag.startBounds.x) / slideScale;
    const dyUser = (nextBounds.y - groupDrag.startBounds.y) / slideScale;
    const transform = `translate(${this.formatSvgNumber(dxUser)} ${this.formatSvgNumber(dyUser)})`;
    let objectCount = 0;
    for (const index of groupDrag.start.keys()) {
      const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
      if (!isSVGGElement(shape)) continue;

      if (!groupDrag.previewOriginalTransforms?.has(index)) {
        groupDrag.previewOriginalTransforms?.set(index, shape.getAttribute('transform'));
      }
      const original = groupDrag.previewOriginalTransforms?.get(index)?.trim() ?? '';
      shape.classList.add('native-powerpoint-shape-drag-preview');
      shape.setAttribute('transform', original ? `${transform} ${original}` : transform);
      objectCount += 1;
    }
    const firstPreview = groupDrag.previewObjectCount === undefined;
    groupDrag.previewObjectCount = objectCount;
    if (firstPreview && objectCount > 0) {
      debugLog('selection', 'PowerPoint group move preview applied', {
        op: 'group-move-preview',
        slide: this.currentSlide,
        shapeIndexes: [...groupDrag.start.keys()],
        objectPreviewCount: objectCount,
        dxUser,
        dyUser,
      });
    }
  }

    /** Build the affine transform shared by all live group-resize previews. */
  private getGroupResizePreviewTransform(
    groupDrag: GroupDragState,
    nextBounds: ShapeTransform,
  ): string | null {
    if (!this.engine || !this.svgEl) return null;
    const slideScale = this.engine.getSlideScale(this.svgEl);
    if (!slideScale) return null;

    const scaleX = (groupDrag.crossedHorizontal ? -1 : 1)
      * nextBounds.cx / groupDrag.startBounds.cx;
    const scaleY = (groupDrag.crossedVertical ? -1 : 1)
      * nextBounds.cy / groupDrag.startBounds.cy;
    const anchorX = groupDrag.handle?.includes('w')
      ? groupDrag.startBounds.x + groupDrag.startBounds.cx
      : groupDrag.handle?.includes('e')
        ? groupDrag.startBounds.x
        : groupDrag.startBounds.x + groupDrag.startBounds.cx / 2;
    const anchorY = groupDrag.handle?.includes('n')
      ? groupDrag.startBounds.y + groupDrag.startBounds.cy
      : groupDrag.handle?.includes('s')
        ? groupDrag.startBounds.y
        : groupDrag.startBounds.y + groupDrag.startBounds.cy / 2;
    const x = anchorX / slideScale;
    const y = anchorY / slideScale;
    return `translate(${this.formatSvgNumber(x)} ${this.formatSvgNumber(y)}) scale(${this.formatSvgNumber(scaleX)} ${this.formatSvgNumber(scaleY)}) translate(${this.formatSvgNumber(-x)} ${this.formatSvgNumber(-y)})`;
  }

  /**
   * Keep every selected object inside the group outline while it is resized.
   * Text gets an inverse affine transform below so only its frame changes: its
   * glyphs remain their natural size and its SVG lines can wrap to the new width.
   */
  private applyGroupResizePreview(groupDrag: GroupDragState, nextBounds: ShapeTransform): void {
    if (!this.svgEl || !this.engine) return;
    const transform = this.getGroupResizePreviewTransform(groupDrag, nextBounds);
    const slideScale = this.engine.getSlideScale(this.svgEl);
    if (!transform || !slideScale) return;
    const scaleX = (groupDrag.crossedHorizontal ? -1 : 1)
      * nextBounds.cx / groupDrag.startBounds.cx;
    const scaleY = (groupDrag.crossedVertical ? -1 : 1)
      * nextBounds.cy / groupDrag.startBounds.cy;
    const anchorX = (groupDrag.handle?.includes('w')
      ? groupDrag.startBounds.x + groupDrag.startBounds.cx
      : groupDrag.handle?.includes('e')
        ? groupDrag.startBounds.x
        : groupDrag.startBounds.x + groupDrag.startBounds.cx / 2) / slideScale;
    const anchorY = (groupDrag.handle?.includes('n')
      ? groupDrag.startBounds.y + groupDrag.startBounds.cy
      : groupDrag.handle?.includes('s')
        ? groupDrag.startBounds.y
        : groupDrag.startBounds.y + groupDrag.startBounds.cy / 2) / slideScale;

    let objectCount = 0;
    let textTransformCount = 0;
    let textReflowCount = 0;
    for (const index of groupDrag.start.keys()) {
      const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
      if (!isSVGGElement(shape)) continue;

      if (!groupDrag.previewOriginalTransforms?.has(index)) {
        groupDrag.previewOriginalTransforms?.set(index, shape.getAttribute('transform'));
      }
      const original = groupDrag.previewOriginalTransforms?.get(index)?.trim() ?? '';
      shape.classList.add('native-powerpoint-shape-drag-preview');
      shape.setAttribute('transform', original ? `${transform} ${original}` : transform);
      objectCount += 1;

      const start = groupDrag.start.get(index);
      const next = groupDrag.latest.get(index);
      const text = shape.querySelector('text');
      if (!start || !next || !isSVGTextElement(text)) continue;

      if (!groupDrag.previewOriginalText?.has(index)) {
        groupDrag.previewOriginalText?.set(index, text.cloneNode(true) as SVGTextElement);
      }
      const originalText = groupDrag.previewOriginalText?.get(index) ?? null;
      try {
        if (this.applyTextResizePreview(
          shape,
          originalText,
          start,
          slideScale,
          scaleX,
          scaleY,
          anchorX,
          anchorY,
          (error) => {
            groupDrag.previewTextReflowError ??= cleanError(error);
          },
        )) {
          textReflowCount += 1;
        }
        textTransformCount += 1;
      } catch (error) {
        // A malformed text frame must not prevent subsequent selected text
        // frames from following a crossed group resize.
        groupDrag.previewTextReflowError ??= cleanError(error);
      }
    }
    groupDrag.previewObjectCount = objectCount;
    groupDrag.previewTextTransformCount = textTransformCount;
    groupDrag.previewTextReflowCount = textReflowCount;
  }

  private restoreGroupShapePreviews(groupDrag: GroupDragState): void {
    for (const [index, original] of groupDrag.previewOriginalTransforms ?? []) {
      const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
      if (!isSVGGElement(shape)) continue;
      shape.classList.remove('native-powerpoint-shape-drag-preview');
      if (original === null) {
        shape.removeAttribute('transform');
      } else {
        shape.setAttribute('transform', original);
      }

      this.restoreShapeTextPreview(
        shape,
        groupDrag.previewOriginalText?.get(index) ?? null,
      );
    }
  }

  /**
   * Restore a cloned text subtree before another preview pass. Reflowing from
   * the rendered baseline avoids accumulating temporary line breaks as the
   * pointer moves back and forth.
   */
  private restoreShapeTextPreview(
    shape: SVGGElement,
    original: SVGTextElement | null,
  ): SVGTextElement | null {
    const current = shape.querySelector('text');
    if (!isSVGTextElement(current)) return null;
    if (!original) return current;

    const replacement = original.cloneNode(true) as SVGTextElement;
    current.replaceWith(replacement);
    return replacement;
  }

  /**
   * Cancel the outer affine scale for a text child while retaining the text
   * anchor's relative position. For a group matrix G and desired translation
   * D, the child receives G⁻¹D, keeping its glyphs readable after a crossed
   * (negative) resize while its in-frame margins scale with the object.
   */
  private getTextResizeCompensationTransform(
    scaleX: number,
    scaleY: number,
    anchorX: number,
    anchorY: number,
    translateX: number,
    translateY: number,
  ): string | null {
    if (!Number.isFinite(scaleX) || !Number.isFinite(scaleY) || scaleX === 0 || scaleY === 0) {
      return null;
    }
    const inverseX = 1 / scaleX;
    const inverseY = 1 / scaleY;
    const offsetX = (translateX - anchorX * (1 - scaleX)) / scaleX;
    const offsetY = (translateY - anchorY * (1 - scaleY)) / scaleY;
    return `matrix(${this.formatSvgNumber(inverseX)} 0 0 ${this.formatSvgNumber(inverseY)} ${this.formatSvgNumber(offsetX)} ${this.formatSvgNumber(offsetY)})`;
  }

  /** Translate the text anchor as its original relative point moves through the group resize. */
  private getRelativeTextPreviewTranslation(
    textAnchor: { x: number; y: number },
    scaleX: number,
    scaleY: number,
    anchorX: number,
    anchorY: number,
  ): { x: number; y: number } {
    return {
      x: (scaleX - 1) * (textAnchor.x - anchorX),
      y: (scaleY - 1) * (textAnchor.y - anchorY),
    };
  }

  /** Find a text line's unscaled SVG anchor, falling back to its shape frame. */
  private getTextPreviewAnchor(
    text: SVGTextElement,
    fallback: { x: number; y: number },
  ): { x: number; y: number } {
    const line = text.querySelector('tspan[data-ooxml-para-idx]');
    const parseCoordinate = (value: string | null): number | null => {
      const first = value?.trim().split(/[\s,]+/)[0] ?? '';
      const coordinate = Number(first);
      return Number.isFinite(coordinate) ? coordinate : null;
    };
    return {
      x: parseCoordinate(line?.getAttribute('x') ?? text.getAttribute('x')) ?? fallback.x,
      y: parseCoordinate(line?.getAttribute('y') ?? text.getAttribute('y')) ?? fallback.y,
    };
  }

  /** Apply a relative-position, no-glyph-scale text preview, then reflow its resized frame. */
  private applyTextResizePreview(
    shape: SVGGElement,
    originalText: SVGTextElement | null,
    start: ShapeTransform,
    slideScale: number,
    scaleX: number,
    scaleY: number,
    anchorX: number,
    anchorY: number,
    onReflowError?: (error: unknown) => void,
  ): boolean {
    const text = this.restoreShapeTextPreview(shape, originalText);
    if (!text || !Number.isFinite(slideScale) || slideScale === 0) return false;

    const textAnchor = this.getTextPreviewAnchor(text, {
      x: start.x / slideScale,
      y: start.y / slideScale,
    });
    const translation = this.getRelativeTextPreviewTranslation(
      textAnchor,
      scaleX,
      scaleY,
      anchorX,
      anchorY,
    );

    const compensation = this.getTextResizeCompensationTransform(
      scaleX,
      scaleY,
      anchorX,
      anchorY,
      translation.x,
      translation.y,
    );
    if (!compensation) return false;

    const base = text.getAttribute('transform')?.trim() ?? '';
    text.setAttribute('transform', base ? `${compensation} ${base}` : compensation);
    if (Math.abs(Math.abs(scaleX) - 1) <= 0.001) return false;
    try {
      return this.reflowShapeTextResizePreview(shape, text);
    } catch (error) {
      onReflowError?.(error);
      return false;
    }
  }

  /**
   * Rebuild temporary visual lines for all run-backed paragraphs in one text
   * frame. It never writes OOXML; the authoritative renderer still owns the
   * persisted layout after pointer-up.
   */
  private reflowShapeTextResizePreview(shape: SVGGElement, text: SVGTextElement): boolean {
    const frame = shape.querySelector(':scope > rect');
    const frameBox = frame ? this.getElementBox(frame) : null;
    if (!frameBox) return false;

    const paragraphs = new Map<string, SVGTSpanElement[]>();
    for (const child of Array.from(text.children)) {
      if (!isSVGTSpanElement(child)) continue;
      const paragraphIndex = child.getAttribute('data-ooxml-para-idx');
      if (paragraphIndex === null || this.collectParagraphRuns(child).length === 0) continue;
      const lines = paragraphs.get(paragraphIndex) ?? [];
      lines.push(child);
      paragraphs.set(paragraphIndex, lines);
    }

    let changed = false;
    let downstreamYOffset = 0;
    for (const lineContainers of paragraphs.values()) {
      const firstLine = lineContainers[0];
      const firstRun = firstLine ? this.collectParagraphRuns(firstLine)[0] : null;
      if (!firstLine || !firstRun) continue;

      const firstLineBox = this.getElementBox(firstLine);
      const baseY = Number(firstLine.getAttribute('y'));
      const lineStep = this.getLivePreviewLineStep(lineContainers, firstRun);
      if (!firstLineBox || !Number.isFinite(baseY) || lineStep === null) continue;

      const inset = Math.max(0, firstLineBox.left - frameBox.left);
      const maxWidth = frameBox.width - inset * 2;
      const measure = this.createInlinePreviewTextMeasurer(firstRun);
      const paragraphText = this.getParagraphPlainText(lineContainers);
      if (!measure || paragraphText.includes('\n') || !Number.isFinite(maxWidth) || maxWidth < 4) continue;

      const lines = wrapTextForPreview(paragraphText, maxWidth, measure);
      if (lines.join('') !== paragraphText) continue;
      const previousRuns = lineContainers.flatMap((line) => this.collectParagraphRuns(line));
      const previousRunTexts = previousRuns.map((run) => run.textContent || '');
      if (previousRunTexts.length === 0) continue;

      const nextRunTexts = redistributeTextAcrossVisualRuns(previousRunTexts, paragraphText);
      const boundaries = [0];
      for (const runText of nextRunTexts) {
        boundaries.push((boundaries[boundaries.length - 1] ?? 0) + runText.length);
      }

      let offset = 0;
      const nextLines = lines.map((lineText, lineIndex) => {
        const line = firstLine.cloneNode(false) as SVGTSpanElement;
        line.setAttribute('y', `${baseY + downstreamYOffset + lineStep * lineIndex}`);
        line.removeAttribute('dy');
        const lineStart = offset;
        const lineEnd = lineStart + lineText.length;
        offset = lineEnd;

        for (let runIndex = 0; runIndex < previousRunTexts.length; runIndex++) {
          const runStart = boundaries[runIndex] ?? lineEnd;
          const runEnd = boundaries[runIndex + 1] ?? runStart;
          const startOffset = Math.max(lineStart, runStart);
          const endOffset = Math.min(lineEnd, runEnd);
          if (endOffset <= startOffset) continue;
          const source = previousRuns[runIndex] ?? firstRun;
          const run = source.cloneNode(true) as SVGTSpanElement;
          run.textContent = paragraphText.slice(startOffset, endOffset);
          line.appendChild(run);
        }

        if (line.children.length === 0) {
          const emptyRun = firstRun.cloneNode(true) as SVGTSpanElement;
          emptyRun.textContent = '';
          line.appendChild(emptyRun);
        }
        return line;
      });

      const parent = firstLine.parentNode;
      if (!parent) continue;
      const anchor = lineContainers[lineContainers.length - 1]?.nextSibling ?? null;
      for (const line of nextLines) parent.insertBefore(line, anchor);
      for (const line of lineContainers) line.remove();
      downstreamYOffset += lineStep * (nextLines.length - lineContainers.length);
      changed ||= nextLines.length !== lineContainers.length
        || nextLines.some((line, index) => line.textContent !== lineContainers[index]?.textContent);
    }
    return changed;
  }

  private updateGroupRotateDrag(event: PointerEvent): void {
    if (!this.groupDrag || !this.engine) return;
    const centerX = this.groupDrag.centerClientX;
    const centerY = this.groupDrag.centerClientY;
    if (centerX === undefined || centerY === undefined || this.groupDrag.startAngle === undefined) return;

    const angle = Math.atan2(event.clientY - centerY, event.clientX - centerX);
    let degrees = this.advanceContinuousRotation(this.groupDrag, angle);
    if (event.shiftKey) degrees = Math.round(degrees / 15) * 15;
    const rotationSnapTarget = this.getCardinalRotationSnap(degrees);
    degrees = rotationSnapTarget ?? degrees;
    const radians = (degrees * Math.PI) / 180;
    const cos = Math.cos(radians);
    const sin = Math.sin(radians);
    const pivotX = this.groupDrag.startBounds.x + this.groupDrag.startBounds.cx / 2;
    const pivotY = this.groupDrag.startBounds.y + this.groupDrag.startBounds.cy / 2;

    this.groupDrag.latest = new Map(
      [...this.groupDrag.start.entries()].map(([index, transform]) => {
        const centerShapeX = transform.x + transform.cx / 2;
        const centerShapeY = transform.y + transform.cy / 2;
        const dx = centerShapeX - pivotX;
        const dy = centerShapeY - pivotY;
        const nextCenterX = pivotX + dx * cos - dy * sin;
        const nextCenterY = pivotY + dx * sin + dy * cos;
        return [index, {
          ...cloneTransform(transform),
          x: nextCenterX - transform.cx / 2,
          y: nextCenterY - transform.cy / 2,
          rot: this.engine!.degreesToOoxml(this.normalizeDegrees(this.engine!.ooxmlToDegrees(transform.rot) + degrees)),
        }];
      })
    );
    this.groupDrag.latestBounds = this.getGroupTransformBounds(this.groupDrag.latest.values())
      ?? cloneTransform(this.groupDrag.startBounds);
    this.groupDrag.rotationSnapTarget = rotationSnapTarget;
    if (this.selectionOverlay) {
      this.selectionOverlay.setCssProps({ transform: `rotate(${degrees}deg)` });
    }
  }

  private finishGroupDrag(event: PointerEvent): void {
    if (!this.groupDrag || event.pointerId !== this.groupDrag.pointerId) return;

    const groupDrag = this.groupDrag;
    this.groupDrag = null;
    this.snapController.endDrag();
    this.restoreGroupShapePreviews(groupDrag);
    this.setResizeCrossedHandleState(false, false);
    this.selectionOverlay?.style.removeProperty('transform');
    this.selectionOverlay?.style.removeProperty('transform-origin');

    if (!groupDrag.moved) return;

    const label = groupDrag.mode === 'resize'
      ? 'Resize objects'
      : groupDrag.mode === 'rotate'
        ? 'Rotate objects'
        : 'Move objects';
    this.suppressNextClick = true;
    const updates = [...groupDrag.latest.entries()].map(([index, transform]) => ({ index, transform }));
    const flipAxes = {
      horizontal: groupDrag.mode === 'resize' && (groupDrag.crossedHorizontal ?? false),
      vertical: groupDrag.mode === 'resize' && (groupDrag.crossedVertical ?? false),
    };
    debugLog('selection', 'PowerPoint group transform preview ended', {
      op: 'group-transform-end',
      slide: this.currentSlide,
      shapeIndexes: [...groupDrag.start.keys()],
      mode: groupDrag.mode,
      handle: groupDrag.handle,
      startBounds: groupDrag.startBounds,
      finalBounds: groupDrag.latestBounds,
      rotationSnapTarget: groupDrag.rotationSnapTarget ?? null,
      totalRotationDegrees: groupDrag.mode === 'rotate'
        ? groupDrag.accumulatedRotationDegrees ?? 0
        : null,
      flippedHorizontal: flipAxes.horizontal,
      flippedVertical: flipAxes.vertical,
      objectPreviewCount: groupDrag.previewObjectCount ?? groupDrag.previewOriginalTransforms?.size ?? 0,
      textPreviewCount: groupDrag.previewTextTransformCount ?? 0,
      textReflowPreviewCount: groupDrag.previewTextReflowCount ?? 0,
      textReflowPreviewError: groupDrag.previewTextReflowError ?? null,
      moved: groupDrag.moved,
    });
    void this.commitGroupTransforms(updates, label, groupDrag.mode, flipAxes);
  }

  private async commitGroupTransforms(
    updates: { index: number; transform: ShapeTransform }[],
    label = 'Move objects',
    mode: DragState['mode'] = 'move',
    flipAxes: { horizontal: boolean; vertical: boolean } = { horizontal: false, vertical: false },
  ): Promise<void> {
    if (!this.engine || updates.length === 0) return;
    const action = mode === 'resize' ? 'resize objects' : mode === 'rotate' ? 'rotate objects' : 'move objects';
    if (!this.ensureEditable(action)) return;

    try {
      const editableUpdates = updates.filter((update) => isEditableShapeIndex(update.index));
      if (editableUpdates.length < updates.length) {
        pptNotice('powerpoint:notice.objectNotEditable');
      }
      if (editableUpdates.length === 0) return;

      const changes: HistoryTransformChange[] = [];
      for (const update of editableUpdates) {
        const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${update.index}"]`);
        if (!isSVGGElement(shape)) continue;
        const before = this.engine.getShapeTransform(shape);
        if (transformsMatch(before, update.transform)) continue;
        changes.push({
          shapeIndex: update.index,
          before: cloneTransform(before),
          after: cloneTransform(update.transform)
        });
      }
      const shouldFlip = flipAxes.horizontal || flipAxes.vertical;
      if (changes.length === 0 && !shouldFlip) return;

      let history: HistoryEntry = shouldFlip
        ? this.historyController.captureSlideXml(this.currentSlide, label)
        : this.historyController.captureTransform(this.currentSlide, changes, label);
      for (const change of changes) {
        await this.engine.updateShapeTransform(this.currentSlide, change.shapeIndex, change.after);
      }
      if (flipAxes.horizontal || flipAxes.vertical) {
        for (const update of editableUpdates) {
          if (flipAxes.horizontal) await this.engine.flipShape(this.currentSlide, update.index, 'horizontal');
          if (flipAxes.vertical) await this.engine.flipShape(this.currentSlide, update.index, 'vertical');
        }
      }
      if (history.kind === 'slideXml') {
        history = this.historyController.completeSlideXml(history);
      }
      this.recordHistoryEntry(history);
      this.markDirty();
      const indices = updates.map((update) => update.index);
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.applyMultiSelection(indices);
        this.scheduleThumbnailRefresh(this.currentSlide);
      }
      debugLog('arrange', 'Committed PowerPoint group transform', {
        slide: this.currentSlide,
        label,
        mode,
        changedCount: changes.length,
        flippedHorizontal: flipAxes.horizontal,
        flippedVertical: flipAxes.vertical,
        indices: editableUpdates.map((update) => update.index)
      });
    } catch (error) {
      errorLog('arrange', 'PowerPoint group transform failed', {
        slide: this.currentSlide,
        label,
        mode,
        indices: updates.map((update) => update.index),
        error
      });
      pptNotice('powerpoint:notice.couldNotMoveObjects', { message: cleanError(error) });
    }
  }

  private getSelectedBox(): { left: number; top: number; width: number; height: number } | null {
    const selected = this.getSelectedShapeElement();
    if (selected && this.pictureUsesImageSelectionBounds(selected)) {
      const box = this.getPictureSelectionBox(selected);
      if (box) return box;
    }

    if (
      this.selectedTransform
      && this.shapeHasRotation(this.selectedTransform)
    ) {
      return this.getTransformSelectionBox(this.selectedTransform);
    }

    if (!selected) return null;

    return this.getElementBox(this.getShapeSelectionElement(selected));
  }

  private getElementBox(element: Element): { left: number; top: number; width: number; height: number } | null {
    return this.inlineGeometry.getElementBox(element);
  }

  private startDrag(event: PointerEvent, mode: 'move' | 'resize', handle?: HandleName): void {
    if (!this.engine || !this.svgEl || this.selectedTransform === null) return;
    if (!this.ensureEditable(mode === 'move' ? 'move object' : 'resize object')) return;
    if (!isEditableShapeIndex(this.selectedShapeIndex)) {
      pptNotice('powerpoint:notice.objectNotEditable');
      return;
    }

    const startPoint = this.getSvgPoint(event);
    const startBox = this.getSelectedBox();
    if (!startPoint || !startBox) return;

    const previewElement = this.getSelectedShapeElement();
    const freezeShapeDuringResize = this.shouldFreezeTextDuringResize(mode, previewElement);
    const paneEmuScale = this.getPaneEmuScale();
    const previewImageElement = previewElement ? this.getPictureImageElement(previewElement) : null;
    const previewText = previewElement?.querySelector('text');
    const previewImageAttrs = previewImageElement
      ? {
          x: Number(previewImageElement.getAttribute('x') ?? 0),
          y: Number(previewImageElement.getAttribute('y') ?? 0),
          width: Number(previewImageElement.getAttribute('width') ?? 0),
          height: Number(previewImageElement.getAttribute('height') ?? 0),
          transform: previewImageElement.getAttribute('transform'),
        }
      : null;
    const excluded = new Set(this.selectedShapeIndex === null ? [] : [this.selectedShapeIndex]);
    this.snapController.beginDrag(excluded);
    if (previewElement) {
      previewElement.classList.add('native-powerpoint-shape-drag-preview');
    }
    if (previewImageElement) {
      previewImageElement.classList.add('native-powerpoint-shape-drag-preview');
    }

    this.dragState = {
      mode,
      handle,
      pointerId: event.pointerId,
      startPoint,
      startClientX: event.clientX,
      startClientY: event.clientY,
      startBox,
      startTransform: cloneTransform(this.selectedTransform),
      latestTransform: cloneTransform(this.selectedTransform),
      previewElement,
      previewOriginalTransform: previewElement?.getAttribute('transform') ?? null,
      freezeShapeDuringResize,
      previewOriginalText: freezeShapeDuringResize && isSVGTextElement(previewText)
        ? previewText.cloneNode(true) as SVGTextElement
        : null,
      paneEmuScaleX: paneEmuScale?.x,
      paneEmuScaleY: paneEmuScale?.y,
      previewImageElement,
      previewImageAttrs,
    };
    logPptxAction('selection', 'drag', {
      slide: this.currentSlide,
      shapeIndexes: this.selectedShapeIndex === null ? [] : [this.selectedShapeIndex],
      mode,
      handle,
      startTransform: cloneTransform(this.selectedTransform),
      preview: 'shape-group',
      freezeTextDuringResize: freezeShapeDuringResize,
    });
  }

  /** Text uses a nested live reflow preview instead of inheriting the frame's SVG scale. */
  private shouldFreezeTextDuringResize(
    mode: DragState['mode'],
    previewElement: SVGGElement | null,
  ): boolean {
    return mode === 'resize' && previewElement?.querySelector('text') !== null;
  }

  /**
   * Live-drag the actual shape group (not just the selection outline). Move applies
   * an SVG translate; resize scales around the handle's fixed anchor in slide space.
   */
  private getResizeAnchorUser(
    transform: ShapeTransform,
    slideScale: number,
    handle: HandleName | undefined,
  ): { x: number; y: number } {
    const x = transform.x / slideScale;
    const y = transform.y / slideScale;
    const width = transform.cx / slideScale;
    const height = transform.cy / slideScale;
    const right = x + width;
    const bottom = y + height;
    const centerX = x + width / 2;
    const centerY = y + height / 2;

    const anchorX = handle?.includes('w') ? right : handle?.includes('e') ? x : centerX;
    const anchorY = handle?.includes('n') ? bottom : handle?.includes('s') ? y : centerY;
    return { x: anchorX, y: anchorY };
  }

  private updateShapeTransformPreview(transform: ShapeTransform): void {
    if (!this.dragState || !this.engine || !this.svgEl) return;
    const element = this.dragState.previewElement;
    if (!element) return;

    const slideScale = this.engine.getSlideScale(this.svgEl);
    if (!slideScale) return;

    const start = this.dragState.startTransform;
    const dxUser = (transform.x - start.x) / slideScale;
    const dyUser = (transform.y - start.y) / slideScale;
    const sx = (this.dragState.crossedHorizontal ? -1 : 1)
      * (start.cx > 0 ? transform.cx / start.cx : 1);
    const sy = (this.dragState.crossedVertical ? -1 : 1)
      * (start.cy > 0 ? transform.cy / start.cy : 1);
    const base = this.dragState.previewOriginalTransform?.trim() ?? '';
    const parts: string[] = [];

    if (this.dragState.mode === 'rotate') {
      const startDegrees = this.engine.ooxmlToDegrees(start.rot);
      const nextDegrees = this.engine.ooxmlToDegrees(transform.rot);
      const unsignedDelta = this.normalizeDegrees(nextDegrees - startDegrees);
      const deltaDegrees = unsignedDelta > 180 ? unsignedDelta - 360 : unsignedDelta;
      const centerX = (start.x + start.cx / 2) / slideScale;
      const centerY = (start.y + start.cy / 2) / slideScale;
      parts.push(
        `rotate(${this.formatSvgNumber(deltaDegrees)} ${this.formatSvgNumber(centerX)} ${this.formatSvgNumber(centerY)})`
      );
    } else if (this.dragState.mode === 'resize' && (sx !== 1 || sy !== 1)) {
      const anchor = this.getResizeAnchorUser(start, slideScale, this.dragState.handle);
      // Scaling around the opposite, fixed handle already produces the new x/y
      // for west/north resizes. Adding the pointer delta here applies that
      // movement twice, separating the shape's fill from its selection outline.
      parts.push(`translate(${anchor.x} ${anchor.y})`);
      parts.push(`scale(${sx} ${sy})`);
      parts.push(`translate(${-anchor.x} ${-anchor.y})`);
    } else {
      parts.push(`translate(${dxUser} ${dyUser})`);
    }

    if (base) parts.push(base);
    element.setAttribute('transform', parts.join(' '));

    if (this.dragState.mode === 'resize' && this.dragState.freezeShapeDuringResize) {
      const anchor = this.getResizeAnchorUser(start, slideScale, this.dragState.handle);
      this.applyTextResizePreview(
        element,
        this.dragState.previewOriginalText ?? null,
        start,
        slideScale,
        sx,
        sy,
        anchor.x,
        anchor.y,
      );
    }
  }

  /**
   * Undo a live drag preview translate. Only needed when the commit does not
   * re-render (e.g. the shape did not actually move); a committed move replaces
   * the node entirely, so the preview node is already detached by then.
   */
  private restoreShapeDragPreview(element: SVGGElement | null, original: string | null): void {
    if (!element || !element.isConnected) return;
    if (original === null) {
      element.removeAttribute('transform');
    } else {
      element.setAttribute('transform', original);
    }
  }

  private restorePictureImagePreview(
    image: SVGImageElement | null | undefined,
    attrs: DragState['previewImageAttrs'],
  ): void {
    if (!image || !attrs || !image.isConnected) return;
    image.style.removeProperty('will-change');
    image.setAttribute('x', String(attrs.x));
    image.setAttribute('y', String(attrs.y));
    image.setAttribute('width', String(attrs.width));
    image.setAttribute('height', String(attrs.height));
    if (attrs.transform) {
      image.setAttribute('transform', attrs.transform);
    } else {
      image.removeAttribute('transform');
    }
  }

  private computeDragOverlayBox(
    event: PointerEvent,
  ): { left: number; top: number; width: number; height: number } | null {
    if (!this.dragState) return null;

    const dx = event.clientX - this.dragState.startClientX;
    const dy = event.clientY - this.dragState.startClientY;
    const box = { ...this.dragState.startBox };

    if (this.dragState.mode === 'move') {
      box.left += dx;
      box.top += dy;
    } else {
      if (this.dragState.handle?.includes('w')) {
        box.left += dx;
        box.width -= dx;
      }
      if (this.dragState.handle?.includes('e')) {
        box.width += dx;
      }
      if (this.dragState.handle?.includes('n')) {
        box.top += dy;
        box.height -= dy;
      }
      if (this.dragState.handle?.includes('s')) {
        box.height += dy;
      }
      if (box.width < 0) {
        box.left += box.width;
        box.width = -box.width;
      }
      if (box.height < 0) {
        box.top += box.height;
        box.height = -box.height;
      }
      box.width = Math.max(12, box.width);
      box.height = Math.max(12, box.height);
    }

    return box;
  }

  private applySelectionOverlayBox(box: {
    left: number;
    top: number;
    width: number;
    height: number;
  }): void {
    if (!this.selectionOverlay) return;
    this.selectionOverlay.style.removeProperty('transform');
    this.selectionOverlay.style.removeProperty('transform-origin');
    this.selectionOverlay.setCssProps({
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: `${box.width}px`,
      height: `${box.height}px`,
    });
  }

  /** Keep the dragged resize dot on the opposite edge after it crosses. */
  private setResizeCrossedHandleState(horizontal: boolean, vertical: boolean): void {
    if (!this.selectionOverlay) return;
    this.selectionOverlay.toggleClass('native-powerpoint-resize-crossed-horizontal', horizontal);
    this.selectionOverlay.toggleClass('native-powerpoint-resize-crossed-vertical', vertical);
  }

  /**
   * Map the on-screen image AABB (overlay box) to OOXML transform. Screen handle
   * directions swap width/height against cx/cy when the picture is rotated 90°/270°.
   */
  private computePictureTransformFromOverlay(
    overlayBox: { left: number; top: number; width: number; height: number },
  ): ShapeTransform | null {
    if (!this.dragState || !this.engine) return null;

    const start = this.dragState.startTransform;
    const startBox = this.dragState.startBox;
    const paneScaleX = this.dragState.paneEmuScaleX;
    const paneScaleY = this.dragState.paneEmuScaleY;
    if (!paneScaleX || !paneScaleY) return null;

    const minSize = this.engine.pxToEmu(12);
    const sx = startBox.width > 0 ? overlayBox.width / startBox.width : 1;
    const sy = startBox.height > 0 ? overlayBox.height / startBox.height : 1;
    const rotDeg = this.engine.ooxmlToDegrees(start.rot);
    const dimensions = this.solveRotatedRectDimensions(
      overlayBox.width / paneScaleX,
      overlayBox.height / paneScaleY,
      rotDeg,
      start.cx,
      start.cy,
      sx,
      sy,
    );
    const next = cloneTransform(start);

    next.cx = Math.max(minSize, dimensions.width);
    next.cy = Math.max(minSize, dimensions.height);

    const startCenterX = start.x + start.cx / 2;
    const startCenterY = start.y + start.cy / 2;
    const centerDeltaX =
      (overlayBox.left + overlayBox.width / 2 - (startBox.left + startBox.width / 2)) / paneScaleX;
    const centerDeltaY =
      (overlayBox.top + overlayBox.height / 2 - (startBox.top + startBox.height / 2)) / paneScaleY;
    next.x = startCenterX + centerDeltaX - next.cx / 2;
    next.y = startCenterY + centerDeltaY - next.cy / 2;

    return next;
  }

  private paneDeltaToUserDelta(dPaneX: number, dPaneY: number): { x: number; y: number } | null {
    if (!this.svgEl) return null;
    const ctm = this.svgEl.getScreenCTM();
    if (!ctm || ctm.a === 0 || ctm.d === 0) return null;
    return { x: dPaneX / ctm.a, y: dPaneY / ctm.d };
  }

  private paneSizeToUserSize(width: number, height: number): { width: number; height: number } | null {
    if (!this.svgEl) return null;
    const ctm = this.svgEl.getScreenCTM();
    if (!ctm || ctm.a === 0 || ctm.d === 0) return null;
    return { width: width / Math.abs(ctm.a), height: height / Math.abs(ctm.d) };
  }

  /** Scale the rendered {@link SVGImageElement} to match the overlay during picture resize. */
  private updatePictureImagePreview(
    overlayBox: { left: number; top: number; width: number; height: number },
  ): void {
    const dragState = this.dragState;
    if (!dragState?.previewImageElement || !dragState.previewImageAttrs) return;

    const orig = dragState.previewImageAttrs;
    const startBox = dragState.startBox;
    const sx = startBox.width > 0 ? overlayBox.width / startBox.width : 1;
    const sy = startBox.height > 0 ? overlayBox.height / startBox.height : 1;
    const startCenterX = startBox.left + startBox.width / 2;
    const startCenterY = startBox.top + startBox.height / 2;
    const overlayCenterX = overlayBox.left + overlayBox.width / 2;
    const overlayCenterY = overlayBox.top + overlayBox.height / 2;
    const userDelta = this.paneDeltaToUserDelta(overlayCenterX - startCenterX, overlayCenterY - startCenterY);
    const userSize = this.paneSizeToUserSize(overlayBox.width, overlayBox.height);
    if (!userDelta || !userSize) return;

    const rotate = this.parseSvgRotate(orig.transform);
    const rotationDegrees = rotate?.degrees ?? (this.engine ? this.engine.ooxmlToDegrees(dragState.startTransform.rot) : 0);
    const dimensions = this.solveRotatedRectDimensions(
      userSize.width,
      userSize.height,
      rotationDegrees,
      orig.width,
      orig.height,
      sx,
      sy,
    );
    const image = dragState.previewImageElement;
    const picture = image.closest('g[data-ooxml-shape-type="picture"]');
    if (picture?.getAttribute('data-ooxml-blip-stretch') === '1') {
      image.setAttribute('preserveAspectRatio', 'none');
    }
    const width = Math.max(1, dimensions.width);
    const height = Math.max(1, dimensions.height);
    const centerX = orig.x + orig.width / 2 + userDelta.x;
    const centerY = orig.y + orig.height / 2 + userDelta.y;
    const x = centerX - width / 2;
    const y = centerY - height / 2;

    image.setAttribute('x', this.formatSvgNumber(x));
    image.setAttribute('y', this.formatSvgNumber(y));
    image.setAttribute('width', this.formatSvgNumber(width));
    image.setAttribute('height', this.formatSvgNumber(height));

    if (orig.transform && rotate) {
      const originalCenterX = orig.x + orig.width / 2;
      const originalCenterY = orig.y + orig.height / 2;
      const rotateCenterX = centerX + (rotate.centerX === null ? 0 : rotate.centerX - originalCenterX);
      const rotateCenterY = centerY + (rotate.centerY === null ? 0 : rotate.centerY - originalCenterY);
      const nextRotate =
        `rotate(${this.formatSvgNumber(rotate.degrees)},${this.formatSvgNumber(rotateCenterX)},${this.formatSvgNumber(rotateCenterY)})`;
      image.setAttribute('transform', orig.transform.replace(rotate.raw, nextRotate));
    }
  }

  /**
   * Convert a crossing resize into a positive OOXML frame plus flip intent.
   * DrawingML cannot store negative extents, so the shape moves to the other
   * side of its fixed edge and is flipped only when the drag commits.
   */
  private getResizeTransform(
    dx: number,
    dy: number,
  ): { transform: ShapeTransform; crossedHorizontal: boolean; crossedVertical: boolean } | null {
    if (!this.dragState || !this.engine) return null;

    const start = this.dragState.startTransform;
    const next = cloneTransform(start);
    const minSize = this.engine.pxToEmu(12);
    let crossedHorizontal = false;
    let crossedVertical = false;

    if (this.dragState.handle?.includes('w')) {
      const width = start.cx - dx;
      if (width >= minSize) {
        next.x = start.x + dx;
        next.cx = width;
      } else if (width <= -minSize) {
        next.x = start.x + start.cx;
        next.cx = -width;
        crossedHorizontal = true;
      } else {
        next.x = start.x + start.cx - minSize;
        next.cx = minSize;
      }
    }
    if (this.dragState.handle?.includes('e')) {
      const width = start.cx + dx;
      if (width >= minSize) {
        next.cx = width;
      } else if (width <= -minSize) {
        next.x = start.x + width;
        next.cx = -width;
        crossedHorizontal = true;
      } else {
        next.cx = minSize;
      }
    }
    if (this.dragState.handle?.includes('n')) {
      const height = start.cy - dy;
      if (height >= minSize) {
        next.y = start.y + dy;
        next.cy = height;
      } else if (height <= -minSize) {
        next.y = start.y + start.cy;
        next.cy = -height;
        crossedVertical = true;
      } else {
        next.y = start.y + start.cy - minSize;
        next.cy = minSize;
      }
    }
    if (this.dragState.handle?.includes('s')) {
      const height = start.cy + dy;
      if (height >= minSize) {
        next.cy = height;
      } else if (height <= -minSize) {
        next.y = start.y + height;
        next.cy = -height;
        crossedVertical = true;
      } else {
        next.cy = minSize;
      }
    }

    return { transform: next, crossedHorizontal, crossedVertical };
  }

  private handleDragMove = (event: PointerEvent): void => {
    if (this.marquee) {
      this.updateMarquee(event);
      return;
    }
    if (this.groupDrag) {
      this.updateGroupDrag(event);
      return;
    }
    if (!this.dragState || !this.engine || !this.svgEl) return;
    if (event.pointerId !== this.dragState.pointerId) return;

    if (this.dragState.mode === 'rotate') {
      this.updateRotateDrag(event);
      return;
    }

    const point = this.getSvgPoint(event);
    if (!point) return;

    const scale = this.engine.getSlideScale(this.svgEl);
    const dx = (point.x - this.dragState.startPoint.x) * scale;
    const dy = (point.y - this.dragState.startPoint.y) * scale;
    const next = cloneTransform(this.dragState.startTransform);

    if (this.dragState.mode === 'move') {
      next.x += dx;
      next.y += dy;
      const snap = this.snapController.computeSnap(
        { x: next.x, y: next.y, cx: next.cx, cy: next.cy },
        new Set(this.selectedShapeIndex === null ? [] : [this.selectedShapeIndex])
      );
      next.x += snap.dx;
      next.y += snap.dy;
      this.snapController.updateSnapGuides(snap.guideX, snap.guideY);
      this.dragState.latestTransform = next;
      this.selectedTransform = cloneTransform(next);
      this.updateShapeTransformPreview(next);
      this.positionOverlayDuringMove(next);
      return;
    }

    const resized = this.getResizeTransform(dx, dy);
    if (!resized) return;
    this.dragState.crossedHorizontal = resized.crossedHorizontal;
    this.dragState.crossedVertical = resized.crossedVertical;
    this.setResizeCrossedHandleState(resized.crossedHorizontal, resized.crossedVertical);

    const isPicture = this.isPictureShape(this.getSelectedShapeElement());
    if (isPicture) {
      const overlayBox = this.computeDragOverlayBox(event);
      if (!overlayBox) return;
      const nextFromOverlay = this.computePictureTransformFromOverlay(overlayBox);
      if (!nextFromOverlay) return;
      this.dragState.latestTransform = nextFromOverlay;
      this.selectedTransform = cloneTransform(nextFromOverlay);
      this.updatePictureImagePreview(overlayBox);
      const selected = this.getSelectedShapeElement();
      const liveBox = selected ? this.getPictureSelectionBox(selected) : null;
      this.applySelectionOverlayBox(liveBox ?? overlayBox);
      if (selected) {
        this.logSelectionOverlayLayout(
          'drag',
          selected,
          nextFromOverlay,
          liveBox ?? overlayBox,
          'picture-image-bounds-live',
        );
      }
      return;
    }

    this.dragState.latestTransform = resized.transform;
    this.selectedTransform = cloneTransform(resized.transform);
    this.updateShapeTransformPreview(resized.transform);
    if (this.shapeHasRotation(resized.transform)) {
      this.applySelectionOverlayLayout(resized.transform);
    } else {
      this.updateSelectionOverlayDuringDrag(event);
    }
  };

  private updateRotateDrag(event: PointerEvent): void {
    if (!this.engine || !this.dragState) return;

    const centerX = this.dragState.centerClientX ?? 0;
    const centerY = this.dragState.centerClientY ?? 0;
    const angle = Math.atan2(event.clientY - centerY, event.clientX - centerX);
    const deltaDegrees = this.advanceContinuousRotation(this.dragState, angle);
    let degrees = this.engine.ooxmlToDegrees(this.dragState.startTransform.rot) + deltaDegrees;
    if (event.shiftKey) degrees = Math.round(degrees / 15) * 15;
    const rotationSnapTarget = this.getCardinalRotationSnap(degrees);
    degrees = rotationSnapTarget ?? degrees;

    const next = cloneTransform(this.dragState.startTransform);
    next.rot = this.engine.degreesToOoxml(this.normalizeDegrees(degrees));
    this.dragState.latestTransform = next;
    this.dragState.rotationSnapTarget = rotationSnapTarget;
    this.selectedTransform = cloneTransform(next);
    this.updateShapeTransformPreview(next);
    if (this.selectionOverlay) {
      this.selectionOverlay.setCssProps({ transform: `rotate(${degrees}deg)` });
    }
  }

  private handleDragEnd = (event: PointerEvent): void => {
    if (this.marquee) {
      this.finishMarquee(event);
      return;
    }
    if (this.groupDrag) {
      this.finishGroupDrag(event);
      return;
    }
    if (!this.dragState || event.pointerId !== this.dragState.pointerId) return;

    this.snapController.endDrag();
    const transform = cloneTransform(this.dragState.latestTransform);
    const startTransform = cloneTransform(this.dragState.startTransform);
    const moved = !transformsMatch(startTransform, transform);
    const mode = this.dragState.mode;
    const handle = this.dragState.handle;
    const rotationSnapTarget = this.dragState.rotationSnapTarget ?? null;
    const totalRotationDegrees = mode === 'rotate'
      ? this.dragState.accumulatedRotationDegrees ?? 0
      : null;
    const flipAxes = {
      horizontal: mode === 'resize' && (this.dragState.crossedHorizontal ?? false),
      vertical: mode === 'resize' && (this.dragState.crossedVertical ?? false),
    };
    const previewElement = this.dragState.previewElement ?? null;
    const previewOriginalTransform = this.dragState.previewOriginalTransform ?? null;
    const previewOriginalText = this.dragState.previewOriginalText ?? null;
    const previewImageElement = this.dragState.previewImageElement;
    const previewImageAttrs = this.dragState.previewImageAttrs;
    const freezeTextDuringResize = this.dragState.freezeShapeDuringResize ?? false;
    if (previewElement) {
      previewElement.classList.remove('native-powerpoint-shape-drag-preview');
    }
    if (previewImageElement) {
      previewImageElement.classList.remove('native-powerpoint-shape-drag-preview');
    }
    this.setResizeCrossedHandleState(false, false);
    this.dragState = null;
    if (moved) {
      this.suppressNextClick = true;
    }
    debugLog('selection', 'PowerPoint selection drag preview ended', {
      op: 'drag-preview-end',
      slide: this.currentSlide,
      shapeIndexes: this.selectedShapeIndex === null ? [] : [this.selectedShapeIndex],
      mode,
      handle,
      startTransform,
      finalTransform: transform,
      preview: 'shape-group',
      freezeTextDuringResize,
      rotationSnapTarget,
      totalRotationDegrees,
      flippedHorizontal: flipAxes.horizontal,
      flippedVertical: flipAxes.vertical,
      moved,
    });
    this.updateInspectorValues();
    void this.commitTransform(transform, flipAxes).finally(() => {
      this.restoreShapeDragPreview(previewElement, previewOriginalTransform);
      if (previewElement) this.restoreShapeTextPreview(previewElement, previewOriginalText);
      this.restorePictureImagePreview(previewImageElement, previewImageAttrs);
    });
  };

  private async commitTransform(
    transform: ShapeTransform,
    flipAxes: { horizontal: boolean; vertical: boolean } = { horizontal: false, vertical: false },
  ): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('edit object')) return;

    try {
      const startedAt = performance.now();
      const selected = this.getSelectedShapeElement();
      const shapeIndex = selected ? getShapeIndex(selected) : this.selectedShapeIndex;
      if (shapeIndex === null) return;
      if (!isEditableShapeIndex(shapeIndex)) {
        pptNotice('powerpoint:notice.objectNotEditable');
        return;
      }
      const before = selected ? cloneTransform(this.engine.getShapeTransform(selected)) : null;
      const shouldFlip = flipAxes.horizontal || flipAxes.vertical;
      if (before && transformsMatch(before, transform) && !shouldFlip) return;

      // Object moves/resizes record a before/after delta rather than exporting
      // the whole deck, which is what made dragging lag on large presentations.
      // A crossed handle also changes flipH/flipV, so its undo record must retain
      // the complete slide XML rather than only the frame transform.
      let history: HistoryEntry = shouldFlip
        ? this.historyController.captureSlideXml(this.currentSlide, 'Edit layout')
        : before
          ? this.historyController.captureTransform(
            this.currentSlide,
            [{ shapeIndex, before, after: cloneTransform(transform) }],
            'Edit layout'
          )
          : await this.captureHistoryEntry('Edit layout');
      const transformStarted = performance.now();
      await this.engine.updateShapeTransform(this.currentSlide, shapeIndex, transform);
      if (flipAxes.horizontal) await this.engine.flipShape(this.currentSlide, shapeIndex, 'horizontal');
      if (flipAxes.vertical) await this.engine.flipShape(this.currentSlide, shapeIndex, 'vertical');
      const transformMs = Math.round(performance.now() - transformStarted);
      this.selectedTransform = cloneTransform(transform);
      if (history.kind === 'slideXml') {
        history = this.historyController.completeSlideXml(history);
      }
      this.recordHistoryEntry(history);
      this.markDirty();
      const renderStarted = performance.now();
      const rendered = await this.renderEditedShape(shapeIndex);
      const renderMs = Math.round(performance.now() - renderStarted);
      if (rendered) this.scheduleThumbnailRefresh(this.currentSlide);
      debugLog('inspector', 'Committed PowerPoint object transform', {
        slide: this.currentSlide,
        shapeIndex,
        x: transform.x,
        y: transform.y,
        width: transform.cx,
        height: transform.cy,
        rotation: transform.rot,
        flippedHorizontal: flipAxes.horizontal,
        flippedVertical: flipAxes.vertical,
        transformMs,
        renderMs,
        ms: Math.round(performance.now() - startedAt)
      });
    } catch (error) {
      errorLog('inspector', 'PowerPoint object transform failed', {
        slide: this.currentSlide,
        shapeIndex: this.selectedShapeIndex,
        error
      });
      pptNotice('powerpoint:notice.couldNotUpdateObject', { message: cleanError(error) });
    }
  }

  private getSvgPoint(event: MouseEvent | PointerEvent): PointerPoint | null {
    if (!this.svgEl) return null;

    const matrix = this.svgEl.getScreenCTM();
    if (!matrix) return null;

    const point = this.svgEl.createSVGPoint();
    point.x = event.clientX;
    point.y = event.clientY;
    const result = point.matrixTransform(matrix.inverse());
    return { x: result.x, y: result.y };
  }

  private updateSelectionOverlayDuringDrag(event: PointerEvent): void {
    const box = this.computeDragOverlayBox(event);
    if (!box) return;
    this.applySelectionOverlayBox(box);
  }

  private updateInspectorValues(): void {
    if (!this.engine || !this.selectedTransform) return;

    if (this.xInput) this.xInput.value = String(Math.round(this.engine.emuToPx(this.selectedTransform.x) * 100) / 100);
    if (this.yInput) this.yInput.value = String(Math.round(this.engine.emuToPx(this.selectedTransform.y) * 100) / 100);
    if (this.widthInput) this.widthInput.value = String(Math.round(this.engine.emuToPx(this.selectedTransform.cx) * 100) / 100);
    if (this.heightInput) this.heightInput.value = String(Math.round(this.engine.emuToPx(this.selectedTransform.cy) * 100) / 100);
    if (this.rotationInput) this.rotationInput.value = String(Math.round(this.engine.ooxmlToDegrees(this.selectedTransform.rot) * 100) / 100);
  }

  private updateSlideScale(): void {
    if (!this.canvasPane || !this.slideSurface || !this.svgEl) return;

    const size = getSvgIntrinsicSize(this.svgEl);
    if (!size) {
      this.updateSelectionOverlay();
      return;
    }

    const computedStyle = window.getComputedStyle(this.canvasPane);
    const horizontalPadding =
      (Number.parseFloat(computedStyle.paddingLeft) || 0) +
      (Number.parseFloat(computedStyle.paddingRight) || 0);
    const verticalPadding =
      (Number.parseFloat(computedStyle.paddingTop) || 0) +
      (Number.parseFloat(computedStyle.paddingBottom) || 0);
    const availableWidth = Math.max(1, this.canvasPane.clientWidth - horizontalPadding);
    const availableHeight = Math.max(1, this.canvasPane.clientHeight - verticalPadding);
    const fitScale = Math.min(1, availableWidth / size.width, availableHeight / size.height);
    const scale = Math.max(0.05, fitScale * this.zoomLevel);
    const width = Math.max(1, Math.floor(size.width * scale));
    const height = Math.max(1, Math.floor(size.height * scale));

    this.slideSurface.addClass('is-scaled');
    this.slideSurface.style.setProperty('--native-powerpoint-slide-width', `${width}px`);
    this.slideSurface.style.setProperty('--native-powerpoint-slide-height', `${height}px`);
    this.updateSelectionOverlay();
    this.refreshActiveInlineEditorGeometry();
  }

  private handleCanvasWheel(event: WheelEvent): void {
    if (!this.canvasPane || !this.slideSurface || !this.svgEl || !this.engine) return;
    if (!this.isActivePowerPointView()) return;

    event.preventDefault();
    event.stopPropagation();

    // Browsers synthesize Ctrl+wheel for trackpad pinch gestures. All other
    // wheel input, including two-finger trackpad scrolling, pans the canvas.
    if (!event.ctrlKey) {
      const deltaX = this.normalizeWheelDelta(event, event.deltaX);
      const deltaY = this.normalizeWheelDelta(event, event.deltaY);
      if (deltaX === 0 && deltaY === 0) return;

      this.canvasPane.scrollLeft += deltaX;
      this.canvasPane.scrollTop += deltaY;
      debugLog('view', 'Panned PowerPoint canvas', {
        deltaX,
        deltaY,
        scrollLeft: this.canvasPane.scrollLeft,
        scrollTop: this.canvasPane.scrollTop,
        scrollWidth: this.canvasPane.scrollWidth,
        scrollHeight: this.canvasPane.scrollHeight,
        clientWidth: this.canvasPane.clientWidth,
        clientHeight: this.canvasPane.clientHeight,
        source: 'wheel'
      });
      return;
    }

    const delta = this.normalizeWheelDelta(event);
    if (delta === 0) return;

    const nextZoom = this.zoomLevel * Math.pow(2, -delta / 600);
    this.setZoom(nextZoom, { clientX: event.clientX, clientY: event.clientY });
  }

  private normalizeWheelDelta(event: WheelEvent, value = event.deltaY): number {
    if (event.deltaMode === WheelEvent.DOM_DELTA_LINE) {
      return value * 16;
    }

    if (event.deltaMode === WheelEvent.DOM_DELTA_PAGE) {
      return value * Math.max(1, this.canvasPane?.clientHeight ?? 800);
    }

    return value;
  }

  private setZoom(value: number, anchor?: { clientX: number; clientY: number }): void {
    const nextZoom = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, Math.round(value * 1000) / 1000));
    if (nextZoom === this.zoomLevel) return;

    const previousZoom = this.zoomLevel;
    const anchorState = anchor ? this.captureZoomAnchor(anchor) : null;
    this.zoomLevel = nextZoom;
    debugLog('view', 'Changed PowerPoint zoom', {
      previousZoom,
      zoom: nextZoom,
      source: anchor ? 'wheel' : 'toolbar'
    });
    this.updateZoomLabel();
    this.updateSlideScale();
    if (anchorState) {
      this.restoreZoomAnchor(anchorState);
    }
  }

  private captureZoomAnchor(anchor: { clientX: number; clientY: number }): {
    clientX: number;
    clientY: number;
    ratioX: number;
    ratioY: number;
  } | null {
    if (!this.slideSurface) return null;

    const rect = this.slideSurface.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) return null;

    return {
      clientX: anchor.clientX,
      clientY: anchor.clientY,
      ratioX: Math.max(0, Math.min(1, (anchor.clientX - rect.left) / rect.width)),
      ratioY: Math.max(0, Math.min(1, (anchor.clientY - rect.top) / rect.height))
    };
  }

  private restoreZoomAnchor(anchor: { clientX: number; clientY: number; ratioX: number; ratioY: number }): void {
    if (!this.canvasPane || !this.slideSurface) return;

    const rect = this.slideSurface.getBoundingClientRect();
    const nextX = rect.left + rect.width * anchor.ratioX;
    const nextY = rect.top + rect.height * anchor.ratioY;
    this.canvasPane.scrollLeft += nextX - anchor.clientX;
    this.canvasPane.scrollTop += nextY - anchor.clientY;
    this.updateSelectionOverlay();
    this.refreshActiveInlineEditorGeometry();
  }

  private refreshActiveInlineEditorGeometry(): void {
    if (!this.activeEditor || !this.activeEditorTarget) return;

    const box = this.getElementBox(this.activeEditorTarget);
    if (!box) return;

    this.positionTextRunEditor(this.activeEditor, box);
    this.updateInlineCaret(this.activeEditor, this.activeEditorTarget);
  }

  private markDirty(): void {
    this.session.applyEdit();
  }

  private clearAutosave(): void {
    this.session.clearAutosave();
  }

  private async preserveUnsavedChangesForTeardown(reason: string): Promise<boolean> {
    return this.session.preserveUnsavedChangesForTeardown(reason);
  }

  private resetLoadedPresentation(): void {
    this.filmstripRendered = false;
    this.filmstripRenderScheduled = false;
    this.session.reset();
    this.slideFilmstripController.dispose();
    this.removeActiveEditor();
    this.historyController.clear();
    this.engine = null;
		this.presentationWordCount = 0;
		this.presentationWordCountEditVersion = -1;
		this.onWordCountClear();
    this.loadedFile = null;
    this.sourcePackage = null;
    this.sourceBuffer = null;
    this.isViewOnly = false;
    this.viewOnlyReason = '';
    this.selectedShapeIndex = null;
    this.selectedTransform = null;
    this.hasShownGeneratedTextNotice = false;
    this.fontSubstitutions = [];
    this.findController.reset();
    this.svgEl = null;
    this.dragState = null;
    this.cancelMarquee();
  }

  private updateEditingAvailability(): void {
    const canEdit = this.canEdit();
    const disabledReason = this.viewOnlyReason || this.t('powerpoint:notice.openEditableFirst');

    for (const button of this.editButtons) {
      const baseTitle = button.dataset.baseTitle || button.getAttribute('aria-label') || this.tb('edit');
      button.disabled = !canEdit;
      button.toggleClass('is-disabled', !canEdit);
      button.setAttribute('aria-label', canEdit ? baseTitle : this.t('powerpoint:accessibility.editDisabled', { label: baseTitle, reason: disabledReason }));
      button.setAttribute('aria-disabled', String(!canEdit));
    }

    this.historyController.updateAvailability();
    this.updateObjectClipboardAvailability();
  }

  private updateObjectClipboardAvailability(): void {
    const hasSelection = this.getSelectedIndices().length > 0;
    const hasSingleSelection = this.selectedShapeIndex !== null;
    const canEdit = this.canEdit();
    this.updateObjectClipboardButton(this.copyButton, hasSelection);
    this.updateObjectClipboardButton(this.pasteButton, canEdit && Boolean(this.objectClipboard));
    this.updateObjectClipboardButton(this.duplicateButton, canEdit && hasSingleSelection);
    this.updateArrangeAvailability();
  }

  private updateObjectClipboardButton(button: HTMLButtonElement | null, enabled: boolean): void {
    if (!button) return;
    button.disabled = !enabled;
    button.toggleClass('is-disabled', !enabled);
    button.setAttribute('aria-disabled', String(!enabled));
  }

  private updateSlideCounter(): void {
    this.updateHeaderTitle();
    const count = this.engine?.slideCount || 0;
    this.slideCounterEl?.setText(this.t('powerpoint:present.slideCount', {
      current: count ? this.currentSlide + 1 : 0,
      total: count
    }));

    if (this.thumbnailContainer) {
      this.thumbnailContainer.querySelectorAll('.native-powerpoint-thumbnail').forEach((thumbnail, index) => {
        thumbnail.toggleClass('active', index === this.currentSlide);
      });
    }
  }
}
