// Internal type and interface declarations for the PowerPoint view. These are
// type-only (no runtime code) and were extracted from NativePowerPointView.ts to
// shrink that file and make the view's data shapes easy to find.

import type {
  GeneratedTextEdit,
  ParagraphAlignment,
  ParagraphTextRange,
  RunTarget
} from '../PresentationEngine';
import type { ShapeTransform } from 'pptx-svg';

export type SaveState = 'idle' | 'dirty' | 'saving' | 'saved' | 'failed' | 'recovered' | 'view-only';
export type HandleName = 'nw' | 'ne' | 'sw' | 'se' | 'n' | 'e' | 's' | 'w';
export type SvgSecurityDecision = 'compatibility' | 'yolo' | null;

export interface PointerPoint {
  x: number;
  y: number;
}

export interface DragState {
  mode: 'move' | 'resize' | 'rotate';
  handle?: HandleName;
  pointerId: number;
  startPoint: PointerPoint;
  startClientX: number;
  startClientY: number;
  startBox: { left: number; top: number; width: number; height: number };
  startTransform: ShapeTransform;
  latestTransform: ShapeTransform;
  centerClientX?: number;
  centerClientY?: number;
  startAngle?: number;
  /** Most recent pointer angle, used to unwrap rotation across the ±180° seam. */
  lastAngle?: number;
  /** Unwrapped pointer rotation, preserving complete clockwise/counterclockwise turns. */
  accumulatedRotationDegrees?: number;
  /** Cardinal angle chosen for the live rotation preview, if any. */
  rotationSnapTarget?: number | null;
  /** Live-preview element (the dragged shape group) and its original transform. */
  previewElement?: SVGGElement | null;
  previewOriginalTransform?: string | null;
  /** Text glyphs stay unscaled while the temporary resize preview reflows them. */
  freezeShapeDuringResize?: boolean;
  /** Original text subtree restored when the temporary resize preview ends. */
  previewOriginalText?: SVGTextElement | null;
  /** The resize handle has crossed the opposite horizontal edge. */
  crossedHorizontal?: boolean;
  /** The resize handle has crossed the opposite vertical edge. */
  crossedVertical?: boolean;
  /** Pane pixels per EMU; cached at drag start for overlay positioning. */
  paneEmuScaleX?: number;
  paneEmuScaleY?: number;
  /** Live picture resize preview: inner {@link SVGImageElement} attrs at drag start. */
  previewImageElement?: SVGImageElement | null;
  previewImageAttrs?: {
    x: number;
    y: number;
    width: number;
    height: number;
    transform: string | null;
  } | null;
  /** Crop clip rect (`clipPath > rect`) attrs at drag start, when the picture is cropped. */
  previewClipRectElement?: SVGRectElement | null;
  previewClipAttrs?: {
    x: number;
    y: number;
    width: number;
    height: number;
  } | null;
}

export interface MarqueeState {
  pointerId: number;
  startClientX: number;
  startClientY: number;
  additive: boolean;
  base: number[];
  moved: boolean;
}

export interface GroupDragState {
  mode: DragState['mode'];
  handle?: HandleName;
  pointerId: number;
  startPoint: PointerPoint;
  startClientX: number;
  startClientY: number;
  /** Rendered union box used to position the shared group outline. */
  startBox: { left: number; top: number; width: number; height: number };
  /** Union of the selected OOXML frames at drag start. */
  startBounds: ShapeTransform;
  /** Current group frame used to keep the shared outline in sync. */
  latestBounds: ShapeTransform;
  centerClientX?: number;
  centerClientY?: number;
  startAngle?: number;
  /** Most recent pointer angle, used to unwrap rotation across the ±180° seam. */
  lastAngle?: number;
  /** Unwrapped pointer rotation, preserving complete clockwise/counterclockwise turns. */
  accumulatedRotationDegrees?: number;
  /** The group resize handle has crossed the opposite horizontal edge. */
  crossedHorizontal?: boolean;
  /** The group resize handle has crossed the opposite vertical edge. */
  crossedVertical?: boolean;
  /** Original SVG transforms for every live group move/resize preview shape. */
  previewOriginalTransforms?: Map<number, string | null>;
  /** Original text subtrees restored after each reflow pass and when the drag ends. */
  previewOriginalText?: Map<number, SVGTextElement>;
  /** Current preview totals, emitted once when the group drag ends. */
  previewObjectCount?: number;
  /** Text frames that received the non-stretching position compensation. */
  previewTextTransformCount?: number;
  previewTextReflowCount?: number;
  /** First temporary reflow failure, emitted once instead of per pointermove. */
  previewTextReflowError?: string | null;
  rotationSnapTarget?: number | null;
  start: Map<number, ShapeTransform>;
  latest: Map<number, ShapeTransform>;
  moved: boolean;
}

export interface PowerPointFindMatch {
  slideIndex: number;
  shapeIndex: number | null;
  text: string;
}

export interface SlideSize {
  width: number;
  height: number;
}

export interface InlineCaretRow {
  top: number;
  height: number;
  centerRatio: number;
}

export interface SvgRectLike {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface SvgInlineCaretGeometry {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  strokeWidth: number;
}

export interface SvgInlineSelectionBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface InlineSelectionDrag {
  editor: HTMLTextAreaElement;
  element: SVGTextElement | SVGTSpanElement;
  anchorOffset: number;
  pointerId: number;
  startClientX: number;
  startClientY: number;
  isSelecting: boolean;
  pendingFrame: number | null;
  pendingClientX: number;
  pendingClientY: number;
  lastLogKey?: string | null;
  lastLogAt?: number;
  cleanup: () => void;
}

export interface InlineCaretPlacement {
  editor: HTMLTextAreaElement;
  element: SVGTextElement | SVGTSpanElement;
  offset: number;
  timestamp: number;
}

export interface CanvasScrollPosition {
  left: number;
  top: number;
}

export interface HistorySnapshotEntry {
  kind: 'snapshot';
  buffer: ArrayBuffer;
  currentSlide: number;
  label: string;
}

export interface HistoryTransformChange {
  shapeIndex: number;
  before: ShapeTransform;
  after: ShapeTransform;
}

/**
 * Lightweight undo record for object move/resize/rotate. Stores only the
 * before/after transforms so the drag-commit path never has to export the whole
 * deck (which is the dominant source of lag on large presentations).
 */
export interface HistoryTransformEntry {
  kind: 'transform';
  slideIndex: number;
  changes: HistoryTransformChange[];
  currentSlide: number;
  label: string;
}

/**
 * Lightweight undo record for edits that only mutate one slide's XML, such as
 * text formatting. Avoids exporting the whole deck before every toolbar click.
 */
export interface HistorySlideXmlEntry {
  kind: 'slideXml';
  slideIndex: number;
  beforeXml: string;
  afterXml: string;
  currentSlide: number;
  label: string;
}

export type HistoryEntry = HistorySnapshotEntry | HistoryTransformEntry | HistorySlideXmlEntry;

export interface ShapeTextEditTarget {
  kind: 'shape-paragraph';
  shapeIndex: number;
  paragraphIndex: number;
  runIndex: number;
  text: string;
  element: SVGTextElement | SVGTSpanElement;
  runElements: SVGTSpanElement[];
}

export interface GeneratedTextEditTarget extends GeneratedTextEdit {
  shapeIndex: number;
  text: string;
  element: SVGTextElement;
}

export type TextEditTarget = GeneratedTextEditTarget | ShapeTextEditTarget;

export interface TextToolbarControls {
  bold: HTMLButtonElement;
  italic: HTMLButtonElement;
  underline: HTMLButtonElement;
  fontSizeInput: HTMLInputElement;
  textColorBar: HTMLElement;
  highlightBar: HTMLElement;
  alignButtons: Record<ParagraphAlignment, HTMLButtonElement>;
}

export interface TextStyleContext {
  shapeIndex: number;
  run: RunTarget | null;
  anchor: { left: number; top: number; width: number; height: number };
}

export interface InlineRangeSelection {
  shapeIndex: number;
  ranges: ParagraphTextRange[];
}

export interface ToolbarFormattingSnapshot extends TextStyleContext {
  ranges: ParagraphTextRange[] | null;
}

export type DistributeAxis = 'horizontal' | 'vertical';

export type MenuDropdownEntry =
  | 'separator'
  | { label: string; icon?: string; onClick: () => void; disabled?: boolean };
