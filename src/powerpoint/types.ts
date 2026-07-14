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
  /** Live-preview element (the dragged shape group) and its original transform. */
  previewElement?: SVGGElement | null;
  previewOriginalTransform?: string | null;
  /** Text boxes keep their rendered contents unchanged until resize commit. */
  freezeShapeDuringResize?: boolean;
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
  pointerId: number;
  startPoint: PointerPoint;
  startClientX: number;
  startClientY: number;
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
