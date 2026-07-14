import { isSVGTSpanElement } from '../domGuards';
import type { SvgRectLike } from './types';

export interface InlineLocalBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface InlineCaretGeometry {
  left: number;
  top: number;
  height: number;
}

export interface LeafCharEntry {
  span: SVGTextContentElement;
  count: number;
  start: number;
}

export interface LeafCharInfo {
  entries: LeafCharEntry[];
  total: number;
}

/**
 * Pure SVG text hit-testing / caret geometry.
 *
 * This logic depends only on the rendered SVG and the canvas pane element, so it
 * runs identically inside the Obsidian view and inside a headless-Chrome harness.
 * Keep it free of any Obsidian-specific globals (no `activeDocument`, no editor
 * state) so `scripts/smoke-selection-geometry.mjs` exercises the real shipped
 * code rather than a copy.
 */
export class InlineTextGeometry {
  constructor(private readonly getCanvasPane: () => HTMLElement | null) {}

  private get pane(): HTMLElement | null {
    return this.getCanvasPane();
  }

  getElementBox(element: Element): InlineLocalBox | null {
    const pane = this.pane;
    if (!pane) return null;

    const paneRect = pane.getBoundingClientRect();
    const shapeRect = element.getBoundingClientRect();
    return {
      left: shapeRect.left - paneRect.left + pane.scrollLeft,
      top: shapeRect.top - paneRect.top + pane.scrollTop,
      width: shapeRect.width,
      height: shapeRect.height
    };
  }

  getScreenFontSize(element: SVGTextElement | SVGTSpanElement): number {
    const style = window.getComputedStyle(element);
    const fontSize = Number.parseFloat(style.fontSize) || 16;
    const matrix = element.getScreenCTM();
    const scale = matrix ? Math.hypot(matrix.a, matrix.b) : 1;
    return Math.max(4, fontSize * scale);
  }

  getLeafTextSpans(element: SVGTextElement | SVGTSpanElement): SVGTextContentElement[] {
    // Include every innermost tspan (bullet markers, runs, etc.), not only
    // data-ooxml-run-idx nodes. Paragraph textContent also contains bullet
    // prefixes, so limiting to runs desynchronizes string indices from glyph
    // geometry and misplaces find/selection highlights.
    const leafTspans = Array.from(element.querySelectorAll('tspan'))
      .filter(isSVGTSpanElement)
      .filter((span) => !span.querySelector('tspan'));
    if (leafTspans.length > 0) {
      return leafTspans;
    }
    return [element];
  }

  isRunTextSpan(span: SVGTextContentElement): boolean {
    return span.hasAttribute('data-ooxml-run-idx');
  }

  /**
   * Editor-string length of a leaf tspan. Every offset that flows through this
   * module — inline-editor caret offsets, find's `indexOf` offsets, and the
   * per-render `data-ooxml-char-*` run stamps — indexes the run/leaf
   * *textContent* string. The per-leaf count must therefore be the string
   * length, not `getNumberOfChars()`: the glyph count drifts from the string at
   * ligatures/shaping and was the original offset-drift source. We keep
   * `getNumberOfChars()` solely as the valid bound for SVG glyph-index probes
   * (`getStartPositionOfChar` / `getEndPositionOfChar` / `getExtentOfChar`).
   */
  leafCharCount(span: SVGTextContentElement): number {
    return (span.textContent ?? '').length;
  }

  /** Rendered glyph count for `span` (the addressable range for per-char geometry). */
  getGlyphCount(span: SVGTextContentElement): number {
    try {
      const glyphs = span.getNumberOfChars();
      if (Number.isFinite(glyphs) && glyphs >= 0) return glyphs;
    } catch {
      // jsdom and detached nodes throw; fall through to the string length.
    }
    return (span.textContent ?? '').length;
  }

  /** Clamp a string index to a valid glyph index before probing SVG char geometry. */
  clampGlyphIndex(span: SVGTextContentElement, index: number): number {
    const glyphs = this.getGlyphCount(span);
    if (glyphs <= 0) return 0;
    return Math.max(0, Math.min(glyphs - 1, index));
  }

  /** Character counts for OOXML runs only — matches the inline editor / engine offsets. */
  getRunCharInfo(element: SVGTextElement | SVGTSpanElement): LeafCharInfo {
    const entries: LeafCharEntry[] = [];
    let total = 0;
    for (const span of this.getLeafTextSpans(element)) {
      if (!this.isRunTextSpan(span)) continue;
      const count = this.leafCharCount(span);
      if (count <= 0) continue;
      entries.push({ span, count, start: total });
      total += count;
    }
    if (entries.length > 0) {
      return { entries, total };
    }
    return this.getLeafCharInfo(element);
  }

  /** Map a geometry (all-leaf) character index to a run-only offset within `element`. */
  geometryIndexToRunOffset(element: SVGTextElement | SVGTSpanElement, geometryIndex: number): number {
    const clamped = Math.max(0, geometryIndex);
    let runOffset = 0;
    let leafOffset = 0;
    for (const span of this.getLeafTextSpans(element)) {
      const count = this.leafCharCount(span);
      if (count <= 0) continue;

      if (clamped <= leafOffset + count) {
        if (this.isRunTextSpan(span)) {
          runOffset += Math.max(0, clamped - leafOffset);
        }
        return runOffset;
      }

      if (this.isRunTextSpan(span)) {
        runOffset += count;
      }
      leafOffset += count;
    }
    return runOffset;
  }

  /** Map a run-only offset to a geometry (all-leaf) character index within `element`. */
  runOffsetToGeometryIndex(element: SVGTextElement | SVGTSpanElement, runOffset: number): number {
    const clamped = Math.max(0, runOffset);
    let runSeen = 0;
    let leafOffset = 0;
    for (const span of this.getLeafTextSpans(element)) {
      const count = this.leafCharCount(span);
      if (count <= 0) continue;

      if (this.isRunTextSpan(span)) {
        if (clamped <= runSeen + count) {
          return leafOffset + Math.max(0, clamped - runSeen);
        }
        runSeen += count;
      }
      leafOffset += count;
    }
    return leafOffset;
  }

  getLeafCharInfo(element: SVGTextElement | SVGTSpanElement): LeafCharInfo {
    const entries: LeafCharEntry[] = [];
    let total = 0;
    for (const span of this.getLeafTextSpans(element)) {
      const count = this.leafCharCount(span);
      if (count <= 0) continue;
      entries.push({ span, count, start: total });
      total += count;
    }
    return { entries, total };
  }

  private transformSvgRectToLocalBox(
    rect: SvgRectLike,
    matrix: DOMMatrix,
    paneRect: DOMRect
  ): InlineLocalBox | null {
    const scrollLeft = this.pane?.scrollLeft ?? 0;
    const scrollTop = this.pane?.scrollTop ?? 0;
    const points = [
      new DOMPoint(rect.x, rect.y),
      new DOMPoint(rect.x + rect.width, rect.y),
      new DOMPoint(rect.x, rect.y + rect.height),
      new DOMPoint(rect.x + rect.width, rect.y + rect.height)
    ].map((point) => point.matrixTransform(matrix));
    if (points.some((point) => !Number.isFinite(point.x) || !Number.isFinite(point.y))) {
      return null;
    }

    const xs = points.map((point) => point.x - paneRect.left + scrollLeft);
    const ys = points.map((point) => point.y - paneRect.top + scrollTop);
    const left = Math.min(...xs);
    const top = Math.min(...ys);
    return {
      left,
      top,
      width: Math.max(...xs) - left,
      height: Math.max(...ys) - top
    };
  }

  private getFallbackInlineCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    preferredHeight: number
  ): InlineCaretGeometry | null {
    const box = this.getElementBox(element);
    if (!box) return null;

    const height = Math.max(6, preferredHeight);
    return {
      left: box.left,
      top: box.top + Math.max(0, (box.height - height) / 2),
      height
    };
  }

  getSvgTextCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    offset: number,
    preferredHeight = this.getScreenFontSize(element) * 1.08
  ): InlineCaretGeometry | null {
    const text = element.textContent || '';
    if (!text) return this.getFallbackInlineCaretGeometry(element, preferredHeight);

    const paneRect = this.pane?.getBoundingClientRect();
    if (!paneRect) return null;

    const { entries, total } = this.getLeafCharInfo(element);
    if (entries.length === 0 || total <= 0) {
      return this.getFallbackInlineCaretGeometry(element, preferredHeight);
    }

    const normalizedOffset = Math.max(0, Math.min(total, offset));
    const useStart = normalizedOffset <= 0;
    const globalCharIndex = useStart ? 0 : normalizedOffset - 1;

    let entry = entries[0];
    for (const candidate of entries) {
      if (globalCharIndex >= candidate.start && globalCharIndex < candidate.start + candidate.count) {
        entry = candidate;
        break;
      }
      entry = candidate;
    }
    if (!entry) return this.getFallbackInlineCaretGeometry(element, preferredHeight);

    // `entry.count` is the string length; the SVG per-char probes index glyphs,
    // so clamp the string offset into the rendered glyph range first.
    const glyphIndex = this.clampGlyphIndex(entry.span, globalCharIndex - entry.start);
    const matrix = entry.span.getScreenCTM();
    if (!matrix) return null;

    let position: DOMPoint;
    let extent: SvgRectLike | null = null;
    try {
      position = useStart
        ? entry.span.getStartPositionOfChar(glyphIndex)
        : entry.span.getEndPositionOfChar(glyphIndex);
      extent = entry.span.getExtentOfChar(glyphIndex);
    } catch {
      return null;
    }

    const point = new DOMPoint(position.x, position.y).matrixTransform(matrix);
    const localLeft = point.x - paneRect.left + (this.pane?.scrollLeft ?? 0);
    const fallbackBox = this.getElementBox(element);
    let top = fallbackBox ? fallbackBox.top + Math.max(0, (fallbackBox.height - preferredHeight) / 2) : 0;
    let height = Math.max(6, preferredHeight);

    if (extent) {
      const bounds = this.transformSvgRectToLocalBox(extent, matrix, paneRect);
      if (bounds && bounds.height > 0) {
        // Snap to the real glyph row so the caret fills the actual line height
        // and aligns vertically to the text, instead of floating at the click Y.
        height = Math.max(6, bounds.height);
        top = bounds.top;
      }
    }

    return { left: localLeft, top, height };
  }

  getInlineTextOffsetFromSvgGeometry(
    element: SVGTextElement | SVGTSpanElement,
    localClientX: number,
    localClientY: number,
    textLength: number
  ): number | null {
    // Wrapped lines store glyphs in child tspans; the container's own
    // getNumberOfChars() under-counts (e.g. 20 vs 24) and caps selection
    // before the true line end. Leaf totals match getSvgTextCaretGeometry.
    const { total: leafTotal } = this.getLeafCharInfo(element);
    if (leafTotal <= 0) return null;

    const maxOffset = Math.min(textLength, leafTotal);
    const rowOffsets: number[] = [];
    for (let offset = 0; offset < maxOffset; offset++) {
      const geometry = this.getSvgTextCaretGeometry(element, offset);
      if (!geometry) continue;
      const centerY = geometry.top + geometry.height / 2;
      const previousOffset = rowOffsets.at(-1);
      const previous = previousOffset !== undefined
        ? this.getSvgTextCaretGeometry(element, previousOffset)
        : null;
      const previousY = previous ? previous.top + previous.height / 2 : centerY;
      if (rowOffsets.length === 0 || Math.abs(centerY - previousY) > Math.max(4, geometry.height * 0.45)) {
        rowOffsets.push(offset);
      }
    }

    let rowStart = 0;
    let rowEnd = maxOffset;
    if (rowOffsets.length > 1) {
      let bestRowIndex = 0;
      let bestRowDistance = Number.POSITIVE_INFINITY;
      for (let rowIndex = 0; rowIndex < rowOffsets.length; rowIndex++) {
        const rowOffset = rowOffsets[rowIndex] ?? 0;
        const geometry = this.getSvgTextCaretGeometry(element, rowOffset);
        if (!geometry) continue;
        const centerY = geometry.top + geometry.height / 2;
        const rowDistance = Math.abs(localClientY - centerY);
        if (rowDistance < bestRowDistance) {
          bestRowDistance = rowDistance;
          bestRowIndex = rowIndex;
        }
      }
      rowStart = rowOffsets[bestRowIndex] ?? 0;
      rowEnd = rowOffsets[bestRowIndex + 1] ?? maxOffset;
    }

    let bestOffset = rowStart;
    let bestDistance = Number.POSITIVE_INFINITY;
    for (let offset = rowStart; offset <= rowEnd; offset++) {
      const geometry = this.getSvgTextCaretGeometry(element, offset);
      if (!geometry) continue;

      const centerY = geometry.top + geometry.height / 2;
      const dx = localClientX - geometry.left;
      const dy = localClientY - centerY;
      const distance = dx * dx + dy * dy * 2.25;
      if (distance < bestDistance) {
        bestDistance = distance;
        bestOffset = offset;
      }
    }

    if (!Number.isFinite(bestDistance)) return null;

    // Click to the right of the last glyph on this row → caret after the line.
    const endCaret = this.getSvgTextCaretGeometry(element, rowEnd);
    if (endCaret && localClientX >= endCaret.left - 1) {
      return rowEnd;
    }

    return bestOffset;
  }

  getInlineTextOffsetAtClientPointForElement(
    element: SVGTextElement | SVGTSpanElement,
    clientX: number,
    clientY: number | undefined,
    box: InlineLocalBox
  ): number {
    const text = element.textContent ?? '';
    if (text.length === 0) return 0;

    const pane = this.pane;
    const paneRect = pane?.getBoundingClientRect();
    const localClientX = paneRect
      ? clientX - paneRect.left + (pane?.scrollLeft ?? 0)
      : clientX;
    const localClientY = paneRect && clientY !== undefined
      ? clientY - paneRect.top + (pane?.scrollTop ?? 0)
      : box.top + box.height / 2;

    const geometryOffset = this.getInlineTextOffsetFromSvgGeometry(element, localClientX, localClientY, text.length);
    return geometryOffset ?? Math.max(0, Math.min(text.length, Math.round(text.length * ((localClientX - box.left) / Math.max(1, box.width)))));
  }

  snapWrappedRunLocalToLineEnd(
    container: SVGTextElement | SVGTSpanElement,
    runLocal: number,
    runTotal: number,
    localClientX: number
  ): number {
    if (runLocal >= runTotal || runTotal <= 0) return runLocal;

    const geometryEnd = this.runOffsetToGeometryIndex(container, runTotal);
    const endCaret = this.getSvgTextCaretGeometry(container, geometryEnd);
    if (endCaret && localClientX >= endCaret.left - 1) {
      return runTotal;
    }

    if (runLocal < runTotal - 1) return runLocal;

    const lastGeometry = this.runOffsetToGeometryIndex(container, runTotal - 1);
    const lastStart = this.getSvgTextCaretGeometry(container, lastGeometry);
    const lastEnd = this.getSvgTextCaretGeometry(container, lastGeometry + 1);
    if (lastStart && lastEnd && localClientX >= (lastStart.left + lastEnd.left) / 2) {
      return runTotal;
    }

    return runLocal;
  }
}
