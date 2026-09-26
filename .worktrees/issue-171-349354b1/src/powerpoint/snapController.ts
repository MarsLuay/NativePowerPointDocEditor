import type { PresentationEngine } from '../PresentationEngine';
import { isSVGGElement } from '../domGuards';
import { debugLog } from '../logger';
import { SNAP_THRESHOLD_PX } from './constants';
import { getShapeIndex, getSvgIntrinsicSize } from './svgUtils';
import type { PointerPoint } from './types';

interface ElementBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * The slice of `NativePowerPointView` that snapping reaches back into. Kept
 * deliberately small: snapping is pure geometry plus guide-DOM, so it borrows
 * only the engine, the slide SVG/pane elements, and two coordinate helpers.
 */
export interface SnapHost {
  readonly engine: PresentationEngine | null;
  readonly svgEl: SVGSVGElement | null;
  readonly canvasPane: HTMLElement | null;
  readonly slideSurface: HTMLElement | null;
  emuPointToPane(emuX: number, emuY: number): PointerPoint | null;
  getElementBox(element: Element): ElementBox | null;
}

export interface SnapResult {
  dx: number;
  dy: number;
  guideX: number | null;
  guideY: number | null;
}

/**
 * Computes alignment snapping for dragged shapes (edges/centers of other shapes
 * and the slide bounds) and renders the snap guide lines. Extracted verbatim
 * from `NativePowerPointView`; it owns only the rendered guide elements.
 */
export class SnapController {
  private snapGuides: HTMLElement[] = [];
  private snapTargetCache: { xs: number[]; ys: number[] } | null = null;
  private verticalGuideEl: HTMLElement | null = null;
  private horizontalGuideEl: HTMLElement | null = null;
  private lastGuideX: number | null | undefined = undefined;
  private lastGuideY: number | null | undefined = undefined;

  constructor(private readonly host: SnapHost) {}

  /** Cache snap targets once per drag so pointermove does not rescan the slide. */
  beginDrag(excluded: Set<number>): void {
    this.snapTargetCache = this.getSnapTargets(excluded);
    this.lastGuideX = undefined;
    this.lastGuideY = undefined;
  }

  endDrag(): void {
    if (this.lastGuideX !== null && this.lastGuideX !== undefined
      || this.lastGuideY !== null && this.lastGuideY !== undefined) {
      debugLog('snap', 'PowerPoint snap applied on commit', {
        guideXEmu: this.lastGuideX,
        guideYEmu: this.lastGuideY
      });
    }
    this.clearSnapGuides();
    this.snapTargetCache = null;
  }

  private getSnapTargets(excluded: Set<number>): { xs: number[]; ys: number[] } {
    const xs: number[] = [];
    const ys: number[] = [];
    if (!this.host.engine || !this.host.svgEl) return { xs, ys };

    this.host.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      if (!isSVGGElement(shape)) return;
      if (shape.parentElement?.closest('g[data-ooxml-shape-idx]')) return;
      const index = getShapeIndex(shape);
      if (index === null || excluded.has(index)) return;
      const transform = this.host.engine?.getShapeTransform(shape);
      if (!transform) return;
      xs.push(transform.x, transform.x + transform.cx / 2, transform.x + transform.cx);
      ys.push(transform.y, transform.y + transform.cy / 2, transform.y + transform.cy);
    });

    const size = getSvgIntrinsicSize(this.host.svgEl);
    const scale = this.host.engine.getSlideScale(this.host.svgEl);
    if (size && scale) {
      const width = size.width * scale;
      const height = size.height * scale;
      xs.push(0, width / 2, width);
      ys.push(0, height / 2, height);
    }
    return { xs, ys };
  }

  computeSnap(
    box: { x: number; y: number; cx: number; cy: number },
    excluded: Set<number>
  ): SnapResult {
    const result: SnapResult = { dx: 0, dy: 0, guideX: null, guideY: null };
    if (!this.host.engine || !this.host.svgEl) return result;

    const ctm = this.host.svgEl.getScreenCTM();
    const scale = this.host.engine.getSlideScale(this.host.svgEl);
    if (!ctm || !scale || ctm.a === 0 || ctm.d === 0) return result;

    const thresholdX = (SNAP_THRESHOLD_PX * scale) / ctm.a;
    const thresholdY = (SNAP_THRESHOLD_PX * scale) / ctm.d;
    const targets = this.snapTargetCache ?? this.getSnapTargets(excluded);
    const xLines = [box.x, box.x + box.cx / 2, box.x + box.cx];
    const yLines = [box.y, box.y + box.cy / 2, box.y + box.cy];
    let bestX = thresholdX + 1;
    let bestY = thresholdY + 1;

    for (const line of xLines) {
      for (const target of targets.xs) {
        const distance = Math.abs(target - line);
        if (distance <= thresholdX && distance < bestX) {
          bestX = distance;
          result.dx = target - line;
          result.guideX = target;
        }
      }
    }
    for (const line of yLines) {
      for (const target of targets.ys) {
        const distance = Math.abs(target - line);
        if (distance <= thresholdY && distance < bestY) {
          bestY = distance;
          result.dy = target - line;
          result.guideY = target;
        }
      }
    }
    return result;
  }

  updateSnapGuides(guideXEmu: number | null, guideYEmu: number | null): void {
    if (this.lastGuideX === guideXEmu && this.lastGuideY === guideYEmu) return;
    const hadGuide = this.lastGuideX !== null && this.lastGuideX !== undefined
      || this.lastGuideY !== null && this.lastGuideY !== undefined;
    const hasGuide = guideXEmu !== null || guideYEmu !== null;
    this.lastGuideX = guideXEmu;
    this.lastGuideY = guideYEmu;

    if (!hadGuide && hasGuide) {
      debugLog('snap', 'PowerPoint snap guides engaged', { guideXEmu, guideYEmu });
    } else if (hadGuide && !hasGuide) {
      debugLog('snap', 'PowerPoint snap guides cleared');
    }

    if (!this.host.canvasPane || !this.host.slideSurface) {
      this.clearSnapGuides();
      return;
    }

    const surface = this.host.getElementBox(this.host.slideSurface);
    if (!surface) return;

    if (guideXEmu === null) {
      if (this.verticalGuideEl) {
        this.verticalGuideEl.remove();
        this.snapGuides = this.snapGuides.filter((guide) => guide !== this.verticalGuideEl);
        this.verticalGuideEl = null;
      }
    } else {
      const point = this.host.emuPointToPane(guideXEmu, 0);
      if (point) {
        if (!this.verticalGuideEl) {
          this.verticalGuideEl = this.host.canvasPane.createDiv({
            cls: 'native-powerpoint-snap-guide native-powerpoint-snap-guide-vertical'
          });
          this.snapGuides.push(this.verticalGuideEl);
        }
        this.verticalGuideEl.setCssProps({
          left: `${point.x}px`,
          top: `${surface.top}px`,
          height: `${surface.height}px`
        });
      }
    }

    if (guideYEmu === null) {
      if (this.horizontalGuideEl) {
        this.horizontalGuideEl.remove();
        this.snapGuides = this.snapGuides.filter((guide) => guide !== this.horizontalGuideEl);
        this.horizontalGuideEl = null;
      }
    } else {
      const point = this.host.emuPointToPane(0, guideYEmu);
      if (point) {
        if (!this.horizontalGuideEl) {
          this.horizontalGuideEl = this.host.canvasPane.createDiv({
            cls: 'native-powerpoint-snap-guide native-powerpoint-snap-guide-horizontal'
          });
          this.snapGuides.push(this.horizontalGuideEl);
        }
        this.horizontalGuideEl.setCssProps({
          left: `${surface.left}px`,
          top: `${point.y}px`,
          width: `${surface.width}px`
        });
      }
    }
  }

  clearSnapGuides(): void {
    if (this.lastGuideX !== null && this.lastGuideX !== undefined
      || this.lastGuideY !== null && this.lastGuideY !== undefined) {
      debugLog('snap', 'PowerPoint snap guides cleared');
    }
    for (const guide of this.snapGuides) {
      guide.remove();
    }
    this.snapGuides = [];
    this.verticalGuideEl = null;
    this.horizontalGuideEl = null;
    this.lastGuideX = undefined;
    this.lastGuideY = undefined;
  }
}
