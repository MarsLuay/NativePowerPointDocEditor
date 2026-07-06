import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';
import type { ShapeTransform } from 'pptx-svg';
import { isSVGGElement } from '../domGuards';
import type { PresentationEngine } from '../PresentationEngine';
import { cleanError } from './runtimeCompat';
import type { SnapController } from './snapController';
import type { DragState, GroupDragState, HandleName, HistoryEntry, MarqueeState, PointerPoint } from './types';
import { cloneTransform, getShapeIndex, transformsMatch } from './svgUtils';

export interface SelectionDragHost {
  readonly t: TranslateFn;
  readonly engine: PresentationEngine | null;
  readonly svgEl: SVGSVGElement | null;
  readonly canvasPane: HTMLElement | null;
  selectedShapeIndex: number | null;
  readonly selectedShapeIndices: Set<number>;
  selectedTransform: ShapeTransform | null;
  readonly snapController: SnapController;
  suppressNextClick: boolean;
  currentSlide: number;
  ensureEditable(action: string): boolean;
  canEdit(): boolean;
  getSelectedShapeElement(): SVGGElement | null;
  getElementBox(element: Element): { left: number; top: number; width: number; height: number } | null;
  emuPointToPane(emuX: number, emuY: number): { x: number; y: number } | null;
  applyMultiSelection(indices: number[]): void;
  clearSelection(options?: { skipTextCommit?: boolean }): void;
  captureHistoryEntry(label: string): Promise<HistoryEntry>;
  recordHistoryEntry(entry: HistoryEntry): void;
  markDirty(): void;
  renderCurrentSlide(keepSelection?: boolean): Promise<boolean>;
  renderEditedShape(shapeIndex: number): Promise<boolean>;
  renderThumbnails(): Promise<void>;
  updateInspectorValues(): void;
  updateTextToolbar(): void;
}

export class SelectionDragController {
  dragState: DragState | null = null;
  marquee: MarqueeState | null = null;
  marqueeEl: HTMLElement | null = null;
  groupDrag: GroupDragState | null = null;
  multiSelectionBoxes: HTMLElement[] = [];
  selectionOverlay: HTMLElement | null = null;
  private readonly notice: TranslateNoticeFn;

  constructor(private readonly host: SelectionDragHost) {
    this.notice = createTranslateNotice(this.host.t);
  }

  clearDragState(): void {
    this.dragState = null;
    this.groupDrag = null;
    this.cancelMarquee();
  }

  positionOverlayFromTransform(transform: ShapeTransform): void {
      if (!this.selectionOverlay) return;
      const topLeft = this.host.emuPointToPane(transform.x, transform.y);
      const bottomRight = this.host.emuPointToPane(transform.x + transform.cx, transform.y + transform.cy);
      if (!topLeft || !bottomRight) return;
      this.selectionOverlay.setCssProps({
        left: `${topLeft.x}px`,
        top: `${topLeft.y}px`,
        width: `${Math.max(0, bottomRight.x - topLeft.x)}px`,
        height: `${Math.max(0, bottomRight.y - topLeft.y)}px`
      });
    }

  startRotateDrag(event: PointerEvent): void {
      if (!this.host.engine || this.host.selectedTransform === null || !this.selectionOverlay) return;
      if (!this.host.ensureEditable('rotate object')) return;

      const rect = this.selectionOverlay.getBoundingClientRect();
      const centerClientX = rect.left + rect.width / 2;
      const centerClientY = rect.top + rect.height / 2;
      const startBox = this.getSelectedBox();
      if (!startBox) return;

      this.dragState = {
        mode: 'rotate',
        pointerId: event.pointerId,
        startPoint: { x: event.clientX, y: event.clientY },
        startClientX: event.clientX,
        startClientY: event.clientY,
        startBox,
        startTransform: cloneTransform(this.host.selectedTransform),
        latestTransform: cloneTransform(this.host.selectedTransform),
        centerClientX,
        centerClientY,
        startAngle: Math.atan2(event.clientY - centerClientY, event.clientX - centerClientX)
      };
    }
  updateSelectionOverlay(): void {
      this.updateMultiSelectionBoxes();
      if (!this.host.canvasPane || this.host.selectedShapeIndex === null) {
        this.removeSelectionOverlay();
        this.host.updateTextToolbar();
        return;
      }

      const box = this.getSelectedBox();
      if (!box) {
        this.removeSelectionOverlay();
        this.host.updateTextToolbar();
        return;
      }

      if (!this.selectionOverlay) {
        this.selectionOverlay = this.host.canvasPane.createDiv({ cls: 'native-powerpoint-selection-box' });
        if (this.host.canEdit()) {
          // Edge hit-zones first so the corner dots stack above them at overlaps.
          // Each edge stretches the object along a single axis.
          for (const handle of ['n', 'e', 's', 'w'] as HandleName[]) {
            const edgeEl = this.selectionOverlay.createDiv({ cls: `native-powerpoint-resize-edge native-powerpoint-resize-${handle}` });
            edgeEl.addEventListener('pointerdown', (event) => {
              event.preventDefault();
              event.stopPropagation();
              this.startDrag(event, 'resize', handle);
            });
          }

          for (const handle of ['nw', 'ne', 'sw', 'se'] as HandleName[]) {
            const handleEl = this.selectionOverlay.createDiv({ cls: `native-powerpoint-resize-handle native-powerpoint-resize-${handle}` });
            handleEl.addEventListener('pointerdown', (event) => {
              event.preventDefault();
              event.stopPropagation();
              this.startDrag(event, 'resize', handle);
            });
          }

          const rotateStem = this.selectionOverlay.createDiv({ cls: 'native-powerpoint-rotate-stem' });
          rotateStem.setAttribute('aria-hidden', 'true');
          const rotateHandle = this.selectionOverlay.createDiv({ cls: 'native-powerpoint-rotate-handle' });
          rotateHandle.setAttribute('aria-label', this.host.t('powerpoint:accessibility.rotateObject'));
          rotateHandle.setAttribute('data-tooltip', this.host.t('powerpoint:accessibility.rotateObject'));
          rotateHandle.addEventListener('pointerdown', (event) => {
            event.preventDefault();
            event.stopPropagation();
            this.startRotateDrag(event);
          });
        }
      }

      this.selectionOverlay.style.removeProperty('transform');
      this.selectionOverlay.setCssProps({
        left: `${box.left}px`,
        top: `${box.top}px`,
        width: `${box.width}px`,
        height: `${box.height}px`
      });

      this.host.updateTextToolbar();
    }

  removeSelectionOverlay(): void {
      this.selectionOverlay?.remove();
      this.selectionOverlay = null;
    }

    updateMultiSelectionBoxes(): void {
      this.removeMultiSelectionBoxes();
      if (!this.host.canvasPane || !this.host.svgEl || this.host.selectedShapeIndices.size <= 1) return;

      for (const index of this.host.selectedShapeIndices) {
        const shape = this.host.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
        if (!isSVGGElement(shape)) continue;

        const box = this.host.getElementBox(shape);
        if (!box) continue;

        const boxEl = this.host.canvasPane.createDiv({
          cls: 'native-powerpoint-selection-box native-powerpoint-multi-selection-box'
        });
        boxEl.setCssProps({
          left: `${box.left}px`,
          top: `${box.top}px`,
          width: `${box.width}px`,
          height: `${box.height}px`
        });
        this.multiSelectionBoxes.push(boxEl);
      }
    }

  removeMultiSelectionBoxes(): void {
      for (const box of this.multiSelectionBoxes) {
        box.remove();
      }
      this.multiSelectionBoxes = [];
    }

  collectShapesInClientRect(left: number, top: number, right: number, bottom: number): number[] {
      if (!this.host.svgEl) return [];

      const indices: number[] = [];
      this.host.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
        if (!isSVGGElement(shape)) return;
        if (shape.parentElement?.closest('g[data-ooxml-shape-idx]')) return;

        const index = getShapeIndex(shape);
        if (index === null) return;

        const rect = shape.getBoundingClientRect();
        if (rect.width === 0 && rect.height === 0) return;

        const intersects =
          rect.left <= right && rect.right >= left && rect.top <= bottom && rect.bottom >= top;
        if (intersects) indices.push(index);
      });
      return indices;
    }

  previewSelectionClasses(indices: Set<number>): void {
      this.host.svgEl?.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
        const index = getShapeIndex(shape);
        if (index !== null && indices.has(index)) {
          shape.addClass('native-powerpoint-shape-selected');
        } else {
          shape.removeClass('native-powerpoint-shape-selected');
        }
      });
    }

  beginMarquee(event: PointerEvent, additive: boolean): void {
      if (!this.host.canvasPane) return;

      this.cancelMarquee();
      this.marquee = {
        pointerId: event.pointerId,
        startClientX: event.clientX,
        startClientY: event.clientY,
        additive,
        base: [...this.host.selectedShapeIndices],
        moved: false
      };
    }

    updateMarquee(event: PointerEvent): void {
      if (!this.marquee || event.pointerId !== this.marquee.pointerId || !this.host.canvasPane) return;

      const deltaX = event.clientX - this.marquee.startClientX;
      const deltaY = event.clientY - this.marquee.startClientY;
      if (!this.marquee.moved && Math.hypot(deltaX, deltaY) < 4) return;
      if (!this.marquee.moved) {
        this.removeSelectionOverlay();
        this.removeMultiSelectionBoxes();
      }
      this.marquee.moved = true;

      if (!this.marqueeEl) {
        this.marqueeEl = this.host.canvasPane.createDiv({ cls: 'native-powerpoint-marquee-box' });
      }

      const paneRect = this.host.canvasPane.getBoundingClientRect();
      const left = Math.min(event.clientX, this.marquee.startClientX);
      const top = Math.min(event.clientY, this.marquee.startClientY);
      const width = Math.abs(deltaX);
      const height = Math.abs(deltaY);
      this.marqueeEl.setCssProps({
        left: `${left - paneRect.left + this.host.canvasPane.scrollLeft}px`,
        top: `${top - paneRect.top + this.host.canvasPane.scrollTop}px`,
        width: `${width}px`,
        height: `${height}px`
      });

      const hits = this.collectShapesInClientRect(left, top, left + width, top + height);
      const preview = new Set<number>(this.marquee.additive ? this.marquee.base : []);
      hits.forEach((index) => preview.add(index));
      this.previewSelectionClasses(preview);
    }

  finishMarquee(event: PointerEvent): void {
      if (!this.marquee || event.pointerId !== this.marquee.pointerId) return;

      const marquee = this.marquee;
      this.marquee = null;
      this.marqueeEl?.remove();
      this.marqueeEl = null;

      if (!marquee.moved) {
        if (marquee.additive) {
          this.host.suppressNextClick = true;
          this.host.applyMultiSelection(marquee.base);
        } else {
          this.host.clearSelection();
        }
        return;
      }

      this.host.suppressNextClick = true;
      const left = Math.min(event.clientX, marquee.startClientX);
      const top = Math.min(event.clientY, marquee.startClientY);
      const right = Math.max(event.clientX, marquee.startClientX);
      const bottom = Math.max(event.clientY, marquee.startClientY);
      const hits = this.collectShapesInClientRect(left, top, right, bottom);
      const finalSet = new Set<number>(marquee.additive ? marquee.base : []);
      hits.forEach((index) => finalSet.add(index));
      this.host.applyMultiSelection([...finalSet]);
    }

  cancelMarquee(): void {
      this.marquee = null;
      this.marqueeEl?.remove();
      this.marqueeEl = null;
    }

  startGroupDrag(event: PointerEvent): void {
      if (!this.host.engine || !this.host.svgEl) return;

      const startPoint = this.getSvgPoint(event);
      if (!startPoint) return;

      const start = new Map<number, ShapeTransform>();
      for (const index of this.host.selectedShapeIndices) {
        const shape = this.host.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
        if (isSVGGElement(shape)) {
          start.set(index, cloneTransform(this.host.engine.getShapeTransform(shape)));
        }
      }
      if (start.size === 0) return;

      this.groupDrag = {
        pointerId: event.pointerId,
        startPoint,
        startClientX: event.clientX,
        startClientY: event.clientY,
        start,
        latest: new Map(start),
        moved: false
      };
    }

  updateGroupDrag(event: PointerEvent): void {
      if (!this.groupDrag || event.pointerId !== this.groupDrag.pointerId || !this.host.engine || !this.host.svgEl) {
        return;
      }

      const deltaClientX = event.clientX - this.groupDrag.startClientX;
      const deltaClientY = event.clientY - this.groupDrag.startClientY;
      if (!this.groupDrag.moved && Math.hypot(deltaClientX, deltaClientY) < 3) return;
      this.groupDrag.moved = true;

      const point = this.getSvgPoint(event);
      if (!point) return;

      const scale = this.host.engine.getSlideScale(this.host.svgEl);
      const dx = (point.x - this.groupDrag.startPoint.x) * scale;
      const dy = (point.y - this.groupDrag.startPoint.y) * scale;
      let minX = Infinity;
      let minY = Infinity;
      let maxX = -Infinity;
      let maxY = -Infinity;
      this.groupDrag.start.forEach((transform, index) => {
        const next = cloneTransform(transform);
        next.x += dx;
        next.y += dy;
        this.groupDrag?.latest.set(index, next);
        minX = Math.min(minX, next.x);
        minY = Math.min(minY, next.y);
        maxX = Math.max(maxX, next.x + next.cx);
        maxY = Math.max(maxY, next.y + next.cy);
      });

      const snap = Number.isFinite(minX)
        ? this.host.snapController.computeSnap(
            { x: minX, y: minY, cx: maxX - minX, cy: maxY - minY },
            new Set(this.host.selectedShapeIndices)
          )
        : { dx: 0, dy: 0, guideX: null, guideY: null };
      if (snap.dx !== 0 || snap.dy !== 0) {
        this.groupDrag.latest.forEach((transform) => {
          transform.x += snap.dx;
          transform.y += snap.dy;
        });
      }
      this.host.snapController.updateSnapGuides(snap.guideX, snap.guideY);

      const ctm = this.host.svgEl.getScreenCTM();
      const snapClientX = ctm && ctm.a !== 0 ? (snap.dx * ctm.a) / scale : 0;
      const snapClientY = ctm && ctm.d !== 0 ? (snap.dy * ctm.d) / scale : 0;
      const cssTransform = `translate(${deltaClientX + snapClientX}px, ${deltaClientY + snapClientY}px)`;
      for (const box of this.multiSelectionBoxes) {
        box.style.transform = cssTransform;
      }
    }

  finishGroupDrag(event: PointerEvent): void {
      if (!this.groupDrag || event.pointerId !== this.groupDrag.pointerId) return;

      const groupDrag = this.groupDrag;
      this.groupDrag = null;
      this.host.snapController.clearSnapGuides();
      for (const box of this.multiSelectionBoxes) {
        box.style.removeProperty('transform');
      }

      if (!groupDrag.moved) return;

      this.host.suppressNextClick = true;
      const updates = [...groupDrag.latest.entries()].map(([index, transform]) => ({ index, transform }));
      void this.commitGroupTransforms(updates);
    }

  async commitGroupTransforms(
      updates: { index: number; transform: ShapeTransform }[],
      label = 'Move objects'
    ): Promise<void> {
      if (!this.host.engine || updates.length === 0) return;
      if (!this.host.ensureEditable('move objects')) return;

      try {
        const changed = updates.filter((update) => {
          const shape = this.host.svgEl?.querySelector(`g[data-ooxml-shape-idx="${update.index}"]`);
          return !(
            isSVGGElement(shape)
            && this.host.engine !== null
            && transformsMatch(this.host.engine.getShapeTransform(shape), update.transform)
          );
        });
        if (changed.length === 0) return;

        const history = await this.host.captureHistoryEntry(label);
        for (const update of changed) {
          await this.host.engine.updateShapeTransform(this.host.currentSlide, update.index, update.transform);
        }
        this.host.recordHistoryEntry(history);
        this.host.markDirty();
        const indices = updates.map((update) => update.index);
        const rendered = await this.host.renderCurrentSlide();
        if (rendered) {
          this.host.applyMultiSelection(indices);
          await this.host.renderThumbnails();
        }
      } catch (error) {
        this.notice('powerpoint:notice.couldNotMoveObjects', { message: cleanError(error) });
      }
    }

  getSelectedBox(): { left: number; top: number; width: number; height: number } | null {
      const selected = this.host.getSelectedShapeElement();
      if (!selected) return null;

      return this.host.getElementBox(selected);
    }

  startDrag(event: PointerEvent, mode: 'move' | 'resize', handle?: HandleName): void {
      if (!this.host.engine || !this.host.svgEl || this.host.selectedTransform === null) return;
      if (!this.host.ensureEditable(mode === 'move' ? 'move object' : 'resize object')) return;

      const startPoint = this.getSvgPoint(event);
      const startBox = this.getSelectedBox();
      if (!startPoint || !startBox) return;

      this.dragState = {
        mode,
        handle,
        pointerId: event.pointerId,
        startPoint,
        startClientX: event.clientX,
        startClientY: event.clientY,
        startBox,
        startTransform: cloneTransform(this.host.selectedTransform),
        latestTransform: cloneTransform(this.host.selectedTransform)
      };
    }

  handleDragMove = (event: PointerEvent): void => {
      if (this.marquee) {
        this.updateMarquee(event);
        return;
      }
      if (this.groupDrag) {
        this.updateGroupDrag(event);
        return;
      }
      if (!this.dragState || !this.host.engine || !this.host.svgEl) return;
      if (event.pointerId !== this.dragState.pointerId) return;

      if (this.dragState.mode === 'rotate') {
        this.updateRotateDrag(event);
        return;
      }

      const point = this.getSvgPoint(event);
      if (!point) return;

      const scale = this.host.engine.getSlideScale(this.host.svgEl);
      const dx = (point.x - this.dragState.startPoint.x) * scale;
      const dy = (point.y - this.dragState.startPoint.y) * scale;
      const next = cloneTransform(this.dragState.startTransform);

      if (this.dragState.mode === 'move') {
        next.x += dx;
        next.y += dy;
        const snap = this.host.snapController.computeSnap(
          { x: next.x, y: next.y, cx: next.cx, cy: next.cy },
          new Set(this.host.selectedShapeIndex === null ? [] : [this.host.selectedShapeIndex])
        );
        next.x += snap.dx;
        next.y += snap.dy;
        this.host.snapController.updateSnapGuides(snap.guideX, snap.guideY);
        this.dragState.latestTransform = next;
        this.host.selectedTransform = cloneTransform(next);
        this.host.updateInspectorValues();
        this.positionOverlayFromTransform(next);
        return;
      }

      const minSize = this.host.engine.pxToEmu(12);
      const start = this.dragState.startTransform;
      if (this.dragState.handle?.includes('w')) {
        next.x += dx;
        next.cx -= dx;
        if (next.cx < minSize) {
          next.cx = minSize;
          // Pin the (fixed) east edge instead of letting the shape drift left.
          next.x = start.x + start.cx - minSize;
        }
      }
      if (this.dragState.handle?.includes('e')) {
        next.cx += dx;
        if (next.cx < minSize) next.cx = minSize;
      }
      if (this.dragState.handle?.includes('n')) {
        next.y += dy;
        next.cy -= dy;
        if (next.cy < minSize) {
          next.cy = minSize;
          // Pin the (fixed) south edge instead of letting the shape drift up.
          next.y = start.y + start.cy - minSize;
        }
      }
      if (this.dragState.handle?.includes('s')) {
        next.cy += dy;
        if (next.cy < minSize) next.cy = minSize;
      }

      this.dragState.latestTransform = next;
      this.host.selectedTransform = cloneTransform(next);
      this.host.updateInspectorValues();
      this.updateSelectionOverlayDuringDrag(event);
    };

  updateRotateDrag(event: PointerEvent): void {
      if (!this.host.engine || !this.dragState) return;

      const centerX = this.dragState.centerClientX ?? 0;
      const centerY = this.dragState.centerClientY ?? 0;
      const angle = Math.atan2(event.clientY - centerY, event.clientX - centerX);
      const deltaDegrees = ((angle - (this.dragState.startAngle ?? 0)) * 180) / Math.PI;
      let degrees = this.host.engine.ooxmlToDegrees(this.dragState.startTransform.rot) + deltaDegrees;
      if (event.shiftKey) degrees = Math.round(degrees / 15) * 15;
      degrees = ((degrees % 360) + 360) % 360;

      const next = cloneTransform(this.dragState.startTransform);
      next.rot = this.host.engine.degreesToOoxml(degrees);
      this.dragState.latestTransform = next;
      this.host.selectedTransform = cloneTransform(next);
      this.host.updateInspectorValues();
      if (this.selectionOverlay) {
        this.selectionOverlay.style.transform = `rotate(${degrees}deg)`;
      }
    }

  handleDragEnd = (event: PointerEvent): void => {
      if (this.marquee) {
        this.finishMarquee(event);
        return;
      }
      if (this.groupDrag) {
        this.finishGroupDrag(event);
        return;
      }
      if (!this.dragState || event.pointerId !== this.dragState.pointerId) return;

      this.host.snapController.clearSnapGuides();
      const transform = cloneTransform(this.dragState.latestTransform);
      this.dragState = null;
      void this.commitTransform(transform);
    };

  async commitTransform(transform: ShapeTransform): Promise<void> {
      if (!this.host.engine || this.host.selectedShapeIndex === null) return;
      if (!this.host.ensureEditable('edit object')) return;

      try {
        const selected = this.host.getSelectedShapeElement();
        const shapeIndex = selected ? getShapeIndex(selected) : this.host.selectedShapeIndex;
        if (shapeIndex === null) return;
        if (selected && transformsMatch(this.host.engine.getShapeTransform(selected), transform)) return;

        const history = await this.host.captureHistoryEntry('Edit layout');
        await this.host.engine.updateShapeTransform(this.host.currentSlide, shapeIndex, transform);
        this.host.selectedTransform = cloneTransform(transform);
        this.host.recordHistoryEntry(history);
        this.host.markDirty();
        const rendered = await this.host.renderEditedShape(shapeIndex);
        if (rendered) await this.host.renderThumbnails();
      } catch (error) {
        this.notice('powerpoint:notice.couldNotUpdateObject', { message: cleanError(error) });
      }
    }

  getSvgPoint(event: PointerEvent): PointerPoint | null {
      if (!this.host.svgEl) return null;

      const matrix = this.host.svgEl.getScreenCTM();
      if (!matrix) return null;

      const point = this.host.svgEl.createSVGPoint();
      point.x = event.clientX;
      point.y = event.clientY;
      const result = point.matrixTransform(matrix.inverse());
      return { x: result.x, y: result.y };
    }

  updateSelectionOverlayDuringDrag(event: PointerEvent): void {
      if (!this.dragState || !this.selectionOverlay) return;

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
        box.width = Math.max(12, box.width);
        box.height = Math.max(12, box.height);
      }

      this.selectionOverlay.setCssProps({
        left: `${box.left}px`,
        top: `${box.top}px`,
        width: `${box.width}px`,
        height: `${box.height}px`
      });
    }
}
