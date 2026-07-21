import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';
import type { ShapeTransform } from 'pptx-svg';

import { isSVGGElement } from '../domGuards';
import { debugLog, errorLog, logPptxAction } from '../logger';
import type { PresentationEngine, ShapeReorderMode } from '../PresentationEngine';
import { cleanError } from './runtimeCompat';
import type { PresentationSession } from './session/PresentationSession';
import { cloneTransform } from './svgUtils';
import type { DistributeAxis, HistoryEntry } from './types';

export interface ArrangeHost {
  readonly t: TranslateFn;
  readonly session: PresentationSession;
  readonly engine: PresentationEngine | null;
  readonly svgEl: SVGSVGElement | null;
  currentSlide: number;
  ensureEditable(action: string): boolean;
  canEdit(): boolean;
  getSelectedIndices(): number[];
  captureHistoryEntry(label: string): Promise<HistoryEntry>;
  recordHistoryEntry(entry: HistoryEntry): void;
  markDirty(): void;
  renderCurrentSlide(keepSelection?: boolean): Promise<boolean>;
  renderThumbnails(): Promise<void>;
  applyMultiSelection(indices: number[]): void;
  selectShape(shapeIndex: number): void;
  commitGroupTransforms(
    updates: { index: number; transform: ShapeTransform }[],
    label: string
  ): Promise<void>;
  createIconButton(container: HTMLElement, icon: string, label: string, onClick: () => void): HTMLButtonElement;
  updateToolbarButton(button: HTMLButtonElement | null, enabled: boolean): void;
}

export class ArrangeController {
  private distributeButtons: HTMLButtonElement[] = [];
  private zOrderButtons: HTMLButtonElement[] = [];
  /** A reorder reloads the package and shifts numeric shape indexes. */
  private reorderInFlight = false;
  private groupButton: HTMLButtonElement | null = null;
  private ungroupButton: HTMLButtonElement | null = null;
  private readonly notice: TranslateNoticeFn;

  constructor(private readonly host: ArrangeHost) {
    this.notice = createTranslateNotice(this.host.t);
  }

  createToolbarGroups(toolbar: HTMLElement): void {
    this.distributeButtons = [];
    this.zOrderButtons = [];

    const alignGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.distributeButtons.push(
      this.host.createIconButton(
        alignGroup,
        'align-horizontal-distribute-center',
        'Distribute horizontally',
        () => void this.distributeSelectedShapes('horizontal')
      ),
      this.host.createIconButton(
        alignGroup,
        'align-vertical-distribute-center',
        'Distribute vertically',
        () => void this.distributeSelectedShapes('vertical')
      )
    );

    const orderGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.zOrderButtons.push(
      this.host.createIconButton(orderGroup, 'bring-to-front', 'Bring to front', () => void this.reorderSelection('front')),
      this.host.createIconButton(orderGroup, 'arrow-up', 'Bring forward', () => void this.reorderSelection('forward')),
      this.host.createIconButton(orderGroup, 'arrow-down', 'Send backward', () => void this.reorderSelection('backward')),
      this.host.createIconButton(orderGroup, 'send-to-back', 'Send to back', () => void this.reorderSelection('back'))
    );

    const groupGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.groupButton = this.host.createIconButton(groupGroup, 'group', 'Group objects', () => void this.groupSelection());
    this.ungroupButton = this.host.createIconButton(groupGroup, 'ungroup', 'Ungroup objects', () => void this.ungroupSelection());
  }

  updateArrangeAvailability(): void {
    const canEdit = this.host.canEdit();
    const count = this.host.getSelectedIndices().filter((index) => index >= 0).length;
    for (const button of this.distributeButtons) {
      this.host.updateToolbarButton(button, canEdit && count >= 3);
    }
    for (const button of this.zOrderButtons) {
      this.host.updateToolbarButton(button, canEdit && count >= 1 && !this.reorderInFlight);
    }
    this.host.updateToolbarButton(this.groupButton, canEdit && count >= 2);
    this.host.updateToolbarButton(this.ungroupButton, canEdit && this.isSingleGroupSelected());
  }

  async nudgeSelection(key: string, large: boolean): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('move objects')) return;

    const stepEmu = this.host.engine.pxToEmu(large ? 10 : 1);
    let dx = 0;
    let dy = 0;
    if (key === 'ArrowLeft') dx = -stepEmu;
    else if (key === 'ArrowRight') dx = stepEmu;
    else if (key === 'ArrowUp') dy = -stepEmu;
    else if (key === 'ArrowDown') dy = stepEmu;

    const boxes = this.collectSelectedTransforms();
    if (boxes.length === 0) return;

    const updates = boxes.map(({ index, transform }) => {
      const next = cloneTransform(transform);
      next.x += dx;
      next.y += dy;
      return { index, transform: next };
    });
    debugLog('arrange', 'Nudging PowerPoint objects', {
      op: 'nudge',
      slide: this.host.currentSlide,
      count: updates.length,
      key,
      large
    });
    await this.host.commitGroupTransforms(updates, 'Nudge objects');
  }

  private collectSelectedTransforms(): { index: number; transform: ShapeTransform }[] {
    if (!this.host.engine || !this.host.svgEl) return [];
    const result: { index: number; transform: ShapeTransform }[] = [];
    for (const index of this.host.getSelectedIndices()) {
      const shape = this.host.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
      if (isSVGGElement(shape)) {
        result.push({ index, transform: cloneTransform(this.host.engine.getShapeTransform(shape)) });
      }
    }
    return result;
  }

  private async distributeSelectedShapes(axis: DistributeAxis): Promise<void> {
    if (!this.host.ensureEditable('distribute objects')) return;

    const boxes = this.collectSelectedTransforms();
    if (boxes.length < 3) {
      this.notice('powerpoint:notice.selectThreeToDistribute');
      return;
    }

    const horizontal = axis === 'horizontal';
    const center = (transform: ShapeTransform): number =>
      horizontal ? transform.x + transform.cx / 2 : transform.y + transform.cy / 2;
    const sorted = [...boxes].sort((a, b) => center(a.transform) - center(b.transform));
    const first = sorted[0];
    const last = sorted[sorted.length - 1];
    if (!first || !last) return;

    const start = center(first.transform);
    const step = (center(last.transform) - start) / (sorted.length - 1);
    const updates = sorted.map((box, position) => {
      const next = cloneTransform(box.transform);
      const target = start + step * position;
      if (horizontal) {
        next.x = Math.round(target - next.cx / 2);
      } else {
        next.y = Math.round(target - next.cy / 2);
      }
      return { index: box.index, transform: next };
    });

    await this.host.commitGroupTransforms(updates, 'Distribute objects');
    debugLog('arrange', 'Distributed PowerPoint objects', {
      op: 'distribute',
      slide: this.host.currentSlide,
      count: updates.length,
      axis
    });
  }

  async reorderSelection(mode: ShapeReorderMode): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('reorder objects')) return;

    // A second click before the first package reload completes carries the old
    // numeric shape index. It can therefore reorder the neighboring object and
    // undo the first one-step move. Keep this action single-flight.
    if (this.reorderInFlight) {
      debugLog('arrange', 'Ignored overlapping PowerPoint reorder', {
        op: 'reorder',
        slide: this.host.currentSlide,
        shapeIndexes: this.host.getSelectedIndices().filter((index) => index >= 0),
        mode,
      });
      return;
    }

    const indices = this.host.getSelectedIndices().filter((index) => index >= 0);
    if (indices.length === 0) return;

    this.reorderInFlight = true;
    this.updateArrangeAvailability();
    try {
      logPptxAction('arrange', 'reorder', {
        slide: this.host.currentSlide,
        shapeIndexes: indices,
        mode,
        intersectingOnly: true,
      });
      const history = await this.host.captureHistoryEntry('Reorder objects');
      const newIndices = await this.host.session.applyCommand({
        type: 'reorder-shapes',
        slideIndex: this.host.currentSlide,
        shapeIndexes: indices,
        mode,
        intersectingOnly: true,
      }) as number[] | null;
      if (newIndices === null) {
        debugLog('arrange', 'Skipped overlap-aware PowerPoint reorder without an intersecting object', {
          op: 'reorder',
          slide: this.host.currentSlide,
          shapeIndexes: indices,
          mode,
        });
        return;
      }
      this.host.recordHistoryEntry(history);
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        this.host.applyMultiSelection(newIndices.filter((index) => index >= 0));
        await this.host.renderThumbnails();
      }
      debugLog('arrange', 'Reordered PowerPoint objects', {
        op: 'reorder',
        slide: this.host.currentSlide,
        sourceShapeIndexes: indices,
        finalShapeIndexes: newIndices,
        mode,
        rendered,
      });
    } catch (error) {
      errorLog('arrange', 'PowerPoint object reorder failed', { indices, mode, error });
      this.notice('powerpoint:notice.couldNotReorderObjects', { message: cleanError(error) });
    } finally {
      this.reorderInFlight = false;
      this.updateArrangeAvailability();
    }
  }

  private async groupSelection(): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('group objects')) return;

    const indices = this.host.getSelectedIndices().filter((index) => index >= 0);
    if (indices.length < 2) {
      this.notice('powerpoint:notice.selectTwoToGroup');
      return;
    }

    try {
      const history = await this.host.captureHistoryEntry('Group objects');
      const groupIndex = await this.host.session.applyCommand({
        type: 'group-shapes',
        slideIndex: this.host.currentSlide,
        shapeIndexes: indices
      }) as number;
      this.host.recordHistoryEntry(history);
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        this.host.selectShape(groupIndex);
        await this.host.renderThumbnails();
      }
      debugLog('arrange', 'Grouped PowerPoint objects', {
        op: 'group',
        slide: this.host.currentSlide,
        count: indices.length,
        groupIndex
      });
    } catch (error) {
      errorLog('arrange', 'PowerPoint object grouping failed', { indices, error });
      this.notice('powerpoint:notice.couldNotGroupObjects', { message: cleanError(error) });
    }
  }

  private async ungroupSelection(): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('ungroup objects')) return;

    if (!this.isSingleGroupSelected()) {
      this.notice('powerpoint:notice.selectGroupToUngroup');
      return;
    }

    const [groupIndex] = this.host.getSelectedIndices();
    if (groupIndex === undefined) return;

    try {
      const history = await this.host.captureHistoryEntry('Ungroup objects');
      const newIndices = await this.host.session.applyCommand({
        type: 'ungroup-shapes',
        slideIndex: this.host.currentSlide,
        shapeIndex: groupIndex
      }) as number[];
      this.host.recordHistoryEntry(history);
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        this.host.applyMultiSelection(newIndices.filter((index) => index >= 0));
        await this.host.renderThumbnails();
      }
      debugLog('arrange', 'Ungrouped PowerPoint objects', {
        op: 'ungroup',
        slide: this.host.currentSlide,
        groupIndex,
        resultCount: newIndices.length
      });
    } catch (error) {
      errorLog('arrange', 'PowerPoint object ungrouping failed', { groupIndex, error });
      this.notice('powerpoint:notice.couldNotUngroupObjects', { message: cleanError(error) });
    }
  }

  private isSingleGroupSelected(): boolean {
    const indices = this.host.getSelectedIndices().filter((index) => index >= 0);
    if (indices.length !== 1) return false;
    const shape = this.host.svgEl?.querySelector(`g[data-ooxml-shape-idx="${indices[0]}"]`);
    return shape?.getAttribute('data-ooxml-shape-type') === 'group';
  }
}
