import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';
import type { ShapeTransform } from 'pptx-svg';

import type { ChartDataGrid, ChartDataUpdate } from '../ChartData';
import type { FontSubstitution } from '../FontFidelity';
import { debugLog, errorLog } from '../logger';
import type { PresentationEngine } from '../PresentationEngine';
import { GENERATED_GRID_SELECTOR } from './constants';
import { cleanError } from './runtimeCompat';
import { cloneTransform } from './svgUtils';
import type { HistoryEntry } from './types';

export interface InspectorHost {
  readonly t: TranslateFn;
  readonly engine: PresentationEngine | null;
  readonly inspectorEl: HTMLElement | null;
  readonly selectedShapeIndex: number | null;
  readonly selectedShapeIndices: Set<number>;
  readonly selectedTransform: ShapeTransform | null;
  currentSlide: number;
  readonly isViewOnly: boolean;
  readonly viewOnlyReason: string;
  readonly fontSubstitutions: FontSubstitution[];
  ensureEditable(action: string): boolean;
  canEdit(): boolean;
  getSelectedShapeElement(): Element | null;
  captureHistoryEntry(label: string): Promise<HistoryEntry>;
  recordHistoryEntry(entry: HistoryEntry): void;
  markDirty(): void;
  renderCurrentSlide(keepSelection?: boolean): Promise<boolean>;
  renderThumbnails(): Promise<void>;
  renderEditedShape(shapeIndex: number): Promise<boolean>;
  commitTransform(transform: ShapeTransform): Promise<void>;
}

export class InspectorController {
  private xInput: HTMLInputElement | null = null;
  private yInput: HTMLInputElement | null = null;
  private widthInput: HTMLInputElement | null = null;
  private heightInput: HTMLInputElement | null = null;
  private rotationInput: HTMLInputElement | null = null;
  private readonly notice: TranslateNoticeFn;

  constructor(private readonly host: InspectorHost) {
    this.notice = createTranslateNotice(this.host.t);
  }

  render(): void {
    if (!this.host.inspectorEl) return;

    const t = this.host.t;
    this.host.inspectorEl.empty();
    this.host.inspectorEl.createDiv({ cls: 'native-powerpoint-inspector-title', text: t('powerpoint:inspector.title') });
    this.renderViewOnlyWarning(this.host.inspectorEl);
    this.renderFontFidelity(this.host.inspectorEl);
    this.renderSlideBackgroundControl(this.host.inspectorEl);

    if (!this.host.engine || this.host.selectedShapeIndex === null || !this.host.selectedTransform) {
      this.host.inspectorEl.createDiv({
        cls: 'native-powerpoint-inspector-empty',
        text: this.host.selectedShapeIndices.size > 1
          ? t('powerpoint:inspector.emptyMultiSelect', { count: this.host.selectedShapeIndices.size })
          : t('powerpoint:inspector.emptySelectObject')
      });
      this.xInput = null;
      this.yInput = null;
      this.widthInput = null;
      this.heightInput = null;
      this.rotationInput = null;
      return;
    }

    const selected = this.host.getSelectedShapeElement();
    this.host.inspectorEl.createDiv({
      cls: 'native-powerpoint-inspector-subtitle',
      text: t('powerpoint:inspector.objectNumber', { number: this.host.selectedShapeIndex + 1 })
    });
    this.host.inspectorEl.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: selected?.closest(GENERATED_GRID_SELECTOR)
        ? t('powerpoint:inspector.hintEditGenerated')
        : t('powerpoint:inspector.hintEditText')
    });

    const grid = this.host.inspectorEl.createDiv({ cls: 'native-powerpoint-inspector-grid' });
    this.xInput = this.createNumberField(grid, t('powerpoint:inspector.fieldX'), this.host.engine.emuToPx(this.host.selectedTransform.x));
    this.yInput = this.createNumberField(grid, t('powerpoint:inspector.fieldY'), this.host.engine.emuToPx(this.host.selectedTransform.y));
    this.widthInput = this.createNumberField(grid, t('powerpoint:inspector.fieldW'), this.host.engine.emuToPx(this.host.selectedTransform.cx));
    this.heightInput = this.createNumberField(grid, t('powerpoint:inspector.fieldH'), this.host.engine.emuToPx(this.host.selectedTransform.cy));
    this.rotationInput = this.createNumberField(
      grid,
      t('powerpoint:inspector.fieldRot'),
      this.host.engine.ooxmlToDegrees(this.host.selectedTransform.rot)
    );
    this.xInput.disabled = !this.host.canEdit();
    this.yInput.disabled = !this.host.canEdit();
    this.widthInput.disabled = !this.host.canEdit();
    this.heightInput.disabled = !this.host.canEdit();
    this.rotationInput.disabled = !this.host.canEdit();

    const applyLayout = this.host.inspectorEl.createEl('button', {
      cls: 'native-powerpoint-inspector-button',
      text: t('powerpoint:inspector.applyLayout')
    });
    applyLayout.disabled = !this.host.canEdit();
    applyLayout.addEventListener('click', () => void this.applyInspectorTransform());

    if (selected?.getAttribute('data-ooxml-shape-type') === 'chart') {
      const chartData = this.host.engine.getChartDataGrid(this.host.currentSlide, this.host.selectedShapeIndex);
      if (chartData) {
        this.renderChartDataEditor(chartData);
      }
    }
  }

  updateValues(): void {
    if (!this.host.engine || !this.host.selectedTransform) return;

    if (this.xInput) {
      this.xInput.value = String(Math.round(this.host.engine.emuToPx(this.host.selectedTransform.x) * 100) / 100);
    }
    if (this.yInput) {
      this.yInput.value = String(Math.round(this.host.engine.emuToPx(this.host.selectedTransform.y) * 100) / 100);
    }
    if (this.widthInput) {
      this.widthInput.value = String(Math.round(this.host.engine.emuToPx(this.host.selectedTransform.cx) * 100) / 100);
    }
    if (this.heightInput) {
      this.heightInput.value = String(Math.round(this.host.engine.emuToPx(this.host.selectedTransform.cy) * 100) / 100);
    }
    if (this.rotationInput) {
      this.rotationInput.value = String(
        Math.round(this.host.engine.ooxmlToDegrees(this.host.selectedTransform.rot) * 100) / 100
      );
    }
  }

  private renderSlideBackgroundControl(container: HTMLElement): void {
    if (!this.host.engine || this.host.engine.slideCount === 0) return;

    const t = this.host.t;
    const section = container.createDiv({ cls: 'native-powerpoint-slide-background' });
    section.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: t('powerpoint:inspector.slideBackground') });
    section.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: t('powerpoint:inspector.slideBackgroundHint')
    });

    const currentColor = this.host.engine.getSlideBackgroundColor(this.host.currentSlide);
    const row = section.createDiv({ cls: 'native-powerpoint-background-row' });
    const colorInput = row.createEl('input', {
      type: 'color',
      cls: 'native-powerpoint-background-color',
      value: currentColor ? `#${currentColor}` : '#ffffff'
    });
    colorInput.disabled = !this.host.canEdit();

    const applyButton = row.createEl('button', {
      cls: 'native-powerpoint-inspector-button',
      text: t('powerpoint:inspector.apply')
    });
    applyButton.disabled = !this.host.canEdit();
    applyButton.addEventListener('click', () => {
      void this.applySlideBackgroundColor(colorInput.value);
    });
  }

  private async applySlideBackgroundColor(hexColor: string): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('change slide background')) return;

    try {
      const history = await this.host.captureHistoryEntry('Slide background');
      await this.host.engine.setSlideBackgroundColor(this.host.currentSlide, hexColor);
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide(true);
      if (rendered) {
        await this.host.renderThumbnails();
        this.render();
      }
      debugLog('inspector', 'Changed PowerPoint slide background', {
        slide: this.host.currentSlide,
        color: hexColor
      });
    } catch (error) {
      errorLog('inspector', 'PowerPoint slide background change failed', {
        slide: this.host.currentSlide,
        color: hexColor,
        error
      });
      this.notice('powerpoint:notice.couldNotChangeSlideBackground', { message: cleanError(error) });
    }
  }

  private renderChartDataEditor(chartData: ChartDataGrid): void {
    if (!this.host.inspectorEl) return;

    const t = this.host.t;
    const section = this.host.inspectorEl.createDiv({ cls: 'native-powerpoint-chart-data' });
    section.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: t('powerpoint:inspector.chartData') });

    if (!chartData.editable) {
      section.createDiv({
        cls: 'native-powerpoint-inspector-hint',
        text: chartData.reason || t('powerpoint:inspector.chartDataReadOnly')
      });
      return;
    }

    section.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: t('powerpoint:inspector.chartDataHint')
    });

    const viewport = section.createDiv({ cls: 'native-powerpoint-chart-data-scroll' });
    const table = viewport.createEl('table', { cls: 'native-powerpoint-chart-data-grid' });
    const header = table.createEl('thead').createEl('tr');
    header.createEl('th', { text: chartData.categoryLabel });
    chartData.series.forEach((series) => {
      header.createEl('th', { text: series.name });
      if (series.pointLabels !== null) {
        header.createEl('th', { text: t('powerpoint:inspector.seriesLabel', { seriesName: series.name }) });
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
      text: t('powerpoint:inspector.applyChartData')
    });
    apply.disabled = !this.host.canEdit();
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
    input.disabled = !this.host.canEdit();
    return input;
  }

  private async applyChartData(update: ChartDataUpdate): Promise<void> {
    if (!this.host.engine || this.host.selectedShapeIndex === null) return;
    if (!this.host.ensureEditable('edit chart data')) return;

    const chartShapeIndex = this.host.selectedShapeIndex;
    try {
      const history = await this.host.captureHistoryEntry('Edit chart data');
      await this.host.engine.updateChartData(this.host.currentSlide, chartShapeIndex, update);
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderEditedShape(chartShapeIndex);
      if (rendered) await this.host.renderThumbnails();
      debugLog('inspector', 'Updated PowerPoint chart data', {
        slide: this.host.currentSlide,
        shapeIndex: chartShapeIndex,
        categoryCount: update.categories.length,
        seriesCount: update.series.length
      });
    } catch (error) {
      errorLog('inspector', 'PowerPoint chart-data update failed', {
        slide: this.host.currentSlide,
        shapeIndex: chartShapeIndex,
        categoryCount: update.categories.length,
        seriesCount: update.series.length,
        error
      });
      this.notice('powerpoint:notice.couldNotUpdateChartData', { message: cleanError(error) });
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
    if (!this.host.isViewOnly || !this.host.viewOnlyReason) return;

    container.createDiv({
      cls: 'native-powerpoint-view-only-warning',
      text: this.host.viewOnlyReason
    });
  }

  private renderFontFidelity(container: HTMLElement): void {
    if (!this.host.engine) return;

    const t = this.host.t;
    const section = container.createDiv({ cls: 'native-powerpoint-font-fidelity' });
    section.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: t('powerpoint:inspector.fonts') });

    if (this.host.fontSubstitutions.length === 0) {
      section.createDiv({
        cls: 'native-powerpoint-inspector-hint',
        text: t('powerpoint:inspector.fontsAvailable')
      });
      return;
    }

    section.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: t('powerpoint:inspector.fontsSubstituted', { count: this.host.fontSubstitutions.length })
    });
    const list = section.createDiv({ cls: 'native-powerpoint-font-substitution-list' });
    for (const substitution of this.host.fontSubstitutions) {
      const item = list.createDiv({ cls: 'native-powerpoint-font-substitution' });
      item.createSpan({ cls: 'native-powerpoint-font-substitution-source', text: substitution.requested });
      item.createSpan({ cls: 'native-powerpoint-font-substitution-arrow', text: this.host.t('powerpoint:inspector.substitutionArrow') });
      item.createSpan({ cls: 'native-powerpoint-font-substitution-target', text: substitution.substitute });
    }
  }

  private async applyInspectorTransform(): Promise<void> {
    if (!this.host.engine || this.host.selectedShapeIndex === null || !this.host.selectedTransform) return;
    if (!this.host.ensureEditable('edit layout')) return;

    const transform = cloneTransform(this.host.selectedTransform);
    transform.x = this.host.engine.pxToEmu(Number(this.xInput?.value || 0));
    transform.y = this.host.engine.pxToEmu(Number(this.yInput?.value || 0));
    transform.cx = this.host.engine.pxToEmu(Math.max(1, Number(this.widthInput?.value || 1)));
    transform.cy = this.host.engine.pxToEmu(Math.max(1, Number(this.heightInput?.value || 1)));
    transform.rot = this.host.engine.degreesToOoxml(Number(this.rotationInput?.value || 0));
    await this.host.commitTransform(transform);
    debugLog('inspector', 'Applied PowerPoint inspector transform', {
      slide: this.host.currentSlide,
      shapeIndex: this.host.selectedShapeIndex,
      x: transform.x,
      y: transform.y,
      width: transform.cx,
      height: transform.cy,
      rotation: transform.rot
    });
  }
}
