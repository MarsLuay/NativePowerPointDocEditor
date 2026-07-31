import { App, Component, Modal, TFile } from 'obsidian';
import type { ChartDataGrid, ChartDataUpdate } from './ChartData';
import type { TranslateFn } from './i18n/translate';
import { pptT } from './i18n/powerpointNotify';
import { createCheckboxRow } from './menuControls';
import { closeModalDomScope, loadModalDomScope, openModalDomScope } from './modalDomScope';
import {
  draftFromChartGrid,
  importChartDataFromFile,
  type ChartDataLayoutMode,
  type ChartDataTableDraft,
} from './powerpoint/chartDataImport';

const IMAGE_EXTENSIONS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'heic', 'heif']);

/** File picker accept list for PowerPoint image insert/replace (includes HEIC). */
export const POWERPOINT_IMAGE_FILE_ACCEPT =
  'image/png,image/jpeg,image/gif,image/webp,image/bmp,image/heic,image/heif,.heic,.heif';

const clampPercent = (raw: string): number => Math.max(0, Math.min(100, Number(raw) || 0));

export interface ImageCropValues {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export function getImageMimeType(extension: string): string {
  switch (extension.toLowerCase()) {
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'webp':
      return 'image/webp';
    case 'bmp':
      return 'image/bmp';
    case 'heic':
    case 'heics':
      return 'image/heic';
    case 'heif':
    case 'heifs':
      return 'image/heif';
    default:
      return 'image/png';
  }
}

export function isVaultImageFile(file: TFile): boolean {
  return IMAGE_EXTENSIONS.has(file.extension.toLowerCase());
}

export class VaultImageSuggestModal extends Modal {
  private domScope?: Component;

  constructor(
    app: App,
    private readonly onChoose: (file: TFile) => void,
    private readonly t: TranslateFn = pptT
  ) {
    super(app);
  }

  onOpen(): void {
    this.domScope = openModalDomScope();
    const { contentEl } = this;
    this.modalEl.addClass('native-powerpoint-light-surface');
    contentEl.empty();
    contentEl.createEl('h2', { text: this.t('powerpoint:modal.insertImageFromVault') });

    const files = this.app.vault
      .getFiles()
      .filter(isVaultImageFile)
      .sort((left, right) => left.path.localeCompare(right.path));

    if (files.length === 0) {
      contentEl.createEl('p', { text: this.t('powerpoint:modal.noVaultImages') });
      loadModalDomScope(this.domScope);
      return;
    }

    const list = contentEl.createDiv({ cls: 'native-powerpoint-image-picker-list' });
    for (const file of files) {
      const button = list.createEl('button', {
        cls: 'native-powerpoint-image-picker-item',
        text: file.path
      });
      this.domScope.registerDomEvent(button, 'click', () => {
        this.close();
        this.onChoose(file);
      });
    }

    loadModalDomScope(this.domScope);
  }

  onClose(): void {
    closeModalDomScope(this.domScope);
    this.domScope = undefined;
  }
}

export class ImageCropModal extends Modal {
  private domScope?: Component;

  constructor(
    app: App,
    private readonly initial: ImageCropValues,
    private readonly onSubmit: (crop: ImageCropValues) => void,
    private readonly t: TranslateFn = pptT
  ) {
    super(app);
  }

  onOpen(): void {
    this.domScope = openModalDomScope();
    const { contentEl } = this;
    this.modalEl.addClass('native-powerpoint-light-surface');
    contentEl.empty();
    contentEl.createEl('h2', { text: this.t('powerpoint:modal.cropImage') });
    contentEl.createEl('p', {
      cls: 'native-powerpoint-field-hint',
      text: this.t('powerpoint:modal.cropImageHint')
    });

    const form = contentEl.createEl('form', { cls: 'native-powerpoint-insert-table-form' });
    const makeInput = (labelKey: string, value: number): HTMLInputElement => {
      const field = form.createDiv({ cls: 'native-powerpoint-field' });
      field.createEl('label', { text: `${this.t(labelKey)}${this.t('powerpoint:modal.percentSuffix')}` });
      return field.createEl('input', {
        type: 'number',
        attr: {
          min: '0',
          max: '100',
          step: '0.1',
          value: String(Math.round(value * 100) / 100)
        }
      });
    };

    const leftInput = makeInput('powerpoint:modal.cropLeft', this.initial.left);
    const topInput = makeInput('powerpoint:modal.cropTop', this.initial.top);
    const rightInput = makeInput('powerpoint:modal.cropRight', this.initial.right);
    const bottomInput = makeInput('powerpoint:modal.cropBottom', this.initial.bottom);

    const actions = form.createDiv({ cls: 'native-powerpoint-insert-table-actions' });
    const cancelButton = actions.createEl('button', { text: this.t('common:actions.cancel'), type: 'button' });
    const applyButton = actions.createEl('button', {
      text: this.t('powerpoint:inspector.apply'),
      type: 'submit',
      cls: 'native-powerpoint-inspector-button'
    });

    this.domScope.registerDomEvent(cancelButton, 'click', () => this.close());
    this.domScope.registerDomEvent(form, 'submit', (event) => {
      event.preventDefault();
      this.close();
      this.onSubmit({
        left: clampPercent(leftInput.value),
        top: clampPercent(topInput.value),
        right: clampPercent(rightInput.value),
        bottom: clampPercent(bottomInput.value)
      });
    });
    applyButton.focus();
    loadModalDomScope(this.domScope);
  }

  onClose(): void {
    closeModalDomScope(this.domScope);
    this.domScope = undefined;
  }
}

export class InsertTableModal extends Modal {
  private domScope?: Component;

  constructor(
    app: App,
    private readonly onSubmit: (rows: number, cols: number) => void,
    private readonly t: TranslateFn = pptT
  ) {
    super(app);
  }

  onOpen(): void {
    this.domScope = openModalDomScope();
    const { contentEl } = this;
    this.modalEl.addClass('native-powerpoint-light-surface');
    contentEl.empty();
    contentEl.createEl('h2', { text: this.t('powerpoint:modal.insertTable') });

    const form = contentEl.createEl('form', { cls: 'native-powerpoint-insert-table-form' });
    const rowsField = form.createDiv({ cls: 'native-powerpoint-field' });
    rowsField.createEl('label', { text: this.t('powerpoint:modal.rows') });
    const rowsInput = rowsField.createEl('input', {
      type: 'number',
      attr: { min: '1', max: '20', value: '3' }
    });

    const colsField = form.createDiv({ cls: 'native-powerpoint-field' });
    colsField.createEl('label', { text: this.t('powerpoint:modal.columns') });
    const colsInput = colsField.createEl('input', {
      type: 'number',
      attr: { min: '1', max: '10', value: '3' }
    });

    const actions = form.createDiv({ cls: 'native-powerpoint-insert-table-actions' });
    const cancelButton = actions.createEl('button', { text: this.t('common:actions.cancel'), type: 'button' });
    const insertButton = actions.createEl('button', {
      text: this.t('common:actions.insert'),
      type: 'submit',
      cls: 'native-powerpoint-inspector-button'
    });

    this.domScope.registerDomEvent(cancelButton, 'click', () => this.close());
    this.domScope.registerDomEvent(form, 'submit', (event) => {
      event.preventDefault();
      const rows = Math.max(1, Math.min(20, Number(rowsInput.value) || 3));
      const cols = Math.max(1, Math.min(10, Number(colsInput.value) || 3));
      this.close();
      this.onSubmit(rows, cols);
    });
    insertButton.focus();
    loadModalDomScope(this.domScope);
  }

  onClose(): void {
    closeModalDomScope(this.domScope);
    this.domScope = undefined;
  }
}

/** Right-click / Edit Data screen for a chart's category and series grid. */
export class ChartDataEditModal extends Modal {
  private domScope?: Component;
  private layoutMode: ChartDataLayoutMode = 'chart';
  private draft: ChartDataTableDraft;
  private formEl: HTMLFormElement | null = null;
  private tableHost: HTMLElement | null = null;
  private statusEl: HTMLElement | null = null;
  private categoryInputs: HTMLInputElement[] = [];
  private seriesNameInputs: HTMLInputElement[] = [];
  private valueInputs: HTMLInputElement[][] = [];
  private pointLabelInputs: HTMLInputElement[][] = [];
  private fileInput: HTMLInputElement | null = null;

  constructor(
    app: App,
    private readonly chartData: ChartDataGrid,
    private readonly options: {
      canEdit: boolean;
      onSubmit: (update: ChartDataUpdate) => void;
    },
    private readonly t: TranslateFn = pptT,
  ) {
    super(app);
    this.draft = draftFromChartGrid(chartData);
  }

  onOpen(): void {
    this.domScope = openModalDomScope();
    const { contentEl } = this;
    this.modalEl.addClass('native-powerpoint-light-surface');
    this.modalEl.addClass('native-powerpoint-chart-data-modal');
    contentEl.empty();
    contentEl.createEl('h2', { text: this.t('powerpoint:modal.editChartData') });

    if (!this.chartData.editable) {
      contentEl.createEl('p', {
        cls: 'native-powerpoint-field-hint',
        text: this.chartData.reason || this.t('powerpoint:inspector.chartDataReadOnly'),
      });
      const closeButton = contentEl.createEl('button', {
        text: this.t('common:actions.close'),
        type: 'button',
        cls: 'native-powerpoint-inspector-button',
      });
      this.domScope.registerDomEvent(closeButton, 'click', () => this.close());
      loadModalDomScope(this.domScope);
      return;
    }

    contentEl.createEl('p', {
      cls: 'native-powerpoint-field-hint',
      text: this.t('powerpoint:modal.editChartDataHint'),
    });

    const toolbar = contentEl.createDiv({ cls: 'native-powerpoint-chart-data-toolbar' });
    const importButton = toolbar.createEl('button', {
      text: this.t('powerpoint:modal.importExcel'),
      type: 'button',
      cls: 'native-powerpoint-inspector-button',
    });
    importButton.disabled = !this.options.canEdit;

    this.fileInput = toolbar.createEl('input', {
      type: 'file',
      cls: 'native-powerpoint-chart-data-file-input',
      attr: {
        accept: '.xlsx,.xlsm,.csv,.tsv,.txt,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv,text/tab-separated-values',
        'aria-hidden': 'true',
        tabindex: '-1',
      },
    });

    const layoutToggle = createCheckboxRow(toolbar, {
      checked: this.layoutMode === 'excel',
      controlClassName: 'native-powerpoint-chart-data-layout-control',
      copyClassName: 'native-powerpoint-chart-data-layout-copy',
      descriptionClassName: 'native-powerpoint-chart-data-layout-description',
      inputClassName: 'native-powerpoint-chart-data-layout-input',
      label: this.t('powerpoint:modal.excelDataLayout'),
      labelClassName: 'native-powerpoint-chart-data-layout-label',
      rowClassName: 'native-powerpoint-chart-data-layout-row',
      onChange: (checked) => {
        this.commitDraftFromInputs();
        this.layoutMode = checked ? 'excel' : 'chart';
        this.renderTable();
      },
    });
    layoutToggle.disabled = !this.options.canEdit;

    this.statusEl = contentEl.createDiv({ cls: 'native-powerpoint-field-hint native-powerpoint-chart-data-status' });
    this.statusEl.setText(this.layoutHint());

    this.formEl = contentEl.createEl('form', { cls: 'native-powerpoint-chart-data-form' });
    const viewport = this.formEl.createDiv({ cls: 'native-powerpoint-chart-data-scroll' });
    this.tableHost = viewport;
    this.renderTable();

    const actions = this.formEl.createDiv({ cls: 'native-powerpoint-insert-table-actions' });
    const cancelButton = actions.createEl('button', {
      text: this.t('common:actions.cancel'),
      type: 'button',
    });
    const applyButton = actions.createEl('button', {
      text: this.t('powerpoint:inspector.applyChartData'),
      type: 'submit',
      cls: 'native-powerpoint-inspector-button',
    });
    applyButton.disabled = !this.options.canEdit;

    this.domScope.registerDomEvent(importButton, 'click', () => this.fileInput?.click());
    this.domScope.registerDomEvent(this.fileInput, 'change', () => {
      void this.handleImportFile();
    });
    this.domScope.registerDomEvent(cancelButton, 'click', () => this.close());
    this.domScope.registerDomEvent(this.formEl, 'submit', (event) => {
      event.preventDefault();
      if (!this.options.canEdit) return;
      this.commitDraftFromInputs();
      if (this.draft.series.length !== this.chartData.series.length) {
        this.setStatus(this.t('powerpoint:modal.importSeriesCountMismatch', {
          expected: String(this.chartData.series.length),
          actual: String(this.draft.series.length),
        }), true);
        return;
      }
      this.close();
      this.options.onSubmit({
        categories: this.draft.categories,
        series: this.chartData.series.map((series, index) => ({
          values: this.draft.series[index]?.values ?? [],
          pointLabels: series.pointLabels === null
            ? null
            : this.draft.series[index]?.pointLabels ?? series.pointLabels.map(() => ''),
        })),
      });
    });
    applyButton.focus();
    loadModalDomScope(this.domScope);
  }

  onClose(): void {
    closeModalDomScope(this.domScope);
    this.domScope = undefined;
    this.formEl = null;
    this.tableHost = null;
    this.statusEl = null;
    this.fileInput = null;
  }

  private layoutHint(): string {
    return this.layoutMode === 'excel'
      ? this.t('powerpoint:modal.excelDataLayoutHint')
      : this.t('powerpoint:modal.chartDataLayoutHint');
  }

  private setStatus(message: string, isError = false): void {
    if (!this.statusEl) return;
    this.statusEl.setText(message);
    this.statusEl.toggleClass('is-error', isError);
  }

  private commitDraftFromInputs(): void {
    if (!this.tableHost) return;

    if (this.layoutMode === 'excel') {
      const categories = this.categoryInputs.map((input) => input.value);
      this.draft = {
        categoryLabel: this.draft.categoryLabel,
        categories,
        series: this.seriesNameInputs.map((nameInput, seriesIndex) => {
          const previous = this.draft.series[seriesIndex];
          return {
            name: nameInput.value.trim() || `Series ${seriesIndex + 1}`,
            pointLabels: previous?.pointLabels === null || previous?.pointLabels === undefined
              ? null
              : categories.map((_, rowIndex) => previous.pointLabels?.[rowIndex] ?? ''),
            values: categories.map((_, rowIndex) => this.valueInputs[seriesIndex]?.[rowIndex]?.value ?? ''),
          };
        }),
      };
      return;
    }

    this.draft = {
      categoryLabel: this.draft.categoryLabel,
      categories: this.categoryInputs.map((input) => input.value),
      series: this.draft.series.map((series, seriesIndex) => ({
        name: series.name,
        pointLabels: series.pointLabels === null
          ? null
          : this.pointLabelInputs[seriesIndex]?.map((input) => input.value) ?? [],
        values: this.valueInputs[seriesIndex]?.map((input) => input.value) ?? [],
      })),
    };
  }

  private renderTable(): void {
    if (!this.tableHost) return;
    this.tableHost.empty();
    this.categoryInputs = [];
    this.seriesNameInputs = [];
    this.valueInputs = this.draft.series.map(() => [] as HTMLInputElement[]);
    this.pointLabelInputs = this.draft.series.map(() => [] as HTMLInputElement[]);

    const table = this.tableHost.createEl('table', {
      cls: [
        'native-powerpoint-chart-data-grid',
        this.layoutMode === 'excel'
          ? 'native-powerpoint-chart-data-grid-excel'
          : 'native-powerpoint-chart-data-grid-chart',
      ],
    });

    if (this.layoutMode === 'excel') {
      this.renderExcelTable(table);
    } else {
      this.renderChartTable(table);
    }

    this.setStatus(this.layoutHint());
  }

  private renderChartTable(table: HTMLTableElement): void {
    const header = table.createEl('thead').createEl('tr');
    header.createEl('th', { text: this.draft.categoryLabel });
    this.draft.series.forEach((series) => {
      header.createEl('th', { text: series.name });
      if (series.pointLabels !== null) {
        header.createEl('th', {
          text: this.t('powerpoint:inspector.seriesLabel', { seriesName: series.name }),
        });
      }
    });

    const body = table.createEl('tbody');
    this.draft.categories.forEach((category, rowIndex) => {
      const row = body.createEl('tr');
      this.categoryInputs.push(this.createInput(row, category, false));
      this.draft.series.forEach((series, seriesIndex) => {
        this.valueInputs[seriesIndex]?.push(this.createInput(row, series.values[rowIndex] ?? '', true));
        if (series.pointLabels !== null) {
          this.pointLabelInputs[seriesIndex]?.push(
            this.createInput(row, series.pointLabels[rowIndex] ?? '', false),
          );
        }
      });
    });
  }

  private renderExcelTable(table: HTMLTableElement): void {
    const header = table.createEl('thead').createEl('tr');
    // Excel datasheet: blank corner cell, then editable series names.
    header.createEl('th', { cls: 'native-powerpoint-chart-data-corner' }).createSpan({ text: '' });
    this.draft.series.forEach((series, seriesIndex) => {
      const cell = header.createEl('th');
      const input = cell.createEl('input', {
        type: 'text',
        value: series.name,
        attr: { 'aria-label': this.t('powerpoint:modal.seriesName', { index: String(seriesIndex + 1) }) },
      });
      input.disabled = !this.options.canEdit;
      this.seriesNameInputs.push(input);
    });

    const body = table.createEl('tbody');
    this.draft.categories.forEach((category, rowIndex) => {
      const row = body.createEl('tr');
      this.categoryInputs.push(this.createInput(row, category, false));
      this.draft.series.forEach((series, seriesIndex) => {
        this.valueInputs[seriesIndex]?.push(this.createInput(row, series.values[rowIndex] ?? '', true));
      });
    });
  }

  private async handleImportFile(): Promise<void> {
    const file = this.fileInput?.files?.[0];
    if (this.fileInput) this.fileInput.value = '';
    if (!file || !this.options.canEdit) return;

    try {
      this.commitDraftFromInputs();
      const imported = await importChartDataFromFile(file, {
        defaultCategoryLabel: this.chartData.categoryLabel,
        preferLayout: this.layoutMode,
      });

      if (imported.series.length !== this.chartData.series.length) {
        throw new Error(
          this.t('powerpoint:modal.importSeriesCountMismatch', {
            expected: String(this.chartData.series.length),
            actual: String(imported.series.length),
          })
        );
      }

      this.draft = {
        categoryLabel: this.chartData.categoryLabel,
        categories: imported.categories,
        series: this.chartData.series.map((series, index) => ({
          name: imported.series[index]?.name || series.name,
          pointLabels: series.pointLabels === null
            ? null
            : imported.categories.map((_, rowIndex) => series.pointLabels?.[rowIndex] ?? ''),
          values: imported.series[index]?.values ?? [],
        })),
      };
      this.renderTable();
      this.setStatus(this.t('powerpoint:modal.importExcelSuccess', { fileName: file.name }));
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.setStatus(message, true);
    }
  }

  private createInput(row: HTMLTableRowElement, value: string, numeric: boolean): HTMLInputElement {
    const input = row.createEl('td').createEl('input', {
      type: 'text',
      value,
      attr: numeric ? { inputmode: 'decimal' } : {},
    });
    input.disabled = !this.options.canEdit;
    return input;
  }
}

