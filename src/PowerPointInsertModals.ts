import { App, Component, Modal, TFile } from 'obsidian';
import type { TranslateFn } from './i18n/translate';
import { pptT } from './i18n/powerpointNotify';
import { closeModalDomScope, loadModalDomScope, openModalDomScope } from './modalDomScope';

const IMAGE_EXTENSIONS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp']);

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
