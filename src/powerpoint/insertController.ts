import { type App, Notice, type TFile } from 'obsidian';

import { isElement, isNode } from '../domGuards';
import { debugLog, errorLog } from '../logger';
import type { InsertableShapeGeometry, PresentationEngine } from '../PresentationEngine';
import {
  getImageMimeType,
  InsertTableModal,
  VaultImageSuggestModal
} from '../PowerPointInsertModals';
import type { ParagraphListStyle } from '../SlideInsertions';
import { cleanError } from './runtimeCompat';
import type { HistoryEntry, TextEditTarget } from './types';

export interface InsertHost {
  readonly engine: PresentationEngine | null;
  readonly app: App;
  readonly layoutEl: HTMLElement | null;
  readonly selectedShapeIndex: number | null;
  readonly activeEditorTarget: SVGTextElement | SVGTSpanElement | null;
  currentSlide: number;
  ensureEditable(action: string): boolean;
  captureHistoryEntry(label: string): Promise<HistoryEntry>;
  recordHistoryEntry(entry: HistoryEntry): void;
  markDirty(): void;
  renderCurrentSlide(keepSelection?: boolean): Promise<boolean>;
  renderEditedShape(shapeIndex: number): Promise<boolean>;
  renderThumbnails(): Promise<void>;
  selectShape(shapeIndex: number): void;
  createEditIconButton(container: HTMLElement, icon: string, label: string, onClick: () => void): HTMLButtonElement;
  addTextBox(): Promise<void>;
  getTextEditTarget(target: SVGTextElement | SVGTSpanElement | null): TextEditTarget | null;
  registerDocumentPointerDown(handler: (event: PointerEvent) => void, capture?: boolean): void;
  openToolbarPopover(anchor: HTMLElement, build: (popover: HTMLElement) => void): void;
  bindToolbarButton(button: HTMLElement, action: () => void): void;
  closeToolbarPopover(): void;
}

export class InsertController {
  private activeInsertMenu: HTMLElement | null = null;
  private insertTableButton: HTMLButtonElement | null = null;
  private imageFileInput: HTMLInputElement | null = null;

  constructor(private readonly host: InsertHost) {}

  registerMenus(): void {
    const closeMenus = (event: MouseEvent) => {
      const target = isNode(event.target) ? event.target : null;
      if (target && this.activeInsertMenu?.contains(target)) return;
      if (isElement(target) && target.closest('.native-powerpoint-insert-menu-anchor')) return;
      this.closeInsertMenus();
    };
    this.host.registerDocumentPointerDown(closeMenus, true);

    if (!this.host.layoutEl) return;
    const input = this.host.layoutEl.createEl('input', {
      type: 'file',
      cls: 'native-powerpoint-image-file-input'
    });
    input.accept = 'image/png,image/jpeg,image/gif,image/webp,image/bmp';
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      input.value = '';
      if (file) void this.insertImageFromLocalFile(file);
    });
    this.imageFileInput = input;
  }

  createToolbarGroup(toolbar: HTMLElement): void {
    const insertGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });

    const imageButton = this.host.createEditIconButton(insertGroup, 'image', 'Insert image', () => {
      this.toggleInsertMenu(imageButton, [
        { label: 'From vault', onClick: () => this.openVaultImagePicker() },
        { label: 'Upload file', onClick: () => this.imageFileInput?.click() }
      ]);
    });

    const shapeButton = this.host.createEditIconButton(insertGroup, 'shapes', 'Insert shape', () => {
      this.toggleInsertMenu(shapeButton, [
        { label: 'Rectangle', onClick: () => void this.insertShape('rect') },
        { label: 'Ellipse', onClick: () => void this.insertShape('ellipse') },
        { label: 'Rounded rectangle', onClick: () => void this.insertShape('roundRect') },
        { label: 'Line', onClick: () => void this.insertShape('line') },
        { label: 'Arrow', onClick: () => void this.insertShape('rightArrow') }
      ]);
    });

    this.host.createEditIconButton(insertGroup, 'type', 'Insert text box', () => void this.host.addTextBox());
    const tableButton = this.host.createEditIconButton(insertGroup, 'table', 'Insert table', () =>
      this.openTableSizePicker(tableButton)
    );
    this.insertTableButton = tableButton;
    this.host.createEditIconButton(insertGroup, 'bar-chart-3', 'Insert chart', () => void this.insertChart());
    this.host.createEditIconButton(insertGroup, 'list', 'Bulleted list', () => void this.applyListStyle('bullet'));
    this.host.createEditIconButton(insertGroup, 'list-ordered', 'Numbered list', () => void this.applyListStyle('number'));
  }

  toggleInsertMenu(
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

    const menu = activeDocument.body.createDiv({
      cls: 'native-powerpoint-insert-menu native-powerpoint-light-surface'
    });
    menu.dataset.anchorId = anchor.dataset.menuId;
    for (const item of items) {
      const button = menu.createEl('button', {
        cls: 'native-powerpoint-insert-menu-item',
        text: item.label
      });
      button.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
        this.closeInsertMenus();
        item.onClick();
      });
    }

    const rect = anchor.getBoundingClientRect();
    menu.style.left = `${rect.left}px`;
    menu.style.top = `${rect.bottom + 4}px`;
    this.activeInsertMenu = menu;
  }

  closeInsertMenus(): void {
    this.activeInsertMenu?.remove();
    this.activeInsertMenu = null;
  }

  clickImageFileInput(): void {
    this.imageFileInput?.click();
  }

  getInsertTableButton(): HTMLButtonElement | null {
    return this.insertTableButton;
  }

  openVaultImagePicker(): void {
    new VaultImageSuggestModal(this.host.app, (file) => void this.insertImageFromVaultFile(file)).open();
  }

  openInsertTableModal(): void {
    if (!this.host.ensureEditable('insert table')) return;
    new InsertTableModal(this.host.app, (rows, cols) => void this.insertTable(rows, cols)).open();
  }

  // Google Slides-style size picker: a hover grid that matches the look of the
  // other toolbar popovers (color, font) instead of a separate modal dialog.
  openTableSizePicker(anchor: HTMLElement | null): void {
    if (!this.host.ensureEditable('insert table')) return;
    if (!anchor) {
      this.openInsertTableModal();
      return;
    }

    const cols = 10;
    const rows = 8;
    this.host.openToolbarPopover(anchor, (popover) => {
      popover.addClass('native-powerpoint-table-picker');

      const grid = popover.createDiv({ cls: 'native-powerpoint-table-picker-grid' });
      const label = popover.createDiv({
        cls: 'native-powerpoint-table-picker-label',
        text: 'Insert table'
      });

      const cells: HTMLButtonElement[] = [];
      const highlight = (activeCols: number, activeRows: number): void => {
        cells.forEach((cell, index) => {
          const c = index % cols;
          const r = Math.floor(index / cols);
          cell.toggleClass('is-active', c < activeCols && r < activeRows);
        });
        label.setText(activeCols > 0 && activeRows > 0 ? `${activeCols} × ${activeRows}` : 'Insert table');
      };

      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cell = grid.createEl('button', {
            cls: 'native-powerpoint-table-picker-cell',
            attr: { 'aria-label': `${c + 1} × ${r + 1}` }
          });
          cell.addEventListener('pointerenter', () => highlight(c + 1, r + 1));
          this.host.bindToolbarButton(cell, () => {
            this.host.closeToolbarPopover();
            void this.insertTable(r + 1, c + 1);
          });
          cells.push(cell);
        }
      }

      grid.addEventListener('pointerleave', () => highlight(0, 0));
    });
  }

  async insertImageFromVaultFile(file: TFile): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('insert image')) return;

    try {
      const bytes = await this.host.app.vault.readBinary(file);
      const history = await this.host.captureHistoryEntry('Insert image');
      const shapeIndex = this.host.engine.addImage(
        this.host.currentSlide,
        new Uint8Array(bytes),
        getImageMimeType(file.extension)
      );
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        this.host.selectShape(shapeIndex);
        await this.host.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint image from vault', {
        slide: this.host.currentSlide,
        shapeIndex,
        file: file.path,
        bytes: bytes.byteLength
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint vault image insertion failed', { file: file.path, error });
      new Notice(`Could not insert image: ${cleanError(error)}`);
    }
  }

  async insertImageFromLocalFile(file: File): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('insert image')) return;

    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const history = await this.host.captureHistoryEntry('Insert image');
      const shapeIndex = this.host.engine.addImage(
        this.host.currentSlide,
        bytes,
        file.type || getImageMimeType(file.name.split('.').pop() ?? 'png')
      );
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        this.host.selectShape(shapeIndex);
        await this.host.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint image from local file', {
        slide: this.host.currentSlide,
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
      new Notice(`Could not insert image: ${cleanError(error)}`);
    }
  }

  async insertShape(geometry: InsertableShapeGeometry): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('insert shape')) return;

    try {
      const history = await this.host.captureHistoryEntry('Insert shape');
      const shapeIndex = this.host.engine.addShapeGeometry(this.host.currentSlide, geometry);
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        this.host.selectShape(shapeIndex);
        await this.host.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint shape', {
        slide: this.host.currentSlide,
        shapeIndex,
        geometry
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint shape insertion failed', { geometry, error });
      new Notice(`Could not insert shape: ${cleanError(error)}`);
    }
  }

  async insertTable(rows: number, cols: number): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('insert table')) return;

    try {
      const history = await this.host.captureHistoryEntry('Insert table');
      const shapeIndex = await this.host.engine.addTable(this.host.currentSlide, rows, cols);
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        this.host.selectShape(shapeIndex);
        await this.host.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint table', {
        slide: this.host.currentSlide,
        shapeIndex,
        rows,
        columns: cols
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint table insertion failed', { rows, columns: cols, error });
      new Notice(`Could not insert table: ${cleanError(error)}`);
    }
  }

  async insertChart(): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('insert chart')) return;

    try {
      const history = await this.host.captureHistoryEntry('Insert chart');
      const shapeIndex = await this.host.engine.addChart(this.host.currentSlide);
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        this.host.selectShape(shapeIndex);
        await this.host.renderThumbnails();
      }
      debugLog('insert', 'Inserted PowerPoint chart', {
        slide: this.host.currentSlide,
        shapeIndex
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint chart insertion failed', { slide: this.host.currentSlide, error });
      new Notice(`Could not insert chart: ${cleanError(error)}`);
    }
  }

  async applyListStyle(style: ParagraphListStyle): Promise<void> {
    if (!this.host.engine || !this.host.ensureEditable('format text')) return;

    const textTarget = this.host.getTextEditTarget(this.host.activeEditorTarget);
    const shapeIndex = textTarget?.shapeIndex ?? this.host.selectedShapeIndex;
    if (shapeIndex === null) {
      new Notice('Select a text box or place the caret in text first.');
      return;
    }

    const paragraphIndex = textTarget?.kind === 'shape-paragraph' ? textTarget.paragraphIndex : 0;
    try {
      const history = await this.host.captureHistoryEntry(
        style === 'bullet' ? 'Bulleted list' : style === 'number' ? 'Numbered list' : 'Remove list'
      );
      await this.host.engine.applyListStyle(this.host.currentSlide, shapeIndex, paragraphIndex, style);
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderEditedShape(shapeIndex);
      if (rendered) await this.host.renderThumbnails();
      debugLog('insert', 'Applied PowerPoint list style', {
        slide: this.host.currentSlide,
        shapeIndex,
        paragraphIndex,
        style
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint list-style update failed', {
        slide: this.host.currentSlide,
        shapeIndex,
        paragraphIndex,
        style,
        error
      });
      new Notice(`Could not update list style: ${cleanError(error)}`);
    }
  }
}
