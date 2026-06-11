import { App, Modal, TFile } from 'obsidian';

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
  constructor(
    app: App,
    private readonly onChoose: (file: TFile) => void
  ) {
    super(app);
  }

  onOpen(): void {
    const { contentEl } = this;
    this.modalEl.addClass('native-powerpoint-light-surface');
    contentEl.empty();
    contentEl.createEl('h2', { text: 'Insert image from vault' });

    const files = this.app.vault
      .getFiles()
      .filter(isVaultImageFile)
      .sort((left, right) => left.path.localeCompare(right.path));

    if (files.length === 0) {
      contentEl.createEl('p', { text: 'No image files found in this vault.' });
      return;
    }

    const list = contentEl.createDiv({ cls: 'native-powerpoint-image-picker-list' });
    for (const file of files) {
      const button = list.createEl('button', {
        cls: 'native-powerpoint-image-picker-item',
        text: file.path
      });
      button.addEventListener('click', () => {
        this.close();
        this.onChoose(file);
      });
    }
  }
}

export class ImageCropModal extends Modal {
  constructor(
    app: App,
    private readonly initial: ImageCropValues,
    private readonly onSubmit: (crop: ImageCropValues) => void
  ) {
    super(app);
  }

  onOpen(): void {
    const { contentEl } = this;
    this.modalEl.addClass('native-powerpoint-light-surface');
    contentEl.empty();
    contentEl.createEl('h2', { text: 'Crop image' });
    contentEl.createEl('p', {
      cls: 'native-powerpoint-field-hint',
      text: 'Set how far each edge is cropped inward, as a percentage of the image.'
    });

    const form = contentEl.createEl('form', { cls: 'native-powerpoint-insert-table-form' });
    const makeInput = (label: string, value: number): HTMLInputElement => {
      const field = form.createDiv({ cls: 'native-powerpoint-field' });
      field.createEl('label', { text: `${label} (%)` });
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

    const leftInput = makeInput('Left', this.initial.left);
    const topInput = makeInput('Top', this.initial.top);
    const rightInput = makeInput('Right', this.initial.right);
    const bottomInput = makeInput('Bottom', this.initial.bottom);

    const actions = form.createDiv({ cls: 'native-powerpoint-insert-table-actions' });
    const cancelButton = actions.createEl('button', { text: 'Cancel', type: 'button' });
    const applyButton = actions.createEl('button', {
      text: 'Apply',
      type: 'submit',
      cls: 'native-powerpoint-inspector-button'
    });

    cancelButton.addEventListener('click', () => this.close());
    form.addEventListener('submit', (event) => {
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
  }
}

export class InsertTableModal extends Modal {
  constructor(
    app: App,
    private readonly onSubmit: (rows: number, cols: number) => void
  ) {
    super(app);
  }

  onOpen(): void {
    const { contentEl } = this;
    this.modalEl.addClass('native-powerpoint-light-surface');
    contentEl.empty();
    contentEl.createEl('h2', { text: 'Insert table' });

    const form = contentEl.createEl('form', { cls: 'native-powerpoint-insert-table-form' });
    const rowsField = form.createDiv({ cls: 'native-powerpoint-field' });
    rowsField.createEl('label', { text: 'Rows' });
    const rowsInput = rowsField.createEl('input', {
      type: 'number',
      attr: { min: '1', max: '20', value: '3' }
    });

    const colsField = form.createDiv({ cls: 'native-powerpoint-field' });
    colsField.createEl('label', { text: 'Columns' });
    const colsInput = colsField.createEl('input', {
      type: 'number',
      attr: { min: '1', max: '10', value: '3' }
    });

    const actions = form.createDiv({ cls: 'native-powerpoint-insert-table-actions' });
    const cancelButton = actions.createEl('button', { text: 'Cancel', type: 'button' });
    const insertButton = actions.createEl('button', {
      text: 'Insert',
      type: 'submit',
      cls: 'native-powerpoint-inspector-button'
    });

    cancelButton.addEventListener('click', () => this.close());
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      const rows = Math.max(1, Math.min(20, Number(rowsInput.value) || 3));
      const cols = Math.max(1, Math.min(10, Number(colsInput.value) || 3));
      this.close();
      this.onSubmit(rows, cols);
    });
    insertButton.focus();
  }
}
