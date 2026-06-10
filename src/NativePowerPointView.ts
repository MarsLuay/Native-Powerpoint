import { FileView, Notice, Platform, TFile, WorkspaceLeaf, normalizePath, setIcon } from 'obsidian';

import {
  PresentationEngine,
  type GeneratedTextEdit,
  type GeneratedTextKind
} from './PresentationEngine';
import {
  inspectPowerPointPackage,
  summarizePackageMessages,
  validatePowerPointExport,
  validatePowerPointExportContents,
  validatePowerPointPackageStructure,
  type PowerPointPackageInspection
} from './PowerPointPackage';
import { createSvgElementFromString, sanitizeSvg, summarizeSvgSecurityIssues, type SvgSecurityIssue } from './SvgSecurity';
import { applyBackgroundAwareTextHalos } from './TextHalo';
import type { NativePowerPointSettings } from './settings';
import type { ShapeTransform } from 'pptx-svg';
import type { ChartDataGrid, ChartDataUpdate } from './ChartData';
import type { FontSubstitution } from './FontFidelity';
import type { SlideObjectClipboard } from './ShapeClipboard';

export const NATIVE_POWERPOINT_VIEW_TYPE = 'native-powerpoint-view';

export const MODERN_POWERPOINT_EXTENSIONS = [
  'pptx',
  'pptm',
  'ppsx',
  'ppsm',
  'potx',
  'potm'
];

export const LEGACY_POWERPOINT_EXTENSIONS = ['ppt', 'pps', 'pot'];
export const MACRO_ENABLED_POWERPOINT_EXTENSIONS = ['pptm', 'ppsm', 'potm'];
export const EDITABLE_POWERPOINT_EXTENSIONS = ['pptx', 'ppsx', 'potx'];

export const POWERPOINT_EXTENSIONS = [
  ...MODERN_POWERPOINT_EXTENSIONS,
  ...LEGACY_POWERPOINT_EXTENSIONS
];

type SaveState = 'idle' | 'dirty' | 'saving' | 'saved' | 'failed' | 'recovered' | 'view-only';
type HandleName = 'nw' | 'ne' | 'sw' | 'se';
type SvgSecurityDecision = 'compatibility' | 'yolo' | null;

interface PointerPoint {
  x: number;
  y: number;
}

interface DragState {
  mode: 'move' | 'resize';
  handle?: HandleName;
  pointerId: number;
  startPoint: PointerPoint;
  startClientX: number;
  startClientY: number;
  startBox: { left: number; top: number; width: number; height: number };
  startTransform: ShapeTransform;
  latestTransform: ShapeTransform;
}

interface PowerPointFindMatch {
  slideIndex: number;
  shapeIndex: number | null;
  text: string;
}

interface SlideSize {
  width: number;
  height: number;
}

interface InlineCaretRow {
  top: number;
  height: number;
  centerRatio: number;
}

interface SvgInlineCaretGeometry {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  strokeWidth: number;
}

interface SvgInlineSelectionBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface InlineSelectionDrag {
  editor: HTMLTextAreaElement;
  element: SVGTextElement | SVGTSpanElement;
  anchorOffset: number;
  pointerId: number;
  startClientX: number;
  startClientY: number;
  isSelecting: boolean;
  cleanup: () => void;
}

interface InlineCaretPlacement {
  editor: HTMLTextAreaElement;
  element: SVGTextElement | SVGTSpanElement;
  offset: number;
  timestamp: number;
}

interface CanvasScrollPosition {
  left: number;
  top: number;
}

interface HistoryEntry {
  buffer: ArrayBuffer;
  currentSlide: number;
  label: string;
}

interface ShapeTextEditTarget {
  kind: 'shape-run';
  shapeIndex: number;
  paragraphIndex: number;
  runIndex: number;
  text: string;
  element: SVGTSpanElement;
}

interface GeneratedTextEditTarget extends GeneratedTextEdit {
  shapeIndex: number;
  text: string;
  element: SVGTextElement;
}

type TextEditTarget = GeneratedTextEditTarget | ShapeTextEditTarget;

const GENERATED_GRID_SELECTOR =
  'g[data-ooxml-shape-type="table"], g[data-ooxml-shape-type="chart"]';
const HISTORY_LIMIT = 20;
const MIN_ZOOM = 0.25;
const MAX_ZOOM = 3;

export function isPowerPointExtension(extension: string): boolean {
  return POWERPOINT_EXTENSIONS.includes(extension.toLowerCase());
}

export function isModernPowerPointExtension(extension: string): boolean {
  return MODERN_POWERPOINT_EXTENSIONS.includes(extension.toLowerCase());
}

export function isEditablePowerPointExtension(extension: string): boolean {
  return EDITABLE_POWERPOINT_EXTENSIONS.includes(extension.toLowerCase());
}

export function isMacroEnabledPowerPointExtension(extension: string): boolean {
  return MACRO_ENABLED_POWERPOINT_EXTENSIONS.includes(extension.toLowerCase());
}

function cleanError(error: unknown): string {
  return error instanceof Error ? error.message : String(error || 'Unknown error');
}

function cloneTransform(transform: ShapeTransform): ShapeTransform {
  return {
    x: transform.x,
    y: transform.y,
    cx: transform.cx,
    cy: transform.cy,
    rot: transform.rot
  };
}

function getShapeIndex(shape: Element | null): number | null {
  const raw = shape?.getAttribute('data-ooxml-shape-idx');
  if (!raw) return null;

  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

function isPrimaryFindShortcut(evt: KeyboardEvent): boolean {
  const key = evt.key.toLowerCase();
  const isMacFind = evt.metaKey && !evt.ctrlKey;
  const isNonMacFind = evt.ctrlKey && !evt.metaKey && !Platform.isMacOS;
  const hasPrimaryModifier = isMacFind || isNonMacFind;
  return key === 'f' && hasPrimaryModifier && !evt.altKey && !evt.shiftKey;
}

function normalizeSearchText(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

function parseSvgLength(value: string | null): number | null {
  if (!value || value.includes('%')) return null;

  const match = value.trim().match(/^([0-9]*\.?[0-9]+)(?:px)?$/i);
  if (!match) return null;

  const parsed = Number(match[1]);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
}

function getSvgIntrinsicSize(svg: SVGSVGElement): SlideSize | null {
  const viewBox = svg.viewBox.baseVal;
  if (viewBox.width > 0 && viewBox.height > 0) {
    return { width: viewBox.width, height: viewBox.height };
  }

  const width = parseSvgLength(svg.getAttribute('width'));
  const height = parseSvgLength(svg.getAttribute('height'));
  if (width && height) {
    return { width, height };
  }

  return null;
}

function transformsMatch(a: ShapeTransform, b: ShapeTransform): boolean {
  return a.x === b.x && a.y === b.y && a.cx === b.cx && a.cy === b.cy && a.rot === b.rot;
}

function ensureSvgViewBox(svg: SVGSVGElement): void {
  const viewBox = svg.viewBox.baseVal;
  if (viewBox.width > 0 && viewBox.height > 0) return;

  const width = parseSvgLength(svg.getAttribute('width'));
  const height = parseSvgLength(svg.getAttribute('height'));
  if (!width || !height) return;

  svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
  if (!svg.hasAttribute('preserveAspectRatio')) {
    svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
  }
}

function bringGridTextToFront(svg: SVGSVGElement): void {
  const gridGroups = svg.querySelectorAll(GENERATED_GRID_SELECTOR);

  for (const gridGroup of Array.from(gridGroups)) {
    const kind = gridGroup.getAttribute('data-ooxml-shape-type') as GeneratedTextKind;
    const occurrences = new Map<string, number>();
    applyBackgroundAwareTextHalos(gridGroup, kind);
    gridGroup.querySelectorAll('text').forEach((text, labelIndex) => {
      const normalizedText = normalizeSearchText(text.textContent || '');
      const occurrence = occurrences.get(normalizedText) ?? 0;
      occurrences.set(normalizedText, occurrence + 1);

      text.classList.add('native-powerpoint-grid-label');
      text.setAttribute('data-native-powerpoint-generated-kind', kind);
      text.setAttribute('data-native-powerpoint-label-index', String(labelIndex));
      text.setAttribute('data-native-powerpoint-label-occurrence', String(occurrence));
    });

    const containers = [gridGroup, ...Array.from(gridGroup.querySelectorAll('g'))];
    for (const container of containers) {
      const textChildren = Array.from(container.children).filter(
        (child) => child.tagName.toLowerCase() === 'text'
      );
      for (const textChild of textChildren) {
        container.appendChild(textChild);
      }
    }
  }
}

function markEditableTextRuns(svg: SVGSVGElement): void {
  svg.querySelectorAll('tspan[data-ooxml-run-idx]').forEach((run) => {
    if (!run.closest(GENERATED_GRID_SELECTOR)) {
      run.classList.add('native-powerpoint-editable-text');
    }
  });
}

function normalizeSvgForDisplay(svg: SVGSVGElement): void {
  ensureSvgViewBox(svg);
  bringGridTextToFront(svg);
  markEditableTextRuns(svg);
}

export class NativePowerPointView extends FileView {
  private readonly getSettings: () => NativePowerPointSettings;

  private engine: PresentationEngine | null = null;
  private loadedFile: TFile | null = null;
  private sourcePackage: PowerPointPackageInspection | null = null;
  private sourceBuffer: ArrayBuffer | null = null;
  private currentSlide = 0;
  private zoomLevel = 1;
  private selectedShapeIndex: number | null = null;
  private selectedTransform: ShapeTransform | null = null;
  private saveState: SaveState = 'idle';
  private isViewOnly = false;
  private viewOnlyReason = '';
  private isLoading = false;
  private isDirty = false;
  private editVersion = 0;
  private saveTimer: number | null = null;
  private savePromise: Promise<void> = Promise.resolve();
  private dragState: DragState | null = null;
  private activeEditor: HTMLTextAreaElement | null = null;
  private activeEditorCommit: (() => Promise<void>) | null = null;
  private activeInlineCaret: SVGLineElement | null = null;
  private activeInlineSelectionRects: SVGRectElement[] = [];
  private inlineSelectionDrag: InlineSelectionDrag | null = null;
  private lastInlineCaretPlacement: InlineCaretPlacement | null = null;
  private suppressNextTextClick = false;
  private activeInlineCaretRow: InlineCaretRow | null = null;
  private activeEditorTarget: SVGTextElement | SVGTSpanElement | null = null;
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private isRestoringHistory = false;
  private svgSecurityDecision: SvgSecurityDecision = null;
  private findMatches: PowerPointFindMatch[] = [];
  private currentFindMatchIndex = 0;

  private layoutEl: HTMLElement | null = null;
  private thumbnailContainer: HTMLElement | null = null;
  private canvasPane: HTMLElement | null = null;
  private slideSurface: HTMLElement | null = null;
  private inspectorEl: HTMLElement | null = null;
  private statusEl: HTMLElement | null = null;
  private slideCounterEl: HTMLElement | null = null;
  private svgEl: SVGSVGElement | null = null;
  private selectionOverlay: HTMLElement | null = null;
  private canvasResizeObserver: ResizeObserver | null = null;
  private findPanelEl: HTMLElement | null = null;
  private findInputEl: HTMLInputElement | null = null;
  private findStatusEl: HTMLElement | null = null;
  private editButtons: HTMLButtonElement[] = [];
  private undoButton: HTMLButtonElement | null = null;
  private redoButton: HTMLButtonElement | null = null;
  private xInput: HTMLInputElement | null = null;
  private yInput: HTMLInputElement | null = null;
  private widthInput: HTMLInputElement | null = null;
  private heightInput: HTMLInputElement | null = null;
  private rotationInput: HTMLInputElement | null = null;
  private hasShownGeneratedTextNotice = false;
  private fontSubstitutions: FontSubstitution[] = [];
  private objectClipboard: SlideObjectClipboard | null = null;
  private copyButton: HTMLButtonElement | null = null;
  private pasteButton: HTMLButtonElement | null = null;
  private duplicateButton: HTMLButtonElement | null = null;

  constructor(leaf: WorkspaceLeaf, getSettings: () => NativePowerPointSettings) {
    super(leaf);
    this.getSettings = getSettings;
  }

  canAcceptExtension(extension: string): boolean {
    return isPowerPointExtension(extension);
  }

  getViewType(): string {
    return NATIVE_POWERPOINT_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.loadedFile?.basename || this.file?.basename || 'Native PowerPoint';
  }

  getIcon(): string {
    return 'presentation';
  }

  async onOpen(): Promise<void> {
    this.contentEl.empty();
    this.contentEl.addClass('native-powerpoint-view');
    this.createLayout();
    this.registerKeyboardHandlers();
    this.renderInspector();
  }

  async onLoadFile(file: TFile): Promise<void> {
    if (this.engine && this.loadedFile && this.loadedFile.path !== file.path) {
      const preserved = await this.preserveUnsavedChangesForTeardown('switching files');
      if (!preserved) {
        new Notice(`Could not switch files because unsaved edits from ${this.loadedFile.name} could not be preserved.`);
        return;
      }

      this.resetLoadedPresentation();
    }

    await this.loadPresentation(file);
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    const preserved = await this.preserveUnsavedChangesForTeardown('switching files');
    if (preserved) {
      this.resetLoadedPresentation();
    }
  }

  async onClose(): Promise<void> {
    const preserved = await this.preserveUnsavedChangesForTeardown('closing the view');
    if (!preserved) {
      new Notice('Native PowerPoint could not close safely. Keep this tab open and retry saving or closing after resolving the vault write error.');
      return;
    }

    this.resetLoadedPresentation();
    this.contentEl.removeClass('native-powerpoint-view');
    this.file = null;
  }

  async saveCurrentPresentation(): Promise<boolean> {
    const file = this.loadedFile || this.file;
    const engine = this.engine;
    const sourcePackage = this.sourcePackage;
    const sourceBuffer = this.sourceBuffer;
    if (!file || !engine || !sourcePackage || !sourceBuffer || !isModernPowerPointExtension(file.extension)) {
      new Notice('Open a modern PowerPoint file to save it.');
      return false;
    }

    if (!isEditablePowerPointExtension(file.extension)) {
      new Notice(this.viewOnlyReason || 'This PowerPoint format is view-only in Native PowerPoint.');
      return false;
    }

    if (!this.ensureEditable('save changes')) {
      return false;
    }

    const targetVersion = this.editVersion;
    this.clearAutosave();
    this.setSaveState('saving');

    const run = async () => {
      const output = await engine.export();
      const exportedPackage = await this.validateExportBeforeSave(output, engine, sourcePackage, sourceBuffer);
      await this.app.vault.modifyBinary(file, output);

      if (this.engine === engine && this.loadedFile?.path === file.path) {
        this.sourcePackage = exportedPackage;
        this.sourceBuffer = output;

        if (this.editVersion === targetVersion) {
          this.isDirty = false;
          this.setSaveState('saved');
        } else {
          this.setSaveState('dirty');
          this.scheduleAutosave();
        }
      }
    };

    this.savePromise = this.savePromise.then(run, run);

    try {
      await this.savePromise;
      return true;
    } catch (error) {
      this.setSaveState('failed');
      new Notice(`Could not save ${file.name}: ${cleanError(error)}`);
      return false;
    }
  }

  private async validateExportBeforeSave(
    output: ArrayBuffer,
    engine = this.engine,
    sourcePackage = this.sourcePackage,
    sourceBuffer = this.sourceBuffer
  ): Promise<PowerPointPackageInspection> {
    if (!engine || !sourcePackage || !sourceBuffer) {
      throw new Error('Cannot verify the PowerPoint package before saving.');
    }

    const exportedPackage = inspectPowerPointPackage(output);
    const validation = validatePowerPointExport(sourcePackage, exportedPackage, engine.slideCount);
    if (!validation.ok) {
      throw new Error(`Export validation failed: ${summarizePackageMessages(validation.errors)}`);
    }

    const contentValidation = await validatePowerPointExportContents(sourceBuffer, output);
    if (!contentValidation.ok) {
      throw new Error(`Export validation failed: ${summarizePackageMessages(contentValidation.errors)}`);
    }

    await PresentationEngine.validateRoundTrip(output, engine.slideCount);
    return exportedPackage;
  }

  private canEdit(): boolean {
    const file = this.loadedFile || this.file;
    return Boolean(
      this.engine &&
      file &&
      !this.isViewOnly &&
      isEditablePowerPointExtension(file.extension)
    );
  }

  private ensureEditable(action: string): boolean {
    if (this.canEdit()) return true;

    const reason = this.viewOnlyReason || 'This PowerPoint file is view-only in Native PowerPoint.';
    new Notice(`Cannot ${action}: ${reason}`);
    return false;
  }

  private shouldOpenViewOnly(file: TFile, sourcePackage: PowerPointPackageInspection): boolean {
    return isMacroEnabledPowerPointExtension(file.extension) || sourcePackage.hasVbaProject;
  }

  private getViewOnlyReason(file: TFile, sourcePackage: PowerPointPackageInspection): string {
    if (isMacroEnabledPowerPointExtension(file.extension)) {
      return 'Macro-enabled PowerPoint files are view-only until macro preservation is verified.';
    }

    if (sourcePackage.hasVbaProject) {
      return 'This package contains a macro project, so editing is disabled until macro preservation is verified.';
    }

    return '';
  }

  private createLayout(): void {
    this.contentEl.empty();

    const root = this.contentEl.createDiv({ cls: 'native-powerpoint-root' });
    this.layoutEl = root.createDiv({ cls: 'native-powerpoint-layout' });

    const sidebar = this.layoutEl.createDiv({ cls: 'native-powerpoint-sidebar' });
    const sidebarHeader = sidebar.createDiv({ cls: 'native-powerpoint-sidebar-header', text: 'Slides' });
    sidebarHeader.createSpan({ cls: 'native-powerpoint-sidebar-hint', text: 'Native' });
    this.thumbnailContainer = sidebar.createDiv({ cls: 'native-powerpoint-thumbnails' });

    const main = this.layoutEl.createDiv({ cls: 'native-powerpoint-main-content' });
    this.createToolbar(main);
    this.canvasPane = main.createDiv({ cls: 'native-powerpoint-canvas-pane' });
    this.slideSurface = this.canvasPane.createDiv({ cls: 'native-powerpoint-slide-surface' });
    this.registerDomEvent(this.canvasPane, 'pointerdown', this.handleCanvasPanePointerDown, true);
    this.registerCanvasWheelZoom();
    this.observeCanvasPane();

    this.inspectorEl = this.layoutEl.createDiv({ cls: 'native-powerpoint-inspector' });
  }

  private observeCanvasPane(): void {
    this.canvasResizeObserver?.disconnect();
    this.canvasResizeObserver = null;

    if (!this.canvasPane || typeof ResizeObserver === 'undefined') return;

    const observer = new ResizeObserver(() => this.updateSlideScale());
    observer.observe(this.canvasPane);
    this.canvasResizeObserver = observer;
    this.register(() => observer.disconnect());
  }

  private registerCanvasWheelZoom(): void {
    if (!this.canvasPane) return;

    const pane = this.canvasPane;
    const handleWheel = (event: WheelEvent) => this.handleCanvasWheel(event);
    pane.addEventListener('wheel', handleWheel, { passive: false });
    this.register(() => pane.removeEventListener('wheel', handleWheel));
  }

  private createToolbar(main: HTMLElement): void {
    this.editButtons = [];
    const toolbar = main.createDiv({ cls: 'native-powerpoint-toolbar' });

    const navGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.createIconButton(navGroup, 'chevron-left', 'Previous slide', () => void this.goToSlide(this.currentSlide - 1));
    this.slideCounterEl = navGroup.createDiv({ cls: 'native-powerpoint-page-counter', text: '0 / 0' });
    this.createIconButton(navGroup, 'chevron-right', 'Next slide', () => void this.goToSlide(this.currentSlide + 1));

    const historyGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.undoButton = this.createIconButton(historyGroup, 'undo-2', 'Undo', () => void this.undo());
    this.redoButton = this.createIconButton(historyGroup, 'redo-2', 'Redo', () => void this.redo());

    const slideGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.createEditIconButton(slideGroup, 'plus', 'Add slide', () => void this.addSlide());
    this.createEditIconButton(slideGroup, 'trash-2', 'Delete slide', () => void this.deleteSlide());
    this.createEditIconButton(slideGroup, 'arrow-left-to-line', 'Move slide left', () => void this.moveSlide(-1));
    this.createEditIconButton(slideGroup, 'arrow-right-to-line', 'Move slide right', () => void this.moveSlide(1));

    const objectGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.createEditIconButton(objectGroup, 'type', 'Add text box', () => void this.addTextBox());
    this.copyButton = this.createIconButton(objectGroup, 'copy', 'Copy selected object (Ctrl+C)', () => void this.copySelectedShape());
    this.pasteButton = this.createIconButton(objectGroup, 'clipboard-paste', 'Paste object (Ctrl+V)', () => void this.pasteCopiedShape());
    this.duplicateButton = this.createIconButton(objectGroup, 'copy-plus', 'Duplicate selected object (Ctrl+D)', () => void this.duplicateSelectedShape());
    this.createEditIconButton(objectGroup, 'eraser', 'Delete selected object', () => void this.deleteSelectedShape());

    const zoomGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.createIconButton(zoomGroup, 'zoom-out', 'Zoom out', () => this.setZoom(this.zoomLevel - 0.1));
    this.createIconButton(zoomGroup, 'zoom-in', 'Zoom in', () => this.setZoom(this.zoomLevel + 0.1));

    const searchGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.createIconButton(searchGroup, 'search', 'Find in presentation', () => this.openFindPanel());
    this.createFindPanel(toolbar);

    this.statusEl = toolbar.createDiv({ cls: 'native-powerpoint-save-status', text: 'Ready' });
    this.updateEditingAvailability();
    this.updateHistoryAvailability();
    this.updateObjectClipboardAvailability();
  }

  private createFindPanel(toolbar: HTMLElement): void {
    const panel = toolbar.createDiv({ cls: 'native-powerpoint-find-panel' });
    this.findPanelEl = panel;

    const input = panel.createEl('input', {
      cls: 'native-powerpoint-find-input',
      type: 'search',
      attr: {
        'aria-label': 'Find text in presentation',
        placeholder: 'Find text'
      }
    });
    this.findInputEl = input;

    this.findStatusEl = panel.createDiv({ cls: 'native-powerpoint-find-status', text: 'No search' });

    const previousButton = panel.createEl('button', {
      cls: 'native-powerpoint-find-btn',
      attr: { 'aria-label': 'Previous match', title: 'Previous match' }
    });
    setIcon(previousButton, 'chevron-up');

    const nextButton = panel.createEl('button', {
      cls: 'native-powerpoint-find-btn',
      attr: { 'aria-label': 'Next match', title: 'Next match' }
    });
    setIcon(nextButton, 'chevron-down');

    const closeButton = panel.createEl('button', {
      cls: 'native-powerpoint-find-btn',
      attr: { 'aria-label': 'Close find', title: 'Close find' }
    });
    setIcon(closeButton, 'x');

    input.addEventListener('input', () => {
      void this.refreshFindMatches({ reveal: true });
    });
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        this.closeFindPanel();
      } else if (event.key === 'Enter') {
        event.preventDefault();
        void this.moveFindMatch(event.shiftKey ? -1 : 1);
      }
    });
    previousButton.addEventListener('click', () => void this.moveFindMatch(-1));
    nextButton.addEventListener('click', () => void this.moveFindMatch(1));
    closeButton.addEventListener('click', () => this.closeFindPanel());
  }

  private createIconButton(container: HTMLElement, icon: string, label: string, onClick: () => void): HTMLButtonElement {
    const button = container.createEl('button', {
      cls: 'native-powerpoint-toolbar-btn',
      attr: { 'aria-label': label, title: label }
    });
    setIcon(button, icon);
    button.addEventListener('click', onClick);
    return button;
  }

  private createEditIconButton(container: HTMLElement, icon: string, label: string, onClick: () => void): HTMLButtonElement {
    const button = this.createIconButton(container, icon, label, () => {
      if (this.ensureEditable(label.toLowerCase())) {
        onClick();
      }
    });
    button.dataset.baseTitle = label;
    this.editButtons.push(button);
    return button;
  }

  private async captureHistoryEntry(label: string): Promise<HistoryEntry> {
    if (!this.engine) {
      throw new Error('Open a loaded PowerPoint file first.');
    }

    return {
      buffer: await this.engine.export(),
      currentSlide: this.currentSlide,
      label
    };
  }

  private recordHistoryEntry(entry: HistoryEntry): void {
    this.undoStack.push(entry);
    if (this.undoStack.length > HISTORY_LIMIT) {
      this.undoStack.shift();
    }
    this.redoStack = [];
    this.updateHistoryAvailability();
  }

  private clearHistory(): void {
    this.undoStack = [];
    this.redoStack = [];
    this.isRestoringHistory = false;
    this.updateHistoryAvailability();
  }

  private async undo(): Promise<void> {
    await this.restoreHistoryEntry(this.undoStack, this.redoStack, 'undo');
  }

  private async redo(): Promise<void> {
    await this.restoreHistoryEntry(this.redoStack, this.undoStack, 'redo');
  }

  private async restoreHistoryEntry(
    source: HistoryEntry[],
    destination: HistoryEntry[],
    action: 'undo' | 'redo'
  ): Promise<void> {
    if (!this.engine || this.isRestoringHistory || source.length === 0) return;
    if (!this.ensureEditable(action)) return;
    if (this.activeEditor) {
      this.activeEditor.blur();
      return;
    }

    const entry = source[source.length - 1];
    if (!entry) return;

    this.isRestoringHistory = true;
    this.clearAutosave();
    this.dragState = null;
    this.updateHistoryAvailability();

    try {
      const current = await this.captureHistoryEntry(entry.label);
      await this.engine.restoreSnapshot(entry.buffer);
      source.pop();
      destination.push(current);
      if (destination.length > HISTORY_LIMIT) {
        destination.shift();
      }

      this.currentSlide = Math.max(0, Math.min(entry.currentSlide, this.engine.slideCount - 1));
      this.clearSelection();
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not ${action}: ${cleanError(error)}`);
    } finally {
      this.isRestoringHistory = false;
      this.updateHistoryAvailability();
    }
  }

  private openFindPanel(): void {
    if (!this.engine || this.isLoading) {
      new Notice('Open a loaded PowerPoint file to search it.');
      return;
    }

    this.findPanelEl?.addClass('is-open');
    const seedText = this.getSelectedFindSeedText();
    if (seedText && this.findInputEl && !this.findInputEl.value.trim()) {
      this.findInputEl.value = seedText;
    }

    void this.refreshFindMatches({ reveal: Boolean(this.findInputEl?.value.trim()) });
    window.requestAnimationFrame(() => {
      this.findPanelEl?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      this.findInputEl?.focus();
      this.findInputEl?.select();
    });
  }

  private closeFindPanel(): void {
    this.findPanelEl?.removeClass('is-open');
    this.clearFindHighlight();
  }

  private getSelectedFindSeedText(): string {
    if (
      this.activeEditor
      && this.activeEditor.selectionStart !== null
      && this.activeEditor.selectionEnd !== null
      && this.activeEditor.selectionStart !== this.activeEditor.selectionEnd
    ) {
      return normalizeSearchText(this.activeEditor.value.slice(this.activeEditor.selectionStart, this.activeEditor.selectionEnd));
    }

    return '';
  }

  private async refreshFindMatches(options: { reveal?: boolean } = {}): Promise<void> {
    const query = this.findInputEl?.value.trim() ?? '';
    this.clearFindHighlight();

    if (!this.engine || !query) {
      this.findMatches = [];
      this.currentFindMatchIndex = 0;
      this.updateFindStatus();
      return;
    }

    this.findMatches = this.collectFindMatches(query);
    if (this.findMatches.length === 0) {
      this.currentFindMatchIndex = 0;
      this.updateFindStatus();
      return;
    }

    const currentSlideMatchIndex = this.findMatches.findIndex((match) => match.slideIndex >= this.currentSlide);
    this.currentFindMatchIndex = currentSlideMatchIndex === -1 ? 0 : currentSlideMatchIndex;

    if (options.reveal) {
      await this.revealCurrentFindMatch();
    } else {
      this.applyFindHighlight();
      this.updateFindStatus();
    }
  }

  private collectFindMatches(query: string): PowerPointFindMatch[] {
    const engine = this.engine;
    if (!engine) return [];

    const queryLower = query.toLocaleLowerCase();
    const matches: PowerPointFindMatch[] = [];
    const parser = new DOMParser();

    for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
      const { svg } = engine.renderSlide(slideIndex);
      const slideDocument = parser.parseFromString(svg, 'image/svg+xml');
      const shapeElements = Array.from(slideDocument.querySelectorAll('g[data-ooxml-shape-idx]'));
      let foundShapeMatch = false;

      for (const shape of shapeElements) {
        const shapeIndex = getShapeIndex(shape);
        if (shapeIndex === null) continue;

        const text = normalizeSearchText(shape.textContent ?? '');
        if (!text || !text.toLocaleLowerCase().includes(queryLower)) continue;

        foundShapeMatch = true;
        matches.push({ slideIndex, shapeIndex, text });
      }

      if (!foundShapeMatch) {
        const slideText = normalizeSearchText(slideDocument.documentElement.textContent ?? '');
        if (slideText && slideText.toLocaleLowerCase().includes(queryLower)) {
          matches.push({ slideIndex, shapeIndex: null, text: slideText });
        }
      }
    }

    return matches;
  }

  private async moveFindMatch(direction: -1 | 1): Promise<void> {
    if (!this.findInputEl?.value.trim()) {
      this.openFindPanel();
      return;
    }

    if (this.findMatches.length === 0) {
      await this.refreshFindMatches();
      if (this.findMatches.length === 0) return;
    }

    this.currentFindMatchIndex = (this.currentFindMatchIndex + direction + this.findMatches.length) % this.findMatches.length;
    await this.revealCurrentFindMatch();
  }

  private async revealCurrentFindMatch(): Promise<void> {
    const match = this.findMatches[this.currentFindMatchIndex];
    if (!match) {
      this.updateFindStatus();
      return;
    }

    if (match.slideIndex !== this.currentSlide) {
      this.currentSlide = match.slideIndex;
      this.selectedShapeIndex = null;
      this.selectedTransform = null;
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        await this.renderThumbnails();
      }
    }

    if (match.shapeIndex !== null) {
      this.selectShape(match.shapeIndex);
    } else {
      this.clearSelection();
    }

    this.applyFindHighlight();
    this.updateFindStatus();
  }

  private clearFindHighlight(): void {
    this.svgEl?.querySelectorAll('.native-powerpoint-find-current').forEach((element) => {
      element.removeClass('native-powerpoint-find-current');
    });
  }

  private applyFindHighlight(): void {
    this.clearFindHighlight();
    const match = this.findMatches[this.currentFindMatchIndex];
    if (!match || match.slideIndex !== this.currentSlide || match.shapeIndex === null || !this.svgEl) {
      return;
    }

    const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${match.shapeIndex}"]`);
    shape?.addClass('native-powerpoint-find-current');
  }

  private updateFindStatus(): void {
    if (!this.findStatusEl) return;

    const query = this.findInputEl?.value.trim() ?? '';
    if (!query) {
      this.findStatusEl.setText('No search');
      this.findStatusEl.removeAttribute('title');
      return;
    }

    if (this.findMatches.length === 0) {
      this.findStatusEl.setText('No matches');
      this.findStatusEl.removeAttribute('title');
      return;
    }

    const match = this.findMatches[this.currentFindMatchIndex];
    const slideLabel = match ? `Slide ${match.slideIndex + 1}` : 'Slide';
    this.findStatusEl.setText(`${this.currentFindMatchIndex + 1} / ${this.findMatches.length} | ${slideLabel}`);
    if (match) {
      this.findStatusEl.setAttribute('title', match.text);
    }
  }

  private registerKeyboardHandlers(): void {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (!this.containerEl.isShown()) return;

      if (isPrimaryFindShortcut(event) && this.isActivePowerPointView()) {
        const target = event.target instanceof Element ? event.target : null;
        if (!target?.closest('.modal')) {
          event.preventDefault();
          event.stopImmediatePropagation();
          this.openFindPanel();
          return;
        }
      }

      if (!this.isActivePowerPointView()) return;
      if (this.activeEditor && document.activeElement === this.activeEditor) return;

      const target = event.target instanceof Element ? event.target : null;
      if (target?.closest('input, textarea, select, [contenteditable="true"]')) {
        if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 's') {
          event.preventDefault();
          event.stopImmediatePropagation();
          void this.saveCurrentPresentation();
        }
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        event.stopImmediatePropagation();
        void this.saveCurrentPresentation();
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'z') {
        event.preventDefault();
        event.stopImmediatePropagation();
        void (event.shiftKey ? this.redo() : this.undo());
        return;
      }

      if (event.ctrlKey && !event.metaKey && event.key.toLowerCase() === 'y') {
        event.preventDefault();
        event.stopImmediatePropagation();
        void this.redo();
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'c') {
        if (this.selectedShapeIndex !== null) {
          event.preventDefault();
          event.stopImmediatePropagation();
          void this.copySelectedShape();
        }
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'v') {
        if (this.objectClipboard) {
          event.preventDefault();
          event.stopImmediatePropagation();
          void this.pasteCopiedShape();
        }
        return;
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'd') {
        if (this.selectedShapeIndex !== null) {
          event.preventDefault();
          event.stopImmediatePropagation();
          void this.duplicateSelectedShape();
        }
        return;
      }

      if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
        event.preventDefault();
        void this.goToSlide(this.currentSlide - 1);
      } else if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
        event.preventDefault();
        void this.goToSlide(this.currentSlide + 1);
      } else if (event.key === 'Delete' || event.key === 'Backspace') {
        if (this.selectedShapeIndex !== null) {
          event.preventDefault();
          void this.deleteSelectedShape();
        }
      }
    };

    this.registerDomEvent(window, 'keydown', handleKeyDown, true);
    this.registerDomEvent(document, 'keydown', handleKeyDown, true);

    this.registerDomEvent(window, 'resize', () => this.updateSlideScale());
    this.registerDomEvent(document, 'pointermove', this.handleDragMove, true);
    this.registerDomEvent(document, 'pointerup', this.handleDragEnd, true);
    this.registerDomEvent(document, 'pointerdown', this.handleOutsideSlidePointerDown, true);
  }

  private isActivePowerPointView(): boolean {
    if (this.app.workspace.getActiveViewOfType(NativePowerPointView) === this) {
      return true;
    }

    const workspace = this.app.workspace as typeof this.app.workspace & { activeLeaf?: WorkspaceLeaf | null };
    if (workspace.activeLeaf === this.leaf) {
      return true;
    }

    if (this.contentEl.closest('.workspace-leaf.mod-active')) {
      return true;
    }

    const activeElement = document.activeElement;
    return Boolean(activeElement instanceof Node && this.contentEl.contains(activeElement));
  }

  private async loadPresentation(file: TFile): Promise<void> {
    this.clearAutosave();
    this.removeActiveEditor();
    this.clearHistory();
    this.dragState = null;
    this.isLoading = true;
    this.engine = null;
    this.loadedFile = file;
    this.sourcePackage = null;
    this.sourceBuffer = null;
    this.currentSlide = 0;
    this.selectedShapeIndex = null;
    this.selectedTransform = null;
    this.svgSecurityDecision = null;
    this.isViewOnly = false;
    this.viewOnlyReason = '';
    this.isDirty = false;
    this.editVersion = 0;
    this.findMatches = [];
    this.currentFindMatchIndex = 0;
    this.hasShownGeneratedTextNotice = false;
    this.fontSubstitutions = [];
    if (this.findInputEl) this.findInputEl.value = '';
    this.closeFindPanel();
    this.updateFindStatus();
    this.setSaveState('idle');
    this.updateEditingAvailability();
    this.renderInspector();
    this.showLoading(`Loading ${file.name}...`);

    if (!isModernPowerPointExtension(file.extension)) {
      this.isLoading = false;
      this.showUnsupported(file);
      return;
    }

    try {
      const buffer = await this.app.vault.readBinary(file);
      const sourcePackage = inspectPowerPointPackage(buffer);
      const sourceValidation = validatePowerPointPackageStructure(sourcePackage);
      if (!sourceValidation.ok) {
        throw new Error(summarizePackageMessages(sourceValidation.errors));
      }

      this.sourcePackage = sourcePackage;
      this.sourceBuffer = buffer;
      this.isViewOnly = this.shouldOpenViewOnly(file, sourcePackage);
      this.viewOnlyReason = this.getViewOnlyReason(file, sourcePackage);
      this.engine = await PresentationEngine.load(buffer);
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        await this.renderThumbnails();
      }
      this.setSaveState(this.isViewOnly ? 'view-only' : 'saved');
      if (this.isViewOnly) {
        new Notice(`Opened ${file.name} view-only: ${this.viewOnlyReason}`);
      }
    } catch (error) {
      this.showError(`Could not open ${file.name}: ${cleanError(error)}`);
    } finally {
      this.isLoading = false;
      this.updateSlideCounter();
      this.updateEditingAvailability();
      this.renderInspector();
    }
  }

  private showLoading(message: string): void {
    if (!this.slideSurface) return;
    this.resetSlideSurfaceSizing();
    this.slideSurface.empty();
    this.slideSurface.createDiv({ cls: 'native-powerpoint-loading', text: message });
    this.thumbnailContainer?.empty();
  }

  private showUnsupported(file: TFile): void {
    if (!this.slideSurface) return;
    this.resetSlideSurfaceSizing();
    this.slideSurface.empty();
    this.slideSurface.createDiv({
      cls: 'native-powerpoint-error',
      text: `${file.extension.toUpperCase()} is a legacy binary PowerPoint format. Native editing supports modern Open XML files first: .pptx, .pptm, .ppsx, .ppsm, .potx, and .potm.`
    });
    this.thumbnailContainer?.empty();
    this.setSaveState('idle');
    this.updateEditingAvailability();
  }

  private showError(message: string): void {
    if (!this.slideSurface) return;
    this.resetSlideSurfaceSizing();
    this.slideSurface.empty();
    this.slideSurface.createDiv({ cls: 'native-powerpoint-error', text: message });
    this.thumbnailContainer?.empty();
    this.setSaveState('failed');
    this.updateEditingAvailability();
  }

  private async renderCurrentSlide(keepSelection = false): Promise<boolean> {
    if (!this.engine || !this.slideSurface) return false;

    const selectedShape = keepSelection ? this.selectedShapeIndex : null;
    const { svg } = this.engine.renderSlide(this.currentSlide);
    const safeSvg = this.prepareSvgForRender(svg);

    if (!safeSvg.allowed) {
      this.showUnsafeSvgWarning(safeSvg.issues);
      return false;
    }

    const svgElement = createSvgElementFromString(safeSvg.svg, this.slideSurface.ownerDocument);
    if (!svgElement) {
      this.showError('Could not render this PowerPoint slide because its SVG preview could not be read.');
      return false;
    }

    this.slideSurface.empty();
    this.slideSurface.appendChild(svgElement);
    this.svgEl = svgElement;

    if (this.svgEl) {
      this.fontSubstitutions = this.engine.applyFontFidelity(this.svgEl);
      this.engine.formatChartAxisLabels(this.svgEl, this.currentSlide);
      normalizeSvgForDisplay(this.svgEl);
      this.markGeneratedTextEditability(this.svgEl);
      this.svgEl.addClass('native-powerpoint-slide-svg');
      this.slideSurface.addClass('is-rendered');
      this.updateSlideScale();
      window.requestAnimationFrame(() => this.updateSlideScale());
      this.attachSvgEvents();
    }

    if (selectedShape !== null) {
      this.selectShape(selectedShape);
    } else {
      this.clearSelection();
    }

    this.applyFindHighlight();
    this.updateSlideCounter();
    return true;
  }

  private prepareSvgForRender(svg: string, isThumbnail = false): { svg: string; issues: SvgSecurityIssue[]; allowed: boolean } {
    const settings = this.getSettings();
    if (settings.openWithYoloMode || this.svgSecurityDecision === 'yolo') {
      return {
        svg,
        issues: [],
        allowed: true
      };
    }

    const sanitizerMode =
      this.svgSecurityDecision === 'compatibility' || !settings.hideUnsupportedSvgContent
        ? 'compatibility'
        : 'strict';
    const scannedSvg = sanitizeSvg(svg, { mode: sanitizerMode });
    const shouldWarn =
      this.svgSecurityDecision === null &&
      scannedSvg.issues.length > 0 &&
      !isThumbnail;
    return {
      svg: scannedSvg.svg,
      issues: scannedSvg.issues,
      allowed: !shouldWarn
    };
  }

  private showUnsafeSvgWarning(issues: SvgSecurityIssue[]): void {
    if (!this.slideSurface) return;

    this.svgEl = null;
    this.clearSelection();
    this.resetSlideSurfaceSizing();
    this.slideSurface.empty();
    this.thumbnailContainer?.empty();

    const warning = this.slideSurface.createDiv({ cls: 'native-powerpoint-security-warning' });
    warning.createDiv({ cls: 'native-powerpoint-security-title', text: 'Advanced slide content detected' });
    warning.createEl('p', {
      text: 'This slide has SVG content that would be hidden in the preview, which can make graphics disappear. Compatibility mode hides those parts only in Obsidian and does not delete them from the PPTX. YOLO mode opens the original slide SVG, so only use it for decks you trust.'
    });

    const summary = summarizeSvgSecurityIssues(issues);
    const list = warning.createEl('ul', { cls: 'native-powerpoint-security-list' });
    for (const item of summary.slice(0, 6)) {
      list.createEl('li', { text: item });
    }

    if (summary.length > 6) {
      list.createEl('li', { text: `${summary.length - 6} more issue types hidden` });
    }

    const actions = warning.createDiv({ cls: 'native-powerpoint-security-actions' });
    const openCompatibility = actions.createEl('button', { text: 'Open in compatibility mode' });
    openCompatibility.addClass('mod-warning');
    openCompatibility.addEventListener('click', () => {
      this.svgSecurityDecision = 'compatibility';
      new Notice('Opening this PowerPoint in preview-safe compatibility mode for this session.');
      void this.renderCurrentSlide().then((rendered) => {
        if (rendered) void this.renderThumbnails();
      });
    });

    const openYolo = actions.createEl('button', { text: 'Open with YOLO mode' });
    openYolo.addClass('mod-warning');
    openYolo.addEventListener('click', () => {
      this.svgSecurityDecision = 'yolo';
      new Notice('Opening this PowerPoint with YOLO mode for this session.');
      void this.renderCurrentSlide().then((rendered) => {
        if (rendered) void this.renderThumbnails();
      });
    });

    const rememberYolo = actions.createEl('button', { text: 'Always use YOLO mode' });
    rememberYolo.addClass('mod-warning');
    rememberYolo.addEventListener('click', () => {
      this.svgSecurityDecision = 'yolo';
      void this.getSettings().setOpenWithYoloMode(true)
        .then(() => {
          new Notice('YOLO mode will be used for future PowerPoint files.');
        })
        .catch((error) => {
          new Notice(`Could not remember YOLO mode: ${cleanError(error)}`);
        })
        .finally(() => {
          void this.renderCurrentSlide().then((rendered) => {
            if (rendered) void this.renderThumbnails();
          });
        });
    });
  }

  private resetSlideSurfaceSizing(): void {
    if (!this.slideSurface) return;

    this.slideSurface.removeClass('is-rendered');
    this.slideSurface.removeClass('is-scaled');
    this.slideSurface.style.removeProperty('--native-powerpoint-slide-width');
    this.slideSurface.style.removeProperty('--native-powerpoint-slide-height');

    if (this.svgEl) {
      this.svgEl.style.removeProperty('width');
      this.svgEl.style.removeProperty('height');
      this.svgEl.style.removeProperty('transform');
      this.svgEl.style.removeProperty('transform-origin');
    }
  }

  private attachSvgEvents(): void {
    if (!this.svgEl) return;

    this.svgEl.addEventListener('click', (event) => {
      const target = event.target instanceof Element ? event.target : null;
      if (this.suppressNextTextClick && target?.closest('text')) {
        event.preventDefault();
        event.stopPropagation();
        this.suppressNextTextClick = false;
        return;
      }
      this.suppressNextTextClick = false;

      const shape = target?.closest('g[data-ooxml-shape-idx]') ?? null;
      const shapeIndex = getShapeIndex(shape);
      if (shapeIndex === null) {
        this.clearSelection();
        return;
      }

      this.selectShape(shapeIndex);
      if (target?.closest('text') && target.closest(GENERATED_GRID_SELECTOR)) {
        const textTarget = this.getGeneratedTextEditTarget(target);
        if (textTarget && this.ensureEditable('edit text')) {
          event.preventDefault();
          event.stopPropagation();
          this.startTextEditor(textTarget, event.clientX, event.clientY);
        } else {
          this.showGeneratedTextNotice();
        }
        return;
      }

      const textTarget = this.getTextEditTarget(target);
      if (textTarget && this.ensureEditable('edit text')) {
        event.preventDefault();
        event.stopPropagation();
        this.startTextEditor(textTarget, event.clientX, event.clientY);
      }
    });

    this.svgEl.addEventListener('dblclick', (event) => {
      const target = event.target instanceof Element ? event.target : null;
      const shape = target?.closest('g[data-ooxml-shape-idx]') ?? null;
      const shapeIndex = getShapeIndex(shape);
      if (shapeIndex !== null) {
        event.preventDefault();
        this.selectShape(shapeIndex);
        if (target?.closest('text') && target.closest(GENERATED_GRID_SELECTOR)) {
          const textTarget = this.getGeneratedTextEditTarget(target);
          if (textTarget && this.ensureEditable('edit text')) {
            this.startTextEditor(textTarget, event.clientX, event.clientY);
          } else {
            this.showGeneratedTextNotice();
          }
          return;
        }

        const textTarget = this.getTextEditTarget(target);
        if (this.ensureEditable('edit text')) {
          this.startTextEditor(textTarget, event.clientX, event.clientY);
        }
      }
    });

    this.svgEl.addEventListener('pointerdown', (event) => {
      if (event.button !== 0) return;

      const target = event.target instanceof Element ? event.target : null;
      if (target?.closest('text')) {
        this.handleInlineTextPointerDown(event, target);
        return;
      }

      if (this.selectedShapeIndex === null) return;

      const shape = target?.closest('g[data-ooxml-shape-idx]') ?? null;
      const shapeIndex = getShapeIndex(shape);
      if (shapeIndex !== this.selectedShapeIndex || !this.selectedTransform) return;

      event.preventDefault();
      if (this.ensureEditable('move object')) {
        this.startDrag(event, 'move');
      }
    });
  }

  private async renderThumbnails(): Promise<void> {
    if (!this.engine || !this.thumbnailContainer) return;

    this.thumbnailContainer.empty();

    for (let index = 0; index < this.engine.slideCount; index++) {
      const item = this.thumbnailContainer.createDiv({ cls: 'native-powerpoint-thumbnail' });
      if (index === this.currentSlide) item.addClass('active');

      const preview = item.createDiv({ cls: 'native-powerpoint-thumbnail-preview' });
      try {
        const safeSvg = this.prepareSvgForRender(this.engine.renderSlide(index).svg, true);
        const thumbnailSvg = createSvgElementFromString(safeSvg.svg, preview.ownerDocument);
        if (!thumbnailSvg) {
          throw new Error('Could not read thumbnail SVG.');
        }
        preview.appendChild(thumbnailSvg);
      } catch {
        preview.createDiv({ cls: 'native-powerpoint-thumbnail-error', text: '!' });
      }
      const thumbnailSvg = preview.querySelector('svg');
      if (thumbnailSvg) {
        this.engine.applyFontFidelity(thumbnailSvg);
        this.engine.formatChartAxisLabels(thumbnailSvg, index);
        normalizeSvgForDisplay(thumbnailSvg);
        thumbnailSvg.addClass('native-powerpoint-thumbnail-svg');
      }

      item.createDiv({ cls: 'native-powerpoint-thumbnail-number', text: `${index + 1}` });
      item.addEventListener('click', () => void this.goToSlide(index));
    }
  }

  private async goToSlide(index: number): Promise<void> {
    if (!this.engine || this.isLoading) return;
    if (index < 0 || index >= this.engine.slideCount || index === this.currentSlide) return;

    this.currentSlide = index;
    this.selectedShapeIndex = null;
    this.selectedTransform = null;
    const rendered = await this.renderCurrentSlide();
    if (rendered) {
      await this.renderThumbnails();
      this.renderInspector();
    }
  }

  private async addSlide(): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('add slide')) return;

    try {
      const history = await this.captureHistoryEntry('Add slide');
      const result = await this.engine.addSlide(this.currentSlide);
      this.currentSlide = result.slideIndex;
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not add slide: ${cleanError(error)}`);
    }
  }

  private async deleteSlide(): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('delete slide')) return;

    try {
      const history = await this.captureHistoryEntry('Delete slide');
      const result = await this.engine.deleteSlide(this.currentSlide);
      this.currentSlide = result.slideIndex;
      this.clearSelection();
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not delete slide: ${cleanError(error)}`);
    }
  }

  private async moveSlide(direction: -1 | 1): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('move slide')) return;

    try {
      const history = await this.captureHistoryEntry('Move slide');
      const previousSlide = this.currentSlide;
      const result = await this.engine.moveSlide(this.currentSlide, direction);
      if (result.slideIndex === previousSlide) return;

      this.currentSlide = result.slideIndex;
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not move slide: ${cleanError(error)}`);
    }
  }

  private async addTextBox(): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('add text box')) return;

    try {
      const history = await this.captureHistoryEntry('Add text box');
      const shapeIndex = this.engine.addTextBox(this.currentSlide);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) {
        this.selectShape(shapeIndex);
        this.startTextEditor();
      }
    } catch (error) {
      new Notice(`Could not add text box: ${cleanError(error)}`);
    }
  }

  private async deleteSelectedShape(): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('delete object')) return;

    try {
      const history = await this.captureHistoryEntry('Delete object');
      this.engine.deleteShape(this.currentSlide, this.selectedShapeIndex);
      this.clearSelection();
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not delete object: ${cleanError(error)}`);
    }
  }

  private async copySelectedShape(): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) {
      new Notice('Select a slide object to copy.');
      return;
    }

    try {
      this.objectClipboard = await this.engine.copyShape(this.currentSlide, this.selectedShapeIndex);
      this.updateObjectClipboardAvailability();
      new Notice('Copied slide object.');
    } catch (error) {
      new Notice(`Could not copy object: ${cleanError(error)}`);
    }
  }

  private async pasteCopiedShape(): Promise<void> {
    if (!this.engine || !this.objectClipboard) {
      new Notice('Copy a slide object first.');
      return;
    }
    if (!this.ensureEditable('paste object')) return;

    try {
      const history = await this.captureHistoryEntry('Paste object');
      const shapeIndex = await this.engine.pasteShape(this.objectClipboard, this.currentSlide);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not paste object: ${cleanError(error)}`);
    }
  }

  private async duplicateSelectedShape(): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) {
      new Notice('Select a slide object to duplicate.');
      return;
    }
    if (!this.ensureEditable('duplicate object')) return;

    try {
      const history = await this.captureHistoryEntry('Duplicate object');
      const shapeIndex = await this.engine.duplicateShape(this.currentSlide, this.selectedShapeIndex);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not duplicate object: ${cleanError(error)}`);
    }
  }

  private selectShape(shapeIndex: number): void {
    if (!this.engine || !this.svgEl) return;

    this.selectedShapeIndex = shapeIndex;
    this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      shape.removeClass('native-powerpoint-shape-selected');
    });

    const selected = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!(selected instanceof SVGGElement)) return;

    selected.addClass('native-powerpoint-shape-selected');
    this.selectedTransform = cloneTransform(this.engine.getShapeTransform(selected));
    this.renderInspector();
    this.updateSelectionOverlay();
    this.updateObjectClipboardAvailability();
  }

  private clearSelection(): void {
    this.selectedShapeIndex = null;
    this.selectedTransform = null;
    this.svgEl?.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      shape.removeClass('native-powerpoint-shape-selected');
    });
    this.removeSelectionOverlay();
    this.renderInspector();
    this.updateObjectClipboardAvailability();
  }

  private renderInspector(): void {
    if (!this.inspectorEl) return;

    this.inspectorEl.empty();
    this.inspectorEl.createDiv({ cls: 'native-powerpoint-inspector-title', text: 'Inspector' });
    this.renderViewOnlyWarning(this.inspectorEl);
    this.renderFontFidelity(this.inspectorEl);

    if (!this.engine || this.selectedShapeIndex === null || !this.selectedTransform) {
      this.inspectorEl.createDiv({
        cls: 'native-powerpoint-inspector-empty',
        text: 'Select a slide object to adjust its layout. Click text on the slide to edit it directly.'
      });
      this.xInput = null;
      this.yInput = null;
      this.widthInput = null;
      this.heightInput = null;
      this.rotationInput = null;
      return;
    }

    const selected = this.getSelectedShapeElement();
    this.inspectorEl.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: `Object ${this.selectedShapeIndex + 1}` });
    this.inspectorEl.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: selected?.closest(GENERATED_GRID_SELECTOR)
        ? 'Click highlighted table or chart text on the slide to edit it directly. Generated numeric chart ticks remain read-only.'
        : 'Click text on the slide to edit it directly.'
    });

    const grid = this.inspectorEl.createDiv({ cls: 'native-powerpoint-inspector-grid' });
    this.xInput = this.createNumberField(grid, 'X', this.engine.emuToPx(this.selectedTransform.x));
    this.yInput = this.createNumberField(grid, 'Y', this.engine.emuToPx(this.selectedTransform.y));
    this.widthInput = this.createNumberField(grid, 'W', this.engine.emuToPx(this.selectedTransform.cx));
    this.heightInput = this.createNumberField(grid, 'H', this.engine.emuToPx(this.selectedTransform.cy));
    this.rotationInput = this.createNumberField(grid, 'Rot', this.engine.ooxmlToDegrees(this.selectedTransform.rot));
    this.xInput.disabled = !this.canEdit();
    this.yInput.disabled = !this.canEdit();
    this.widthInput.disabled = !this.canEdit();
    this.heightInput.disabled = !this.canEdit();
    this.rotationInput.disabled = !this.canEdit();

    const applyLayout = this.inspectorEl.createEl('button', { cls: 'native-powerpoint-inspector-button', text: 'Apply layout' });
    applyLayout.disabled = !this.canEdit();
    applyLayout.addEventListener('click', () => void this.applyInspectorTransform());

    if (selected?.getAttribute('data-ooxml-shape-type') === 'chart') {
      const chartData = this.engine.getChartDataGrid(this.currentSlide, this.selectedShapeIndex);
      if (chartData) {
        this.renderChartDataEditor(chartData);
      }
    }
  }

  private renderChartDataEditor(chartData: ChartDataGrid): void {
    if (!this.inspectorEl) return;

    const section = this.inspectorEl.createDiv({ cls: 'native-powerpoint-chart-data' });
    section.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: 'Chart data' });

    if (!chartData.editable) {
      section.createDiv({
        cls: 'native-powerpoint-inspector-hint',
        text: chartData.reason || 'This chart data grid is read-only.'
      });
      return;
    }

    section.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: 'Edit source-backed cells below. Apply updates the chart cache and its embedded Excel workbook.'
    });

    const viewport = section.createDiv({ cls: 'native-powerpoint-chart-data-scroll' });
    const table = viewport.createEl('table', { cls: 'native-powerpoint-chart-data-grid' });
    const header = table.createEl('thead').createEl('tr');
    header.createEl('th', { text: chartData.categoryLabel });
    chartData.series.forEach((series) => {
      header.createEl('th', { text: series.name });
      if (series.pointLabels !== null) {
        header.createEl('th', { text: `${series.name} label` });
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
      text: 'Apply chart data'
    });
    apply.disabled = !this.canEdit();
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
    input.disabled = !this.canEdit();
    return input;
  }

  private async applyChartData(update: ChartDataUpdate): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('edit chart data')) return;

    try {
      const history = await this.captureHistoryEntry('Edit chart data');
      await this.engine.updateChartData(this.currentSlide, this.selectedShapeIndex, update);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not update chart data: ${cleanError(error)}`);
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
    if (!this.isViewOnly || !this.viewOnlyReason) return;

    container.createDiv({
      cls: 'native-powerpoint-view-only-warning',
      text: this.viewOnlyReason
    });
  }

  private async applyTextValue(text: string, target: TextEditTarget | null = null): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('edit text')) return;

    try {
      const previousText = target?.text ?? this.getSelectedShapeElement()?.textContent?.trim() ?? '';
      if (text === previousText) return;

      const history = await this.captureHistoryEntry('Edit text');
      const scrollPosition = this.captureCanvasScroll();
      if (target?.kind === 'shape-run') {
        await this.engine.updateTextRun(
          this.currentSlide,
          target.shapeIndex,
          target.paragraphIndex,
          target.runIndex,
          text
        );
      } else if (target) {
        await this.engine.updateGeneratedText(this.currentSlide, target.shapeIndex, target, text);
      } else {
        await this.engine.updateShapeText(this.currentSlide, this.selectedShapeIndex, text);
      }
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) {
        this.restoreCanvasScrollSoon(scrollPosition);
      }
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not update text: ${cleanError(error)}`);
    }
  }

  private async applyInspectorTransform(): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null || !this.selectedTransform) return;
    if (!this.ensureEditable('edit layout')) return;

    const transform = cloneTransform(this.selectedTransform);
    transform.x = this.engine.pxToEmu(Number(this.xInput?.value || 0));
    transform.y = this.engine.pxToEmu(Number(this.yInput?.value || 0));
    transform.cx = this.engine.pxToEmu(Math.max(1, Number(this.widthInput?.value || 1)));
    transform.cy = this.engine.pxToEmu(Math.max(1, Number(this.heightInput?.value || 1)));
    transform.rot = this.engine.degreesToOoxml(Number(this.rotationInput?.value || 0));
    await this.commitTransform(transform);
  }

  private handleInlineTextPointerDown(event: PointerEvent, target: Element): void {
    const textTarget = target.closest('text') && target.closest(GENERATED_GRID_SELECTOR)
      ? this.getGeneratedTextEditTarget(target)
      : this.getTextEditTarget(target);

    event.preventDefault();
    event.stopPropagation();
    this.clearBrowserTextSelection();

    if (!textTarget) {
      if (target.closest(GENERATED_GRID_SELECTOR)) {
        this.showGeneratedTextNotice();
      }
      return;
    }
    if (!this.ensureEditable('edit text')) return;

    this.suppressNextTextClick = true;
    this.selectShape(textTarget.shapeIndex);
    this.startTextEditor(textTarget, event.clientX, event.clientY);

    const editor = this.activeEditor;
    if (!editor || this.activeEditorTarget !== textTarget.element) return;

    const box = this.getElementBox(textTarget.element);
    if (!box) return;

    this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(textTarget.element, event.clientY, box);
    const offset = this.getInlineTextOffsetAtClientPoint(textTarget.element, editor, event.clientX, event.clientY, box);
    this.focusEditorWithoutCanvasScroll(editor);
    editor.setSelectionRange(offset, offset);
    this.rememberInlineCaretPlacement(editor, textTarget.element, offset);
    this.resetInlineEditorScroll(editor);
    this.updateInlineCaret(editor, textTarget.element);
    this.beginInlineSelectionDrag(editor, textTarget.element, offset, event);
  }

  private beginInlineSelectionDrag(
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement,
    anchorOffset: number,
    event: PointerEvent
  ): void {
    this.stopInlineSelectionDrag();

    const onPointerMove = (moveEvent: PointerEvent) => {
      if (moveEvent.pointerId !== event.pointerId) return;
      moveEvent.preventDefault();
      moveEvent.stopPropagation();
      this.clearBrowserTextSelection();
      this.extendInlineSelectionDrag(moveEvent);
    };
    const onPointerUp = (upEvent: PointerEvent) => {
      if (upEvent.pointerId !== event.pointerId) return;
      upEvent.preventDefault();
      upEvent.stopPropagation();
      this.clearBrowserTextSelection();
      this.extendInlineSelectionDrag(upEvent);
      this.stopInlineSelectionDrag();
    };
    const cleanup = () => {
      document.removeEventListener('pointermove', onPointerMove, true);
      document.removeEventListener('pointerup', onPointerUp, true);
      document.removeEventListener('pointercancel', onPointerUp, true);
    };

    this.inlineSelectionDrag = {
      editor,
      element,
      anchorOffset,
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      isSelecting: false,
      cleanup
    };
    document.addEventListener('pointermove', onPointerMove, true);
    document.addEventListener('pointerup', onPointerUp, true);
    document.addEventListener('pointercancel', onPointerUp, true);
  }

  private extendInlineSelectionDrag(event: PointerEvent): void {
    const drag = this.inlineSelectionDrag;
    if (!drag || this.activeEditor !== drag.editor || this.activeEditorTarget !== drag.element) return;

    const box = this.getElementBox(drag.element);
    if (!box) return;

    if (!drag.isSelecting) {
      if (!this.hasInlineSelectionDragMoved(drag, event)) {
        drag.editor.setSelectionRange(drag.anchorOffset, drag.anchorOffset);
        this.rememberInlineCaretPlacement(drag.editor, drag.element, drag.anchorOffset);
        this.updateInlineCaret(drag.editor, drag.element);
        return;
      }
      drag.isSelecting = true;
      this.lastInlineCaretPlacement = null;
    }

    this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(drag.element, event.clientY, box);
    const focusOffset = this.getInlineTextOffsetAtClientPoint(drag.element, drag.editor, event.clientX, event.clientY, box);
    const selectionStart = Math.min(drag.anchorOffset, focusOffset);
    const selectionEnd = Math.max(drag.anchorOffset, focusOffset);
    const direction = focusOffset < drag.anchorOffset ? 'backward' : 'forward';
    this.focusEditorWithoutCanvasScroll(drag.editor);
    drag.editor.setSelectionRange(selectionStart, selectionEnd, direction);
    this.resetInlineEditorScroll(drag.editor);
    this.updateInlineCaret(drag.editor, drag.element);
  }

  private hasInlineSelectionDragMoved(drag: InlineSelectionDrag, event: PointerEvent): boolean {
    const dx = event.clientX - drag.startClientX;
    const dy = event.clientY - drag.startClientY;
    return Math.hypot(dx, dy) >= 4;
  }

  private stopInlineSelectionDrag(): void {
    this.inlineSelectionDrag?.cleanup();
    this.inlineSelectionDrag = null;
  }

  private clearBrowserTextSelection(): void {
    document.getSelection()?.removeAllRanges();
  }

  private handleOutsideSlidePointerDown = (event: PointerEvent): void => {
    if (!this.activeEditor) return;

    const target = event.target instanceof Node ? event.target : null;
    if (target && this.slideSurface?.contains(target)) return;
    if (target && this.activeEditor.contains(target)) return;

    const clearSelectionAfterCommit = this.isCanvasPaneBackgroundTarget(target);
    if (clearSelectionAfterCommit) {
      event.preventDefault();
      event.stopPropagation();
    }
    this.clearBrowserTextSelection();
    this.commitActiveEditorFromOutside(clearSelectionAfterCommit);
  };

  private handleCanvasPanePointerDown = (event: PointerEvent): void => {
    if (event.button !== 0) return;

    const target = event.target instanceof Node ? event.target : null;
    if (!this.isCanvasPaneBackgroundTarget(target)) return;

    event.preventDefault();
    event.stopPropagation();
    this.clearBrowserTextSelection();
    if (this.activeEditor) {
      this.commitActiveEditorFromOutside(true);
      return;
    }

    this.clearSelection();
    this.removeInlineSelection();
    this.lastInlineCaretPlacement = null;
  };

  private isCanvasPaneBackgroundTarget(target: Node | null): boolean {
    if (!target || !this.canvasPane?.contains(target)) return false;
    if (this.slideSurface?.contains(target)) return false;
    if (this.activeEditor?.contains(target)) return false;

    const element = target instanceof Element ? target : target.parentElement;
    if (element?.closest('.native-powerpoint-selection-box')) return false;

    return true;
  }

  private commitActiveEditorFromOutside(clearSelectionAfterCommit: boolean): void {
    const commit = this.activeEditorCommit;
    if (commit) {
      void commit().finally(() => {
        this.clearBrowserTextSelection();
        if (clearSelectionAfterCommit) {
          this.clearSelection();
        }
      });
      return;
    }

    this.removeActiveEditor();
    if (clearSelectionAfterCommit) {
      this.clearSelection();
    }
  }

  private startTextEditor(target: TextEditTarget | null = null, clientX?: number, clientY?: number): void {
    if (!this.canvasPane || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('edit text')) return;

    const selected = this.getSelectedShapeElement();
    if (!selected) return;
    if (!target && selected.closest(GENERATED_GRID_SELECTOR)) {
      this.showGeneratedTextNotice();
      return;
    }

    if (target && this.activeEditor && this.activeEditorTarget === target.element) {
      const currentBox = this.getElementBox(target.element);
      if (currentBox) {
        this.placeInlineCaret(this.activeEditor, target.element, clientX, clientY, currentBox);
        this.focusEditorWithoutCanvasScroll(this.activeEditor);
        this.resetInlineEditorScroll(this.activeEditor);
      }
      return;
    }

    this.removeActiveEditor();
    const box = target ? this.getElementBox(target.element) : this.getSelectedBox();
    if (!box) return;

    const editor = this.canvasPane.createEl('textarea', {
      cls: `native-powerpoint-inline-editor${target ? ' is-text-run' : ''}`,
      attr: { 'aria-label': 'Edit selected text' }
    });
    const initialText = target?.text ?? selected.textContent?.trim() ?? '';
    editor.value = initialText;
    if (target) {
      const style = window.getComputedStyle(target.element);
      target.element.classList.add('native-powerpoint-text-editing');
      this.activeEditorTarget = target.element;
      this.slideSurface?.addClass('is-inline-text-editing');
      editor.setCssProps({
        color: style.fill,
        fontFamily: style.fontFamily,
        fontSize: `${this.getScreenFontSize(target.element)}px`,
        fontStyle: style.fontStyle,
        fontWeight: style.fontWeight,
        lineHeight: '1.1',
        textAlign: this.getInlineTextAlignment(style.textAnchor)
      });
    }
    if (target) {
      this.positionTextRunEditor(editor, box);
    } else {
      editor.setCssProps({
        left: `${box.left}px`,
        top: `${box.top}px`,
        width: `${Math.max(160, box.width)}px`,
        height: `${Math.max(72, box.height)}px`
      });
    }
    this.activeEditor = editor;
    if (target) {
      this.removeSelectionOverlay();
      this.activeInlineCaret = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      this.activeInlineCaret.classList.add('native-powerpoint-svg-caret');
      this.activeInlineCaret.setAttribute('aria-hidden', 'true');
      this.svgEl?.appendChild(this.activeInlineCaret);
      const updateCaret = () => {
        this.rememberCollapsedInlineCaretPlacement(editor, target.element);
        this.updateInlineCaret(editor, target.element);
      };
      const queueCaretUpdate = () => {
        window.requestAnimationFrame(() => {
          if (this.activeEditor === editor) {
            updateCaret();
          }
        });
      };
      editor.addEventListener('pointerdown', (event) => {
        event.preventDefault();
        const nextBox = this.getElementBox(target.element) ?? box;
        this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(target.element, event.clientY, nextBox);
        const offset = this.getInlineTextOffsetAtClientPoint(target.element, editor, event.clientX, event.clientY, nextBox);
        this.focusEditorWithoutCanvasScroll(editor);
        editor.setSelectionRange(offset, offset);
        this.rememberInlineCaretPlacement(editor, target.element, offset);
        this.resetInlineEditorScroll(editor);
        updateCaret();
      });
      editor.addEventListener('input', () => {
        if (this.activeEditor === editor) {
          target.element.textContent = editor.value;
          const nextBox = this.getElementBox(target.element);
          if (nextBox) {
            this.positionTextRunEditor(editor, nextBox);
          }
          updateCaret();
        }
      });
      editor.addEventListener('click', updateCaret);
      editor.addEventListener('keyup', updateCaret);
      editor.addEventListener('mouseup', updateCaret);
      editor.addEventListener('select', updateCaret);
      editor.addEventListener('keydown', (event) => {
        if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'a') {
          this.lastInlineCaretPlacement = null;
        }
        if (
          event.shiftKey
          && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)
        ) {
          this.lastInlineCaretPlacement = null;
        }
        if (this.handleInlineDeleteKey(event, editor, target.element)) return;
        queueCaretUpdate();
      });
    }

    const commit = async () => {
      if (this.activeEditor !== editor) return;
      this.removeActiveEditor(editor);
      await this.applyTextValue(editor.value, target);
    };
    this.activeEditorCommit = commit;

    editor.addEventListener('keydown', (event) => {
      if (
        event.key === 'Enter'
        && ((event.metaKey || event.ctrlKey) || (target && !event.shiftKey))
      ) {
        event.preventDefault();
        void commit();
      } else if (event.key === 'Escape') {
        if (target) {
          target.element.textContent = initialText;
        }
        this.removeActiveEditor(editor);
      }
    });
    editor.addEventListener('blur', () => void commit());
    this.focusEditorWithoutCanvasScroll(editor);
    if (target) {
      this.placeInlineCaret(editor, target.element, clientX, clientY, box);
    } else {
      this.selectEditorWithoutCanvasScroll(editor);
    }
  }

  private renderFontFidelity(container: HTMLElement): void {
    if (!this.engine) return;

    const section = container.createDiv({ cls: 'native-powerpoint-font-fidelity' });
    section.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: 'Fonts' });

    if (this.fontSubstitutions.length === 0) {
      section.createDiv({
        cls: 'native-powerpoint-inspector-hint',
        text: 'Requested fonts on this slide are available.'
      });
      return;
    }

    section.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: `${this.fontSubstitutions.length} missing font${this.fontSubstitutions.length === 1 ? '' : 's'} substituted. Text wrapping uses the displayed replacement font.`
    });
    const list = section.createDiv({ cls: 'native-powerpoint-font-substitution-list' });
    for (const substitution of this.fontSubstitutions) {
      const item = list.createDiv({ cls: 'native-powerpoint-font-substitution' });
      item.createSpan({ cls: 'native-powerpoint-font-substitution-source', text: substitution.requested });
      item.createSpan({ cls: 'native-powerpoint-font-substitution-arrow', text: '->' });
      item.createSpan({ cls: 'native-powerpoint-font-substitution-target', text: substitution.substitute });
    }
  }

  private removeActiveEditor(editor = this.activeEditor): void {
    if (editor && this.activeEditor && editor !== this.activeEditor) return;

    this.stopInlineSelectionDrag();
    this.activeEditor?.remove();
    this.activeInlineCaret?.remove();
    this.removeInlineSelection();
    this.activeEditor = null;
    this.activeEditorCommit = null;
    this.activeInlineCaret = null;
    this.lastInlineCaretPlacement = null;
    this.activeInlineCaretRow = null;
    this.activeEditorTarget?.classList.remove('native-powerpoint-text-editing');
    this.activeEditorTarget = null;
    this.slideSurface?.removeClass('is-inline-text-editing');
    if (this.svgEl?.isConnected) {
      this.updateSelectionOverlay();
    }
  }

  private captureCanvasScroll(): CanvasScrollPosition | null {
    if (!this.canvasPane) return null;
    return {
      left: this.canvasPane.scrollLeft,
      top: this.canvasPane.scrollTop
    };
  }

  private restoreCanvasScroll(position: CanvasScrollPosition | null): void {
    if (!position || !this.canvasPane) return;
    this.canvasPane.scrollLeft = position.left;
    this.canvasPane.scrollTop = position.top;
  }

  private restoreCanvasScrollSoon(position: CanvasScrollPosition | null): void {
    this.restoreCanvasScroll(position);
    if (!position) return;

    window.requestAnimationFrame(() => this.restoreCanvasScroll(position));
    window.setTimeout(() => this.restoreCanvasScroll(position), 0);
  }

  private focusEditorWithoutCanvasScroll(editor: HTMLTextAreaElement): void {
    const scrollPosition = this.captureCanvasScroll();
    editor.focus({ preventScroll: true });
    this.restoreCanvasScrollSoon(scrollPosition);
  }

  private selectEditorWithoutCanvasScroll(editor: HTMLTextAreaElement): void {
    const scrollPosition = this.captureCanvasScroll();
    editor.select();
    this.restoreCanvasScrollSoon(scrollPosition);
  }

  private getScreenFontSize(element: SVGTextElement | SVGTSpanElement): number {
    const style = window.getComputedStyle(element);
    const fontSize = Number.parseFloat(style.fontSize) || 16;
    const matrix = element.getScreenCTM();
    const scale = matrix ? Math.hypot(matrix.a, matrix.b) : 1;
    return Math.max(4, fontSize * scale);
  }

  private getInlineTextAlignment(textAnchor: string): string {
    if (textAnchor === 'middle') return 'center';
    if (textAnchor === 'end') return 'right';
    return 'left';
  }

  private positionTextRunEditor(
    editor: HTMLTextAreaElement,
    box: { left: number; top: number; width: number; height: number }
  ): void {
    editor.setCssProps({
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: '1px',
      height: '1px'
    });
  }

  private placeInlineCaret(
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement,
    clientX: number | undefined,
    clientY: number | undefined,
    box: { left: number; top: number; width: number; height: number }
  ): void {
    this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(element, clientY, box);
    const text = editor.value;
    if (clientX === undefined || box.width <= 0 || text.length === 0) {
      editor.setSelectionRange(text.length, text.length);
      this.rememberInlineCaretPlacement(editor, element, text.length);
      this.updateInlineCaret(editor, element);
      return;
    }

    const offset = this.getInlineTextOffsetAtClientPoint(element, editor, clientX, clientY, box);
    editor.setSelectionRange(offset, offset);
    this.rememberInlineCaretPlacement(editor, element, offset);
    this.resetInlineEditorScroll(editor);
    this.updateInlineCaret(editor, element);
  }

  private rememberInlineCaretPlacement(
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement,
    offset: number
  ): void {
    this.lastInlineCaretPlacement = {
      editor,
      element,
      offset: Math.max(0, Math.min(offset, editor.value.length)),
      timestamp: Date.now()
    };
  }

  private handleInlineDeleteKey(
    event: KeyboardEvent,
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement
  ): boolean {
    if (event.key !== 'Delete' && event.key !== 'Backspace') return false;
    if (event.metaKey || event.ctrlKey || event.altKey) return false;

    const text = editor.value;
    if (!text) return false;

    const selectionStart = Math.max(0, Math.min(editor.selectionStart ?? text.length, text.length));
    const selectionEnd = Math.max(0, Math.min(editor.selectionEnd ?? text.length, text.length));
    if (selectionStart !== 0 || selectionEnd !== text.length) return false;

    const placement = this.lastInlineCaretPlacement;
    if (
      !placement
      || placement.editor !== editor
      || placement.element !== element
      || Date.now() - placement.timestamp > 30000
    ) {
      return false;
    }

    const caretOffset = Math.max(0, Math.min(placement.offset, text.length));
    const deleteStart = event.key === 'Backspace' ? caretOffset - 1 : caretOffset;
    const deleteEnd = event.key === 'Backspace' ? caretOffset : caretOffset + 1;

    event.preventDefault();
    event.stopPropagation();

    if (deleteStart < 0 || deleteEnd > text.length || deleteStart >= deleteEnd) {
      editor.setSelectionRange(caretOffset, caretOffset);
      this.rememberInlineCaretPlacement(editor, element, caretOffset);
      this.updateInlineCaret(editor, element);
      return true;
    }

    const nextText = text.slice(0, deleteStart) + text.slice(deleteEnd);
    editor.value = nextText;
    element.textContent = nextText;
    editor.setSelectionRange(deleteStart, deleteStart);
    this.rememberInlineCaretPlacement(editor, element, deleteStart);
    this.resetInlineEditorScroll(editor);
    const nextBox = this.getElementBox(element);
    if (nextBox) {
      this.positionTextRunEditor(editor, nextBox);
    }
    this.updateInlineCaret(editor, element);
    return true;
  }

  private rememberCollapsedInlineCaretPlacement(
    editor: HTMLTextAreaElement,
    element: SVGTextElement | SVGTSpanElement
  ): void {
    const textLength = editor.value.length;
    const selectionStart = Math.max(0, Math.min(editor.selectionStart ?? textLength, textLength));
    const selectionEnd = Math.max(0, Math.min(editor.selectionEnd ?? textLength, textLength));
    if (selectionStart === selectionEnd) {
      this.rememberInlineCaretPlacement(editor, element, selectionEnd);
    }
  }

  private updateInlineCaret(editor: HTMLTextAreaElement, element: SVGTextElement | SVGTSpanElement): void {
    if (!this.activeInlineCaret) return;

    const box = this.getElementBox(element);
    if (!box) {
      this.activeInlineCaret.addClass('native-powerpoint-inline-caret-hidden');
      return;
    }

    this.refreshActiveInlineCaretRow(element, box);
    this.updateInlineSelection(editor, element);

    const selectionStart = Math.max(0, Math.min(editor.selectionStart ?? editor.value.length, editor.value.length));
    const selectionEnd = Math.max(0, Math.min(editor.selectionEnd ?? editor.value.length, editor.value.length));
    if (selectionStart !== selectionEnd) {
      this.activeInlineCaret.addClass('native-powerpoint-inline-caret-hidden');
      return;
    }

    const offset = selectionEnd;
    const geometry = this.getSvgInlineCaretGeometry(element, editor, offset, box);
    if (!geometry) {
      this.activeInlineCaret.addClass('native-powerpoint-inline-caret-hidden');
      return;
    }

    this.activeInlineCaret.removeClass('native-powerpoint-inline-caret-hidden');
    this.activeInlineCaret.setAttribute('x1', this.formatSvgNumber(geometry.x1));
    this.activeInlineCaret.setAttribute('y1', this.formatSvgNumber(geometry.y1));
    this.activeInlineCaret.setAttribute('x2', this.formatSvgNumber(geometry.x2));
    this.activeInlineCaret.setAttribute('y2', this.formatSvgNumber(geometry.y2));
    this.activeInlineCaret.setAttribute('stroke-width', this.formatSvgNumber(geometry.strokeWidth));
  }

  private updateInlineSelection(editor: HTMLTextAreaElement, element: SVGTextElement | SVGTSpanElement): void {
    this.removeInlineSelection();

    const selectionStart = Math.max(0, Math.min(editor.selectionStart ?? editor.value.length, editor.value.length));
    const selectionEnd = Math.max(0, Math.min(editor.selectionEnd ?? editor.value.length, editor.value.length));
    const start = Math.min(selectionStart, selectionEnd);
    const end = Math.max(selectionStart, selectionEnd);
    if (start === end) return;

    const boxes = this.getSvgInlineSelectionBoxes(element, start, end);
    const textElement = element instanceof SVGTextElement ? element : element.closest('text');
    const parent = textElement?.parentNode;
    if (!textElement || !parent) return;

    for (const box of boxes) {
      const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect.classList.add('native-powerpoint-svg-selection');
      rect.setAttribute('x', this.formatSvgNumber(box.x));
      rect.setAttribute('y', this.formatSvgNumber(box.y));
      rect.setAttribute('width', this.formatSvgNumber(box.width));
      rect.setAttribute('height', this.formatSvgNumber(box.height));
      parent.insertBefore(rect, textElement);
      this.activeInlineSelectionRects.push(rect);
    }
  }

  private removeInlineSelection(): void {
    for (const rect of this.activeInlineSelectionRects) {
      rect.remove();
    }
    this.activeInlineSelectionRects = [];
  }

  private getSvgInlineCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    editor: HTMLTextAreaElement,
    offset: number,
    box: { left: number; top: number; width: number; height: number }
  ): SvgInlineCaretGeometry | null {
    const screenGeometry = this.getInlineCaretGeometry(element, editor, offset, box);
    const top = this.activeInlineCaretRow?.top ?? screenGeometry.top;
    const height = this.activeInlineCaretRow?.height ?? screenGeometry.height;
    const start = this.localPointToSvgRoot(screenGeometry.left, top);
    const end = this.localPointToSvgRoot(screenGeometry.left, top + height);
    if (!start || !end) return null;

    return {
      x1: start.x,
      y1: start.y,
      x2: end.x,
      y2: end.y,
      strokeWidth: this.getSvgInlineCaretStrokeWidth(element)
    };
  }

  private getSvgInlineSelectionBoxes(
    element: SVGTextElement | SVGTSpanElement,
    start: number,
    end: number
  ): SvgInlineSelectionBox[] {
    const textElement = element as SVGTextContentElement;
    const elementMatrix = element.getScreenCTM();
    const rootMatrix = this.svgEl?.getScreenCTM();
    if (!elementMatrix || !rootMatrix) return [];

    let rootInverse: DOMMatrix;
    try {
      rootInverse = rootMatrix.inverse();
    } catch {
      return [];
    }

    let charCount = 0;
    try {
      charCount = textElement.getNumberOfChars();
    } catch {
      return [];
    }
    if (charCount <= 0) return [];

    const normalizedStart = Math.max(0, Math.min(charCount, start));
    const normalizedEnd = Math.max(normalizedStart, Math.min(charCount, end));
    const rows: SvgInlineSelectionBox[] = [];
    for (let index = normalizedStart; index < normalizedEnd; index++) {
      let charBox: SvgInlineSelectionBox | null = null;
      try {
        charBox = this.transformSvgRectToSvgRoot(textElement.getExtentOfChar(index), elementMatrix, rootInverse);
      } catch {
        charBox = null;
      }
      if (!charBox || charBox.width < 0 || charBox.height <= 0) continue;

      const centerY = charBox.y + charBox.height / 2;
      const row = rows.find((candidate) => (
        Math.abs(centerY - (candidate.y + candidate.height / 2)) < Math.max(2, charBox.height * 0.55)
      ));
      if (row) {
        const left = Math.min(row.x, charBox.x);
        const top = Math.min(row.y, charBox.y);
        const right = Math.max(row.x + row.width, charBox.x + charBox.width);
        const bottom = Math.max(row.y + row.height, charBox.y + charBox.height);
        row.x = left;
        row.y = top;
        row.width = right - left;
        row.height = bottom - top;
      } else {
        rows.push({ ...charBox });
      }
    }

    const padding = this.getSvgInlineSelectionPadding(element);
    return rows.map((box) => ({
      x: box.x - padding,
      y: box.y - padding * 0.5,
      width: box.width + padding * 2,
      height: box.height + padding
    }));
  }

  private getInlineTextOffsetAtClientPoint(
    element: SVGTextElement | SVGTSpanElement,
    editor: HTMLTextAreaElement,
    clientX: number,
    clientY: number | undefined,
    box: { left: number; top: number; width: number; height: number }
  ): number {
    const text = editor.value;
    if (text.length === 0) return 0;

    const paneRect = this.canvasPane?.getBoundingClientRect();
    const localClientX = paneRect
      ? clientX - paneRect.left + (this.canvasPane?.scrollLeft ?? 0)
      : clientX;
    const localClientY = paneRect && clientY !== undefined
      ? clientY - paneRect.top + (this.canvasPane?.scrollTop ?? 0)
      : box.top + box.height / 2;

    const geometryOffset = this.getInlineTextOffsetFromSvgGeometry(element, localClientX, localClientY, text.length);
    if (geometryOffset !== null) {
      return geometryOffset;
    }

    return this.getMeasuredInlineTextOffset(editor, localClientX, box);
  }

  private getInlineTextOffsetFromSvgGeometry(
    element: SVGTextElement | SVGTSpanElement,
    localClientX: number,
    localClientY: number,
    textLength: number
  ): number | null {
    const textElement = element as SVGTextContentElement;
    let charCount = 0;
    try {
      charCount = textElement.getNumberOfChars();
    } catch {
      return null;
    }
    if (charCount <= 0) return null;

    const maxOffset = Math.min(textLength, charCount);
    let bestOffset = 0;
    let bestDistance = Number.POSITIVE_INFINITY;
    for (let offset = 0; offset <= maxOffset; offset++) {
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

    return Number.isFinite(bestDistance) ? bestOffset : null;
  }

  private getInlineCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    editor: HTMLTextAreaElement,
    offset: number,
    box: { left: number; top: number; width: number; height: number }
  ): { left: number; top: number; height: number } {
    const fallbackHeight = this.getInlineCaretHeight(element, box);
    const svgGeometry = this.getSvgTextCaretGeometry(element, offset, fallbackHeight);
    if (svgGeometry) {
      const row = this.activeInlineCaretRow ?? svgGeometry;
      return { left: svgGeometry.left, top: row.top, height: row.height };
    }

    const row = this.activeInlineCaretRow ?? this.getDefaultInlineCaretRow(element, box, fallbackHeight);

    const text = editor.value;
    if (text.length === 0) return { left: box.left, ...row };

    const style = window.getComputedStyle(editor);
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) {
      return { left: box.left + box.width * (offset / text.length), ...row };
    }

    context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    const fullWidth = context.measureText(text).width;
    if (fullWidth <= 0) return { left: box.left, ...row };

    return {
      left: box.left + context.measureText(text.slice(0, offset)).width * (box.width / fullWidth),
      ...row
    };
  }

  private getInlineCaretRowFromClientY(
    element: SVGTextElement | SVGTSpanElement,
    clientY: number | undefined,
    box: { left: number; top: number; width: number; height: number }
  ): InlineCaretRow {
    const height = this.getInlineCaretHeight(element, box);
    const paneRect = this.canvasPane?.getBoundingClientRect();
    if (clientY === undefined || !paneRect) {
      return this.getDefaultInlineCaretRow(element, box, height);
    }

    const localY = clientY - paneRect.top + (this.canvasPane?.scrollTop ?? 0);
    if (box.height <= height * 1.8) {
      return this.getDefaultInlineCaretRow(element, box, height);
    }

    const centerRatio = box.height > 0 ? (localY - box.top) / box.height : 0.5;
    return this.getInlineCaretRowFromRatio(element, box, centerRatio, height);
  }

  private getInlineCaretHeight(element: SVGTextElement | SVGTSpanElement, box: { width: number; height: number }): number {
    const lineCount = this.estimateInlineTextRowCount(element);
    const lineBoxHeight = box.height / Math.max(1, lineCount);
    const screenFontSize = Math.min(this.getScreenFontSize(element), lineBoxHeight);
    const baseHeight = Math.min(lineBoxHeight, screenFontSize || lineBoxHeight);
    return Math.max(6, baseHeight * 0.88);
  }

  private getSvgInlineCaretStrokeWidth(element: SVGTextElement | SVGTSpanElement): number {
    const style = window.getComputedStyle(element);
    const fontSize = Number.parseFloat(style.fontSize) || 16;
    return Math.max(1.25, Math.min(4, fontSize / 14));
  }

  private getSvgInlineSelectionPadding(element: SVGTextElement | SVGTSpanElement): number {
    const style = window.getComputedStyle(element);
    const fontSize = Number.parseFloat(style.fontSize) || 16;
    return Math.max(0.75, Math.min(3, fontSize / 18));
  }

  private refreshActiveInlineCaretRow(
    element: SVGTextElement | SVGTSpanElement,
    box: { left: number; top: number; width: number; height: number }
  ): void {
    if (!this.activeInlineCaretRow) return;

    this.activeInlineCaretRow = this.getInlineCaretRowFromRatio(
      element,
      box,
      this.activeInlineCaretRow.centerRatio
    );
  }

  private estimateInlineTextRowCount(element: SVGTextElement | SVGTSpanElement): number {
    const text = element.textContent || '';
    if (!text) return 1;

    const textElement = element as SVGTextContentElement;
    const matrix = element.getScreenCTM();
    const paneRect = this.canvasPane?.getBoundingClientRect();
    if (!matrix || !paneRect) return 1;

    try {
      const rows: number[] = [];
      const charCount = textElement.getNumberOfChars();
      for (let index = 0; index < charCount; index++) {
        const position = textElement.getStartPositionOfChar(index);
        const point = new DOMPoint(position.x, position.y).matrixTransform(matrix);
        const localY = point.y - paneRect.top + (this.canvasPane?.scrollTop ?? 0);
        if (!rows.some((row) => Math.abs(row - localY) < 4)) {
          rows.push(localY);
        }
      }
      return Math.max(1, rows.length);
    } catch {
      return 1;
    }
  }

  private getDefaultInlineCaretRow(
    _element: SVGTextElement | SVGTSpanElement,
    box: { left: number; top: number; width: number; height: number },
    height: number
  ): InlineCaretRow {
    return this.getInlineCaretRowFromRatio(_element, box, 0.5, height);
  }

  private getInlineCaretRowFromRatio(
    element: SVGTextElement | SVGTSpanElement,
    box: { left: number; top: number; width: number; height: number },
    centerRatio: number,
    height = this.getInlineCaretHeight(element, box)
  ): InlineCaretRow {
    const ratio = Math.max(0, Math.min(1, centerRatio));
    const minCenter = box.top + height / 2;
    const maxCenter = Math.max(minCenter, box.top + box.height - height / 2);
    const center = Math.max(minCenter, Math.min(maxCenter, box.top + box.height * ratio));
    return {
      top: center - height / 2,
      height,
      centerRatio: ratio
    };
  }

  private getMeasuredInlineTextOffset(
    editor: HTMLTextAreaElement,
    localClientX: number,
    box: { left: number; top: number; width: number; height: number }
  ): number {
    const text = editor.value;
    const clickOffset = Math.max(0, Math.min(box.width, localClientX - box.left));
    const style = window.getComputedStyle(editor);
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) {
      return Math.round(text.length * (clickOffset / box.width));
    }

    context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    const measuredWidth = context.measureText(text).width;
    const widthScale = measuredWidth > 0 ? box.width / measuredWidth : 1;
    let previousWidth = 0;
    for (let offset = 1; offset <= text.length; offset++) {
      const width = context.measureText(text.slice(0, offset)).width * widthScale;
      if (clickOffset <= (previousWidth + width) / 2) {
        return offset - 1;
      }
      previousWidth = width;
    }

    return text.length;
  }

  private resetInlineEditorScroll(editor: HTMLTextAreaElement): void {
    editor.scrollLeft = 0;
    editor.scrollTop = 0;
  }

  private getSvgTextCaretLeft(element: SVGTextElement | SVGTSpanElement, offset: number): number | null {
    return this.getSvgTextCaretGeometry(element, offset)?.left ?? null;
  }

  private getSvgTextCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    offset: number,
    preferredHeight = this.getScreenFontSize(element) * 1.08
  ): { left: number; top: number; height: number } | null {
    const text = element.textContent || '';
    if (!text) return this.getFallbackInlineCaretGeometry(element, preferredHeight);

    const textElement = element as SVGTextContentElement;
    const matrix = element.getScreenCTM();
    const paneRect = this.canvasPane?.getBoundingClientRect();
    if (!matrix || !paneRect) return null;

    let position: DOMPoint;
    let extent: DOMRect | SVGRect | null = null;
    try {
      const charCount = textElement.getNumberOfChars();
      if (charCount <= 0) return this.getFallbackInlineCaretGeometry(element, preferredHeight);
      const normalizedOffset = Math.max(0, Math.min(charCount, offset));
      const charIndex = Math.max(0, Math.min(charCount - 1, normalizedOffset <= 0 ? 0 : normalizedOffset - 1));
      position = normalizedOffset <= 0
        ? textElement.getStartPositionOfChar(charIndex)
        : textElement.getEndPositionOfChar(charIndex);
      extent = textElement.getExtentOfChar(charIndex);
    } catch {
      return null;
    }

    const point = new DOMPoint(position.x, position.y).matrixTransform(matrix);
    const localLeft = point.x - paneRect.left + (this.canvasPane?.scrollLeft ?? 0);
    const fallbackBox = this.getElementBox(element);
    let top = fallbackBox ? fallbackBox.top + Math.max(0, (fallbackBox.height - preferredHeight) / 2) : 0;
    let height = Math.max(6, preferredHeight);

    if (extent) {
      const bounds = this.transformSvgRectToLocalBox(extent, matrix, paneRect);
      if (bounds) {
        height = Math.max(6, Math.min(preferredHeight * 1.1, bounds.height || preferredHeight));
        top = bounds.top + Math.max(0, (bounds.height - height) / 2);
      }
    }

    return { left: localLeft, top, height };
  }

  private getFallbackInlineCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    preferredHeight: number
  ): { left: number; top: number; height: number } | null {
    const box = this.getElementBox(element);
    if (!box) return null;

    const height = Math.max(6, preferredHeight);
    return {
      left: box.left,
      top: box.top + Math.max(0, (box.height - height) / 2),
      height
    };
  }

  private localPointToSvgRoot(left: number, top: number): DOMPoint | null {
    if (!this.canvasPane || !this.svgEl) return null;

    const matrix = this.svgEl.getScreenCTM();
    const paneRect = this.canvasPane.getBoundingClientRect();
    if (!matrix || !paneRect) return null;

    const screenPoint = new DOMPoint(
      paneRect.left + left - this.canvasPane.scrollLeft,
      paneRect.top + top - this.canvasPane.scrollTop
    );

    try {
      return screenPoint.matrixTransform(matrix.inverse());
    } catch {
      return null;
    }
  }

  private transformSvgRectToSvgRoot(
    rect: DOMRect | SVGRect,
    elementMatrix: DOMMatrix,
    rootInverse: DOMMatrix
  ): SvgInlineSelectionBox | null {
    const points = [
      new DOMPoint(rect.x, rect.y),
      new DOMPoint(rect.x + rect.width, rect.y),
      new DOMPoint(rect.x, rect.y + rect.height),
      new DOMPoint(rect.x + rect.width, rect.y + rect.height)
    ].map((point) => point.matrixTransform(elementMatrix).matrixTransform(rootInverse));
    if (points.some((point) => !Number.isFinite(point.x) || !Number.isFinite(point.y))) {
      return null;
    }

    const xs = points.map((point) => point.x);
    const ys = points.map((point) => point.y);
    const x = Math.min(...xs);
    const y = Math.min(...ys);
    return {
      x,
      y,
      width: Math.max(...xs) - x,
      height: Math.max(...ys) - y
    };
  }

  private formatSvgNumber(value: number): string {
    return `${Math.round(value * 1000) / 1000}`;
  }

  private transformSvgRectToLocalBox(
    rect: DOMRect | SVGRect,
    matrix: DOMMatrix,
    paneRect: DOMRect
  ): { left: number; top: number; width: number; height: number } | null {
    const scrollLeft = this.canvasPane?.scrollLeft ?? 0;
    const scrollTop = this.canvasPane?.scrollTop ?? 0;
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

  private getSelectedShapeElement(): SVGGElement | null {
    if (!this.svgEl || this.selectedShapeIndex === null) return null;
    const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${this.selectedShapeIndex}"]`);
    return shape instanceof SVGGElement ? shape : null;
  }

  private getTextEditTarget(element: Element | null): TextEditTarget | null {
    const directRun = element?.closest('tspan[data-ooxml-run-idx]') ?? null;
    const textRuns = Array.from(
      element?.closest('text')?.querySelectorAll('tspan[data-ooxml-run-idx]') ?? []
    );
    const run = directRun ?? (textRuns.length === 1 ? textRuns[0] : null);
    if (!(run instanceof SVGTSpanElement) || run.closest(GENERATED_GRID_SELECTOR)) return null;

    const paragraph = run.closest('tspan[data-ooxml-para-idx]');
    const shape = run.closest('g[data-ooxml-shape-idx]');
    const shapeIndex = getShapeIndex(shape);
    const paragraphIndex = Number(paragraph?.getAttribute('data-ooxml-para-idx'));
    const runIndex = Number(run.getAttribute('data-ooxml-run-idx'));
    if (
      shapeIndex === null
      || !Number.isFinite(paragraphIndex)
      || !Number.isFinite(runIndex)
    ) {
      return null;
    }

    return {
      kind: 'shape-run',
      shapeIndex,
      paragraphIndex,
      runIndex,
      text: run.textContent || '',
      element: run
    };
  }

  private getGeneratedTextEditTarget(element: Element | null): GeneratedTextEditTarget | null {
    if (!this.engine) return null;

    const textElement = element?.closest('text');
    const shape = textElement?.closest('g[data-ooxml-shape-idx]');
    const shapeIndex = getShapeIndex(shape ?? null);
    const kind = textElement?.getAttribute('data-native-powerpoint-generated-kind') as GeneratedTextKind | null;
    const labelIndex = Number(textElement?.getAttribute('data-native-powerpoint-label-index'));
    const occurrence = Number(textElement?.getAttribute('data-native-powerpoint-label-occurrence'));
    if (
      !(textElement instanceof SVGTextElement)
      || shapeIndex === null
      || (kind !== 'chart' && kind !== 'table')
      || !Number.isFinite(labelIndex)
      || !Number.isFinite(occurrence)
    ) {
      return null;
    }

    const target: GeneratedTextEditTarget = {
      kind,
      shapeIndex,
      labelIndex,
      occurrence,
      previousText: textElement.textContent || '',
      text: textElement.textContent || '',
      element: textElement
    };
    return this.engine.canUpdateGeneratedText(this.currentSlide, shapeIndex, target) ? target : null;
  }

  private markGeneratedTextEditability(svg: SVGSVGElement): void {
    svg.querySelectorAll('text[data-native-powerpoint-generated-kind]').forEach((text) => {
      if (this.getGeneratedTextEditTarget(text)) {
        text.classList.add('native-powerpoint-editable-text');
      } else {
        text.classList.add('native-powerpoint-generated-readonly');
      }
    });
  }

  private showGeneratedTextNotice(): void {
    if (this.hasShownGeneratedTextNotice) return;

    this.hasShownGeneratedTextNotice = true;
    new Notice(
      'This chart label is generated from numeric scale or data. Edit a highlighted chart title, legend, or category label instead.'
    );
  }

  private updateSelectionOverlay(): void {
    if (!this.canvasPane || this.selectedShapeIndex === null) {
      this.removeSelectionOverlay();
      return;
    }

    const box = this.getSelectedBox();
    if (!box) {
      this.removeSelectionOverlay();
      return;
    }

    if (!this.selectionOverlay) {
      this.selectionOverlay = this.canvasPane.createDiv({ cls: 'native-powerpoint-selection-box' });
      if (this.canEdit()) {
        for (const handle of ['nw', 'ne', 'sw', 'se'] as HandleName[]) {
          const handleEl = this.selectionOverlay.createDiv({ cls: `native-powerpoint-resize-handle native-powerpoint-resize-${handle}` });
          handleEl.addEventListener('pointerdown', (event) => {
            event.preventDefault();
            event.stopPropagation();
            this.startDrag(event, 'resize', handle);
          });
        }
      }
    }

    this.selectionOverlay.setCssProps({
      left: `${box.left}px`,
      top: `${box.top}px`,
      width: `${box.width}px`,
      height: `${box.height}px`
    });
  }

  private removeSelectionOverlay(): void {
    this.selectionOverlay?.remove();
    this.selectionOverlay = null;
  }

  private getSelectedBox(): { left: number; top: number; width: number; height: number } | null {
    const selected = this.getSelectedShapeElement();
    if (!selected) return null;

    return this.getElementBox(selected);
  }

  private getElementBox(element: Element): { left: number; top: number; width: number; height: number } | null {
    if (!this.canvasPane) return null;

    const paneRect = this.canvasPane.getBoundingClientRect();
    const shapeRect = element.getBoundingClientRect();
    return {
      left: shapeRect.left - paneRect.left + this.canvasPane.scrollLeft,
      top: shapeRect.top - paneRect.top + this.canvasPane.scrollTop,
      width: shapeRect.width,
      height: shapeRect.height
    };
  }

  private startDrag(event: PointerEvent, mode: 'move' | 'resize', handle?: HandleName): void {
    if (!this.engine || !this.svgEl || this.selectedTransform === null) return;
    if (!this.ensureEditable(mode === 'move' ? 'move object' : 'resize object')) return;

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
      startTransform: cloneTransform(this.selectedTransform),
      latestTransform: cloneTransform(this.selectedTransform)
    };
  }

  private handleDragMove = (event: PointerEvent): void => {
    if (!this.dragState || !this.engine || !this.svgEl) return;
    if (event.pointerId !== this.dragState.pointerId) return;

    const point = this.getSvgPoint(event);
    if (!point) return;

    const scale = this.engine.getSlideScale(this.svgEl);
    const dx = (point.x - this.dragState.startPoint.x) * scale;
    const dy = (point.y - this.dragState.startPoint.y) * scale;
    const next = cloneTransform(this.dragState.startTransform);

    if (this.dragState.mode === 'move') {
      next.x += dx;
      next.y += dy;
    } else {
      const minSize = this.engine.pxToEmu(12);
      if (this.dragState.handle?.includes('w')) {
        next.x += dx;
        next.cx -= dx;
      }
      if (this.dragState.handle?.includes('e')) {
        next.cx += dx;
      }
      if (this.dragState.handle?.includes('n')) {
        next.y += dy;
        next.cy -= dy;
      }
      if (this.dragState.handle?.includes('s')) {
        next.cy += dy;
      }
      next.cx = Math.max(minSize, next.cx);
      next.cy = Math.max(minSize, next.cy);
    }

    this.dragState.latestTransform = next;
    this.selectedTransform = cloneTransform(next);
    this.updateInspectorValues();
    this.updateSelectionOverlayDuringDrag(event);
  };

  private handleDragEnd = (event: PointerEvent): void => {
    if (!this.dragState || event.pointerId !== this.dragState.pointerId) return;

    const transform = cloneTransform(this.dragState.latestTransform);
    this.dragState = null;
    void this.commitTransform(transform);
  };

  private async commitTransform(transform: ShapeTransform): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('edit object')) return;

    try {
      const selected = this.getSelectedShapeElement();
      if (selected && transformsMatch(this.engine.getShapeTransform(selected), transform)) return;

      const history = await this.captureHistoryEntry('Edit layout');
      this.engine.updateShapeTransform(this.currentSlide, this.selectedShapeIndex, transform);
      this.selectedTransform = cloneTransform(transform);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not update object: ${cleanError(error)}`);
    }
  }

  private getSvgPoint(event: PointerEvent): PointerPoint | null {
    if (!this.svgEl) return null;

    const matrix = this.svgEl.getScreenCTM();
    if (!matrix) return null;

    const point = this.svgEl.createSVGPoint();
    point.x = event.clientX;
    point.y = event.clientY;
    const result = point.matrixTransform(matrix.inverse());
    return { x: result.x, y: result.y };
  }

  private updateSelectionOverlayDuringDrag(event: PointerEvent): void {
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

  private updateInspectorValues(): void {
    if (!this.engine || !this.selectedTransform) return;

    if (this.xInput) this.xInput.value = String(Math.round(this.engine.emuToPx(this.selectedTransform.x) * 100) / 100);
    if (this.yInput) this.yInput.value = String(Math.round(this.engine.emuToPx(this.selectedTransform.y) * 100) / 100);
    if (this.widthInput) this.widthInput.value = String(Math.round(this.engine.emuToPx(this.selectedTransform.cx) * 100) / 100);
    if (this.heightInput) this.heightInput.value = String(Math.round(this.engine.emuToPx(this.selectedTransform.cy) * 100) / 100);
    if (this.rotationInput) this.rotationInput.value = String(Math.round(this.engine.ooxmlToDegrees(this.selectedTransform.rot) * 100) / 100);
  }

  private updateSlideScale(): void {
    if (!this.canvasPane || !this.slideSurface || !this.svgEl) return;

    const size = getSvgIntrinsicSize(this.svgEl);
    if (!size) {
      this.updateSelectionOverlay();
      return;
    }

    const computedStyle = window.getComputedStyle(this.canvasPane);
    const horizontalPadding =
      (Number.parseFloat(computedStyle.paddingLeft) || 0) +
      (Number.parseFloat(computedStyle.paddingRight) || 0);
    const verticalPadding =
      (Number.parseFloat(computedStyle.paddingTop) || 0) +
      (Number.parseFloat(computedStyle.paddingBottom) || 0);
    const availableWidth = Math.max(1, this.canvasPane.clientWidth - horizontalPadding);
    const availableHeight = Math.max(1, this.canvasPane.clientHeight - verticalPadding);
    const fitScale = Math.min(1, availableWidth / size.width, availableHeight / size.height);
    const scale = Math.max(0.05, fitScale * this.zoomLevel);
    const width = Math.max(1, Math.floor(size.width * scale));
    const height = Math.max(1, Math.floor(size.height * scale));

    this.slideSurface.addClass('is-scaled');
    this.slideSurface.style.setProperty('--native-powerpoint-slide-width', `${width}px`);
    this.slideSurface.style.setProperty('--native-powerpoint-slide-height', `${height}px`);
    this.updateSelectionOverlay();
    this.refreshActiveInlineEditorGeometry();
  }

  private handleCanvasWheel(event: WheelEvent): void {
    if (!this.canvasPane || !this.slideSurface || !this.svgEl || !this.engine) return;
    if (!this.isActivePowerPointView()) return;

    event.preventDefault();
    event.stopPropagation();

    const delta = this.normalizeWheelDelta(event);
    if (delta === 0) return;

    const nextZoom = this.zoomLevel * Math.pow(2, -delta / 600);
    this.setZoom(nextZoom, { clientX: event.clientX, clientY: event.clientY });
  }

  private normalizeWheelDelta(event: WheelEvent): number {
    if (event.deltaMode === WheelEvent.DOM_DELTA_LINE) {
      return event.deltaY * 16;
    }

    if (event.deltaMode === WheelEvent.DOM_DELTA_PAGE) {
      return event.deltaY * Math.max(1, this.canvasPane?.clientHeight ?? 800);
    }

    return event.deltaY;
  }

  private setZoom(value: number, anchor?: { clientX: number; clientY: number }): void {
    const nextZoom = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, Math.round(value * 1000) / 1000));
    if (nextZoom === this.zoomLevel) return;

    const anchorState = anchor ? this.captureZoomAnchor(anchor) : null;
    this.zoomLevel = nextZoom;
    this.updateSlideScale();
    if (anchorState) {
      this.restoreZoomAnchor(anchorState);
    }
  }

  private captureZoomAnchor(anchor: { clientX: number; clientY: number }): {
    clientX: number;
    clientY: number;
    ratioX: number;
    ratioY: number;
  } | null {
    if (!this.slideSurface) return null;

    const rect = this.slideSurface.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) return null;

    return {
      clientX: anchor.clientX,
      clientY: anchor.clientY,
      ratioX: Math.max(0, Math.min(1, (anchor.clientX - rect.left) / rect.width)),
      ratioY: Math.max(0, Math.min(1, (anchor.clientY - rect.top) / rect.height))
    };
  }

  private restoreZoomAnchor(anchor: { clientX: number; clientY: number; ratioX: number; ratioY: number }): void {
    if (!this.canvasPane || !this.slideSurface) return;

    const rect = this.slideSurface.getBoundingClientRect();
    const nextX = rect.left + rect.width * anchor.ratioX;
    const nextY = rect.top + rect.height * anchor.ratioY;
    this.canvasPane.scrollLeft += nextX - anchor.clientX;
    this.canvasPane.scrollTop += nextY - anchor.clientY;
    this.updateSelectionOverlay();
    this.refreshActiveInlineEditorGeometry();
  }

  private refreshActiveInlineEditorGeometry(): void {
    if (!this.activeEditor || !this.activeEditorTarget) return;

    const box = this.getElementBox(this.activeEditorTarget);
    if (!box) return;

    this.positionTextRunEditor(this.activeEditor, box);
    this.updateInlineCaret(this.activeEditor, this.activeEditorTarget);
  }

  private markDirty(): void {
    this.isDirty = true;
    this.editVersion++;
    this.setSaveState('dirty');
    if (this.getSettings().autosaveEnabled) {
      this.scheduleAutosave();
    }
  }

  private scheduleAutosave(): void {
    this.clearAutosave();
    if (!this.getSettings().autosaveEnabled) return;

    this.saveTimer = window.setTimeout(() => {
      this.saveTimer = null;
      void this.saveCurrentPresentation();
    }, 1500);
  }

  private clearAutosave(): void {
    if (this.saveTimer !== null) {
      window.clearTimeout(this.saveTimer);
      this.saveTimer = null;
    }
  }

  private async preserveUnsavedChangesForTeardown(reason: string): Promise<boolean> {
    this.clearAutosave();
    await this.savePromise.catch(() => undefined);

    if (!this.isDirty || !this.engine || !this.loadedFile) {
      return true;
    }

    if (this.getSettings().autosaveEnabled) {
      const saved = await this.saveCurrentPresentation();
      if (saved) return true;
    }

    return this.writeRecoveryCopy(reason);
  }

  private async writeRecoveryCopy(reason: string): Promise<boolean> {
    const file = this.loadedFile;
    const engine = this.engine;
    const sourcePackage = this.sourcePackage;
    const sourceBuffer = this.sourceBuffer;
    if (!file || !engine || !sourcePackage || !sourceBuffer) {
      new Notice('Could not create a Native PowerPoint recovery copy because the open presentation is unavailable.');
      return false;
    }

    try {
      const output = await engine.export();
      let isValidated = true;
      let validationError = '';

      try {
        await this.validateExportBeforeSave(output, engine, sourcePackage, sourceBuffer);
      } catch (error) {
        isValidated = false;
        validationError = cleanError(error);
      }

      const recoveryPath = this.getAvailableRecoveryPath(file, isValidated);
      await this.app.vault.createBinary(recoveryPath, output);
      this.isDirty = false;
      this.setSaveState('recovered');

      if (isValidated) {
        new Notice(`Unsaved edits were not written to ${file.name}. Recovery copy created while ${reason}: ${recoveryPath}`, 10000);
      } else {
        new Notice(
          `Save validation failed, so ${file.name} was not overwritten. An unvalidated recovery export was created at ${recoveryPath}. ${validationError}`,
          15000
        );
      }

      return true;
    } catch (error) {
      this.setSaveState('failed');
      new Notice(
        `Could not preserve unsaved edits from ${file.name} while ${reason}: ${cleanError(error)}. The original file was not overwritten.`,
        15000
      );
      return false;
    }
  }

  private getAvailableRecoveryPath(file: TFile, isValidated: boolean): string {
    const slashIndex = file.path.lastIndexOf('/');
    const parentPath = slashIndex === -1 ? '' : file.path.slice(0, slashIndex);
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const recoveryType = isValidated ? 'recovery' : 'unvalidated recovery';
    const baseName = `${file.basename} (Native PowerPoint ${recoveryType} ${timestamp})`;
    let sequence = 0;

    while (true) {
      const suffix = sequence === 0 ? '' : ` ${sequence + 1}`;
      const fileName = `${baseName}${suffix}.${file.extension}`;
      const candidate = normalizePath(parentPath ? `${parentPath}/${fileName}` : fileName);
      if (!this.app.vault.getAbstractFileByPath(candidate)) {
        return candidate;
      }
      sequence++;
    }
  }

  private resetLoadedPresentation(): void {
    this.clearAutosave();
    this.removeActiveEditor();
    this.clearHistory();
    this.engine = null;
    this.loadedFile = null;
    this.sourcePackage = null;
    this.sourceBuffer = null;
    this.isViewOnly = false;
    this.viewOnlyReason = '';
    this.isDirty = false;
    this.editVersion = 0;
    this.savePromise = Promise.resolve();
    this.selectedShapeIndex = null;
    this.selectedTransform = null;
    this.findMatches = [];
    this.currentFindMatchIndex = 0;
    this.hasShownGeneratedTextNotice = false;
    this.fontSubstitutions = [];
    if (this.findInputEl) this.findInputEl.value = '';
    this.closeFindPanel();
    this.updateFindStatus();
    this.svgEl = null;
    this.dragState = null;
  }

  private updateEditingAvailability(): void {
    const canEdit = this.canEdit();
    const disabledReason = this.viewOnlyReason || 'Open an editable PowerPoint file first.';

    for (const button of this.editButtons) {
      const baseTitle = button.dataset.baseTitle || button.getAttribute('aria-label') || 'Edit';
      button.disabled = !canEdit;
      button.toggleClass('is-disabled', !canEdit);
      button.setAttribute('title', canEdit ? baseTitle : `${baseTitle}: ${disabledReason}`);
      button.setAttribute('aria-disabled', String(!canEdit));
    }

    this.updateHistoryAvailability();
    this.updateObjectClipboardAvailability();
  }

  private updateObjectClipboardAvailability(): void {
    const hasSelection = this.selectedShapeIndex !== null;
    const canEdit = this.canEdit();
    this.updateObjectClipboardButton(this.copyButton, hasSelection);
    this.updateObjectClipboardButton(this.pasteButton, canEdit && Boolean(this.objectClipboard));
    this.updateObjectClipboardButton(this.duplicateButton, canEdit && hasSelection);
  }

  private updateObjectClipboardButton(button: HTMLButtonElement | null, enabled: boolean): void {
    if (!button) return;
    button.disabled = !enabled;
    button.toggleClass('is-disabled', !enabled);
    button.setAttribute('aria-disabled', String(!enabled));
  }

  private updateHistoryAvailability(): void {
    const canUseHistory = this.canEdit() && !this.isRestoringHistory;
    this.updateHistoryButton(this.undoButton, 'Undo', 'Ctrl+Z', canUseHistory && this.undoStack.length > 0);
    this.updateHistoryButton(this.redoButton, 'Redo', 'Ctrl+Shift+Z', canUseHistory && this.redoStack.length > 0);
  }

  private updateHistoryButton(
    button: HTMLButtonElement | null,
    label: string,
    shortcut: string,
    enabled: boolean
  ): void {
    if (!button) return;

    const nextEntry = label === 'Undo'
      ? this.undoStack[this.undoStack.length - 1]
      : this.redoStack[this.redoStack.length - 1];
    button.disabled = !enabled;
    button.toggleClass('is-disabled', !enabled);
    button.setAttribute('aria-disabled', String(!enabled));
    button.setAttribute('title', nextEntry ? `${label} ${nextEntry.label} (${shortcut})` : `${label} (${shortcut})`);
  }

  private setSaveState(state: SaveState): void {
    this.saveState = state;
    if (!this.statusEl) return;

    const labels: Record<SaveState, string> = {
      idle: 'Ready',
      dirty: 'Unsaved',
      saving: 'Saving...',
      saved: 'Saved',
      failed: 'Save failed',
      recovered: 'Recovery saved',
      'view-only': 'View-only'
    };

    this.statusEl.setText(labels[state]);
    this.statusEl.dataset.state = state;
  }

  private updateSlideCounter(): void {
    const count = this.engine?.slideCount || 0;
    this.slideCounterEl?.setText(count ? `${this.currentSlide + 1} / ${count}` : '0 / 0');

    if (this.thumbnailContainer) {
      this.thumbnailContainer.querySelectorAll('.native-powerpoint-thumbnail').forEach((thumbnail, index) => {
        thumbnail.toggleClass('active', index === this.currentSlide);
      });
    }
  }
}
