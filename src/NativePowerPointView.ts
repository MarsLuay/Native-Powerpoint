import { FileView, Menu, Notice, Platform, TFile, WorkspaceLeaf, normalizePath, setIcon } from 'obsidian';

import {
  PresentationEngine,
  type GeneratedTextEdit,
  type GeneratedTextKind,
  type ImageCrop,
  type InsertableShapeGeometry,
  type ParagraphAlignment,
  type RunStyleChange,
  type RunStyleInfo,
  type RunTarget,
  type ShapeReorderMode,
  type SlideLayoutKind
} from './PresentationEngine';
import {
  getImageMimeType,
  ImageCropModal,
  InsertTableModal,
  VaultImageSuggestModal,
  type ImageCropValues
} from './PowerPointInsertModals';
import type { ParagraphListStyle } from './SlideInsertions';
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
import { isElement, isNode, isSVGGElement, isSVGTextElement, isSVGTSpanElement } from './domGuards';
import { PowerPointPresentController } from './PowerPointPresent';
import { exportSlideToPng, exportSlidesToPdf, exportSlidesToPngZip } from './PowerPointExport';

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
type HandleName = 'nw' | 'ne' | 'sw' | 'se' | 'n' | 'e' | 's' | 'w';
type SvgSecurityDecision = 'compatibility' | 'yolo' | null;

interface PointerPoint {
  x: number;
  y: number;
}

interface DragState {
  mode: 'move' | 'resize' | 'rotate';
  handle?: HandleName;
  pointerId: number;
  startPoint: PointerPoint;
  startClientX: number;
  startClientY: number;
  startBox: { left: number; top: number; width: number; height: number };
  startTransform: ShapeTransform;
  latestTransform: ShapeTransform;
  centerClientX?: number;
  centerClientY?: number;
  startAngle?: number;
}

interface MarqueeState {
  pointerId: number;
  startClientX: number;
  startClientY: number;
  additive: boolean;
  base: number[];
  moved: boolean;
}

interface GroupDragState {
  pointerId: number;
  startPoint: PointerPoint;
  startClientX: number;
  startClientY: number;
  start: Map<number, ShapeTransform>;
  latest: Map<number, ShapeTransform>;
  moved: boolean;
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

interface SvgRectLike {
  x: number;
  y: number;
  width: number;
  height: number;
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
  pendingFrame: number | null;
  pendingClientX: number;
  pendingClientY: number;
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
  kind: 'shape-paragraph';
  shapeIndex: number;
  paragraphIndex: number;
  runIndex: number;
  text: string;
  element: SVGTextElement | SVGTSpanElement;
  runElements: SVGTSpanElement[];
}

interface GeneratedTextEditTarget extends GeneratedTextEdit {
  shapeIndex: number;
  text: string;
  element: SVGTextElement;
}

type TextEditTarget = GeneratedTextEditTarget | ShapeTextEditTarget;

interface TextToolbarControls {
  bold: HTMLButtonElement;
  italic: HTMLButtonElement;
  underline: HTMLButtonElement;
  fontLabel: HTMLElement;
  fontSizeInput: HTMLInputElement;
  textColorBar: HTMLElement;
  highlightBar: HTMLElement;
  alignButtons: Record<ParagraphAlignment, HTMLButtonElement>;
}

interface TextStyleContext {
  shapeIndex: number;
  run: RunTarget | null;
  anchor: { left: number; top: number; width: number; height: number };
}

const TEXT_TOOLBAR_FONTS = [
  'Arial',
  'Calibri',
  'Cambria',
  'Comic Sans MS',
  'Courier New',
  'Georgia',
  'Garamond',
  'Impact',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana'
];

const TEXT_TOOLBAR_SWATCHES = [
  '000000', '434343', '666666', '999999', 'B7B7B7', 'CCCCCC', 'D9D9D9', 'FFFFFF',
  '980000', 'FF0000', 'FF9900', 'FFFF00', '00FF00', '00FFFF', '4A86E8', '0000FF',
  '9900FF', 'FF00FF', 'E6B8AF', 'FCE5CD', 'FFF2CC', 'D9EAD3', 'D0E0E3', 'C9DAF8'
];

const TEXT_TOOLBAR_MIN_FONT_SIZE = 1;
const TEXT_TOOLBAR_MAX_FONT_SIZE = 400;

const GENERATED_GRID_SELECTOR =
  'g[data-ooxml-shape-type="table"], g[data-ooxml-shape-type="chart"]';
const HISTORY_LIMIT = 20;
const MIN_ZOOM = 0.25;
const MAX_ZOOM = 3;
const SNAP_THRESHOLD_PX = 6;

type AlignMode = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom';
type DistributeAxis = 'horizontal' | 'vertical';

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

/** First face from a CSS `font-family` value, e.g. `"Calibri", sans-serif` → Calibri. */
function parsePrimaryFontFamily(fontFamily: string): string | null {
  const trimmed = fontFamily.trim();
  if (!trimmed) return null;

  if (trimmed.startsWith('"')) {
    const end = trimmed.indexOf('"', 1);
    if (end > 1) return trimmed.slice(1, end);
  } else if (trimmed.startsWith("'")) {
    const end = trimmed.indexOf("'", 1);
    if (end > 1) return trimmed.slice(1, end);
  }

  const comma = trimmed.indexOf(',');
  const primary = (comma >= 0 ? trimmed.slice(0, comma) : trimmed).trim();
  if (!primary || primary === 'inherit' || primary === 'initial' || primary === 'unset') return null;

  const generic = new Set([
    'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-serif', 'ui-sans-serif',
    'ui-monospace', 'ui-rounded', 'emoji', 'math', 'fangsong'
  ]);
  return generic.has(primary.toLowerCase()) ? null : primary;
}

function getToolbarTooltipText(target: HTMLElement): string {
  return (target.getAttribute('aria-label') || '').replace(/\s+/g, ' ').trim();
}

function positionToolbarTooltip(target: HTMLElement, tooltip: HTMLDivElement): void {
  const rect = target.getBoundingClientRect();
  tooltip.style.setProperty('--docxidian-toolbar-tooltip-left', `${Math.round(rect.left + rect.width / 2)}px`);
  tooltip.style.setProperty('--docxidian-toolbar-tooltip-top', `${Math.round(rect.bottom + 8)}px`);
  tooltip.removeClasses(['is-left-aligned', 'is-right-aligned']);

  const tooltipRect = tooltip.getBoundingClientRect();
  const viewportPadding = 8;
  if (tooltipRect.left < viewportPadding) {
    tooltip.style.setProperty('--docxidian-toolbar-tooltip-left', `${viewportPadding}px`);
    tooltip.addClass('is-left-aligned');
  } else if (tooltipRect.right > window.innerWidth - viewportPadding) {
    tooltip.style.setProperty('--docxidian-toolbar-tooltip-left', `${window.innerWidth - viewportPadding}px`);
    tooltip.addClass('is-right-aligned');
  }
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

type MenuDropdownEntry =
  | 'separator'
  | { label: string; icon?: string; onClick: () => void; disabled?: boolean };

export class NativePowerPointView extends FileView {
  private readonly getSettings: () => NativePowerPointSettings;

  private engine: PresentationEngine | null = null;
  private loadedFile: TFile | null = null;
  private sourcePackage: PowerPointPackageInspection | null = null;
  private sourceBuffer: ArrayBuffer | null = null;
  private currentSlide = 0;
  private zoomLevel = 1;
  private selectedShapeIndex: number | null = null;
  private selectedShapeIndices = new Set<number>();
  private selectedSlideIndices = new Set<number>();
  private lastInteractionRegion: 'canvas' | 'thumbnails' = 'canvas';
  private selectedTransform: ShapeTransform | null = null;
  private marquee: MarqueeState | null = null;
  private marqueeEl: HTMLElement | null = null;
  private groupDrag: GroupDragState | null = null;
  private multiSelectionBoxes: HTMLElement[] = [];
  private snapGuides: HTMLElement[] = [];
  private suppressNextClick = false;
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
  private inlineWholeShapeSelection: string | null = null;
  private inlineWholeShapeSelected = false;
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
  private findHighlightRects: SVGRectElement[] = [];

  private layoutEl: HTMLElement | null = null;
  private headerTitleEl: HTMLElement | null = null;
  private zoomLevelEl: HTMLElement | null = null;
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
  private findReplaceInputEl: HTMLInputElement | null = null;
  private findReplaceToggleEl: HTMLButtonElement | null = null;
  private findButtonEl: HTMLButtonElement | null = null;
  private findPanelDismissHandler: ((event: Event) => void) | null = null;
  private findPanelRepositionHandler: (() => void) | null = null;
  private isFindReplaceMode = false;
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
  private alignButtons: HTMLButtonElement[] = [];
  private distributeButtons: HTMLButtonElement[] = [];
  private zOrderButtons: HTMLButtonElement[] = [];
  private groupButton: HTMLButtonElement | null = null;
  private ungroupButton: HTMLButtonElement | null = null;
  private imageFileInput: HTMLInputElement | null = null;
  private replaceImageFileInput: HTMLInputElement | null = null;
  private insertTableButton: HTMLButtonElement | null = null;
  private pendingReplaceShapeIndex: number | null = null;
  private activeInsertMenu: HTMLElement | null = null;
  private activeMenuBarTab: HTMLElement | null = null;
  private activeMenuBarDropdown: HTMLElement | null = null;
  private menuBarCloseTimer: number | null = null;
  private activeShapeTextTarget: ShapeTextEditTarget | null = null;
  private activeTextStyleTarget: ShapeTextEditTarget | null = null;
  private textToolbarEl: HTMLElement | null = null;
  private textToolbarControls: TextToolbarControls | null = null;
  private textToolbarShapeIndex: number | null = null;
  private currentRunStyle: RunStyleInfo | null = null;
  private textColorValue = '000000';
  private textHighlightValue = 'FFFF00';
  private activeToolbarPopover: HTMLElement | null = null;
  private toolbarPopoverCleanup: (() => void) | null = null;
  private presentController: PowerPointPresentController | null = null;
  private thumbnailDragIndex: number | null = null;

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
    this.detachFindPanelDismissHandlers();
    this.findPanelEl?.remove();
    this.findPanelEl = null;
    this.presentController?.dispose();
    this.presentController = null;

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
    this.createHeaderBar(root);
    this.layoutEl = root.createDiv({ cls: 'native-powerpoint-layout' });

    const sidebar = this.layoutEl.createDiv({ cls: 'native-powerpoint-sidebar' });
    this.registerDomEvent(sidebar, 'pointerdown', () => {
      this.lastInteractionRegion = 'thumbnails';
    }, true);
    const sidebarHeader = sidebar.createDiv({ cls: 'native-powerpoint-sidebar-header', text: 'Slides' });
    const addSlideButton = sidebarHeader.createEl('button', {
      cls: 'native-powerpoint-sidebar-add',
      attr: { 'aria-label': 'New slide' }
    });
    setIcon(addSlideButton, 'plus');
    addSlideButton.addEventListener('click', () => void this.addSlideWithLayout('blank'));
    this.thumbnailContainer = sidebar.createDiv({ cls: 'native-powerpoint-thumbnails' });

    const main = this.layoutEl.createDiv({ cls: 'native-powerpoint-main-content' });
    this.createToolbar(main);
    this.canvasPane = main.createDiv({ cls: 'native-powerpoint-canvas-pane' });
    this.slideSurface = this.canvasPane.createDiv({ cls: 'native-powerpoint-slide-surface' });
    this.registerDomEvent(this.canvasPane, 'pointerdown', this.handleCanvasPanePointerDown, true);
    this.registerDomEvent(this.canvasPane, 'contextmenu', this.handleCanvasContextMenu);
    this.registerCanvasWheelZoom();
    this.observeCanvasPane();
    this.registerInsertMenus();
    this.textToolbarEl = null;
    this.textToolbarControls = null;
    this.register(() => this.closeToolbarPopover());

    this.inspectorEl = this.layoutEl.createDiv({ cls: 'native-powerpoint-inspector' });
    this.applyInspectorVisibility();
    this.registerToolbarTooltips(root);
  }

  private applyInspectorVisibility(): void {
    const show = this.getSettings().showInspector;
    this.inspectorEl?.toggleClass('native-powerpoint-inspector-hidden', !show);
  }

  refreshSettings(): void {
    this.applyInspectorVisibility();
  }

  private registerToolbarTooltips(root: HTMLElement): void {
    const TOOLBAR_TOOLTIP_DELAY_MS = 450;
    let activeTarget: HTMLElement | null = null;
    let tooltipEl: HTMLDivElement | null = null;
    let tooltipTimer: number | null = null;

    const clearTooltipTimer = (): void => {
      if (tooltipTimer !== null) {
        window.clearTimeout(tooltipTimer);
        tooltipTimer = null;
      }
    };

    const removeTooltip = (): void => {
      tooltipEl?.remove();
      tooltipEl = null;
    };

    const hideTooltip = (): void => {
      clearTooltipTimer();
      removeTooltip();
      activeTarget = null;
    };

    const getTooltipTarget = (target: EventTarget | null): HTMLElement | null => {
      if (!isNode(target) || !isElement(target)) return null;
      const candidate = target.closest<HTMLElement>(
        '.native-powerpoint-toolbar button, .native-powerpoint-text-toolbar button, .native-powerpoint-find-panel button, .native-powerpoint-rotate-handle'
      );
      if (!candidate || !root.contains(candidate)) return null;
      return candidate;
    };

    const showTooltip = (target: HTMLElement): void => {
      const label = getToolbarTooltipText(target);
      if (!label || !target.isConnected) return;

      removeTooltip();
      const tooltip = activeDocument.body.createDiv({ cls: 'docxidian-toolbar-tooltip', text: label });
      tooltipEl = tooltip;
      positionToolbarTooltip(target, tooltip);
    };

    const scheduleTooltip = (target: HTMLElement): void => {
      if (target === activeTarget) return;
      hideTooltip();
      activeTarget = target;
      tooltipTimer = window.setTimeout(() => {
        tooltipTimer = null;
        if (activeTarget === target) {
          showTooltip(target);
        }
      }, TOOLBAR_TOOLTIP_DELAY_MS);
    };

    const handlePointerOver = (event: PointerEvent): void => {
      const target = getTooltipTarget(event.target);
      if (target) {
        scheduleTooltip(target);
      }
    };

    const handlePointerOut = (event: PointerEvent): void => {
      if (!activeTarget || (isNode(event.relatedTarget) && activeTarget.contains(event.relatedTarget))) {
        return;
      }
      hideTooltip();
    };

    this.registerDomEvent(root, 'pointerover', handlePointerOver);
    this.registerDomEvent(root, 'pointerout', handlePointerOut);
    this.registerDomEvent(window, 'scroll', hideTooltip, true);
    this.registerDomEvent(window, 'resize', hideTooltip);
    this.register(hideTooltip);
  }

  private createHeaderBar(root: HTMLElement): void {
    const headerBar = root.createDiv({ cls: 'native-powerpoint-headerbar' });

    const saveButton = headerBar.createEl('button', {
      cls: 'clickable-icon native-powerpoint-header-save',
      attr: { type: 'button', 'aria-label': 'Save' }
    });
    setIcon(saveButton, 'save');
    saveButton.addEventListener('click', () => void this.saveCurrentPresentation());

    const headerMain = headerBar.createDiv({ cls: 'native-powerpoint-headerbar-main' });
    this.createHeader(headerMain);
    this.createMenuBar(headerMain);
  }

  private createHeader(root: HTMLElement): void {
    const header = root.createDiv({ cls: 'native-powerpoint-header' });

    const title = header.createDiv({ cls: 'native-powerpoint-header-title' });
    this.headerTitleEl = title.createSpan({
      cls: 'native-powerpoint-header-name',
      text: this.getDisplayText()
    });
    this.headerTitleEl.setAttribute('role', 'button');
    this.headerTitleEl.setAttribute('tabindex', '0');
    this.headerTitleEl.setAttribute('aria-label', 'Rename presentation');
    this.headerTitleEl.title = 'Click to rename';
    this.headerTitleEl.addEventListener('click', () => this.beginRenameTitle());
    this.headerTitleEl.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        this.beginRenameTitle();
      }
    });

    this.statusEl = header.createDiv({ cls: 'native-powerpoint-save-status', text: 'Ready' });
  }

  private updateHeaderTitle(): void {
    this.headerTitleEl?.setText(this.getDisplayText());
  }

  private beginRenameTitle(): void {
    const file = this.loadedFile || this.file;
    const titleEl = this.headerTitleEl;
    const parent = titleEl?.parentElement ?? null;
    if (!file || !titleEl || !parent) return;
    if (parent.querySelector('.native-powerpoint-header-name-input')) return;

    const input = parent.createEl('input', {
      cls: 'native-powerpoint-header-name-input',
      type: 'text',
      value: file.basename
    });
    parent.insertBefore(input, titleEl);
    titleEl.hide();
    input.focus();
    input.select();

    let finished = false;
    const cleanup = () => {
      input.remove();
      titleEl.show();
    };
    const commit = () => {
      if (finished) return;
      finished = true;
      const newName = input.value;
      cleanup();
      void this.renameLoadedFile(file, newName);
    };
    const cancel = () => {
      if (finished) return;
      finished = true;
      cleanup();
    };

    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        commit();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        cancel();
      }
    });
    input.addEventListener('blur', () => commit());
  }

  private async renameLoadedFile(file: TFile, rawName: string): Promise<void> {
    const sanitized = rawName.replace(/[\\/:*?"<>|]/g, '').trim();
    if (!sanitized || sanitized === file.basename) return;

    const folder = file.parent?.path ?? '';
    const dir = folder && folder !== '/' ? `${folder}/` : '';
    const newPath = `${dir}${sanitized}.${file.extension}`;
    if (newPath === file.path) return;

    try {
      await this.app.fileManager.renameFile(file, newPath);
    } catch (error) {
      new Notice(`Could not rename file: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      this.updateHeaderTitle();
    }
  }

  private createMenuBar(root: HTMLElement): void {
    const bar = root.createDiv({ cls: 'native-powerpoint-menubar' });
    this.createDropdownTab(bar, 'File', () => this.getFileMenuItems());
    this.createDropdownTab(bar, 'Edit', () => this.getEditMenuItems());
    this.createDropdownTab(bar, 'Insert', () => this.getInsertMenuItems());
    this.createActionTab(bar, 'Search', () => this.openFindPanel());
    this.createActionTab(bar, 'Settings', () => this.openPluginSettings());

    this.registerDomEvent(
      activeDocument,
      'pointerdown',
      (event) => {
        const target = isNode(event.target) ? event.target : null;
        if (target && bar.contains(target)) return;
        if (target && this.activeMenuBarDropdown?.contains(target)) return;
        this.closeMenuBarDropdown();
      },
      true
    );
    this.register(() => this.closeMenuBarDropdown());
  }

  private createDropdownTab(
    bar: HTMLElement,
    label: string,
    getItems: () => MenuDropdownEntry[]
  ): HTMLButtonElement {
    const button = bar.createEl('button', {
      cls: 'native-powerpoint-menubar-item',
      text: label
    });
    button.type = 'button';
    button.addEventListener('click', () => {
      if (this.activeMenuBarTab === button) {
        this.closeMenuBarDropdown();
      } else {
        this.openMenuBarDropdown(button, getItems);
      }
    });
    // Google-style menu bar: hovering a tab reveals its options.
    button.addEventListener('mouseenter', () => this.openMenuBarDropdown(button, getItems));
    button.addEventListener('mouseleave', () => this.scheduleMenuBarClose());
    return button;
  }

  private createActionTab(bar: HTMLElement, label: string, action: () => void): HTMLButtonElement {
    const button = bar.createEl('button', {
      cls: 'native-powerpoint-menubar-item',
      text: label
    });
    button.type = 'button';
    button.addEventListener('click', () => {
      this.closeMenuBarDropdown();
      action();
    });
    // Moving onto a no-dropdown tab dismisses any open menu, like Google's bar.
    button.addEventListener('mouseenter', () => this.closeMenuBarDropdown());
    return button;
  }

  private openMenuBarDropdown(tab: HTMLElement, getItems: () => MenuDropdownEntry[]): void {
    this.cancelMenuBarCloseTimer();
    if (this.activeMenuBarTab === tab && this.activeMenuBarDropdown) return;
    this.closeMenuBarDropdown();

    const dropdown = activeDocument.body.createDiv({
      cls: 'native-powerpoint-menubar-dropdown native-powerpoint-light-surface'
    });
    for (const entry of getItems()) {
      if (entry === 'separator') {
        dropdown.createDiv({ cls: 'native-powerpoint-menubar-dropdown-sep' });
        continue;
      }
      const item = dropdown.createEl('button', { cls: 'native-powerpoint-menubar-dropdown-item' });
      item.type = 'button';
      if (entry.icon) {
        setIcon(item.createSpan({ cls: 'native-powerpoint-menubar-dropdown-icon' }), entry.icon);
      }
      item.createSpan({ cls: 'native-powerpoint-menubar-dropdown-label', text: entry.label });
      if (entry.disabled) {
        item.disabled = true;
      } else {
        item.addEventListener('click', () => {
          this.closeMenuBarDropdown();
          entry.onClick();
        });
      }
    }

    dropdown.addEventListener('mouseenter', () => this.cancelMenuBarCloseTimer());
    dropdown.addEventListener('mouseleave', () => this.scheduleMenuBarClose());

    const rect = tab.getBoundingClientRect();
    dropdown.style.left = `${rect.left}px`;
    dropdown.style.top = `${rect.bottom + 2}px`;

    tab.addClass('is-active');
    this.activeMenuBarTab = tab;
    this.activeMenuBarDropdown = dropdown;
  }

  private scheduleMenuBarClose(): void {
    this.cancelMenuBarCloseTimer();
    this.menuBarCloseTimer = window.setTimeout(() => {
      this.menuBarCloseTimer = null;
      this.closeMenuBarDropdown();
    }, 160);
  }

  private cancelMenuBarCloseTimer(): void {
    if (this.menuBarCloseTimer !== null) {
      window.clearTimeout(this.menuBarCloseTimer);
      this.menuBarCloseTimer = null;
    }
  }

  private closeMenuBarDropdown(): void {
    this.cancelMenuBarCloseTimer();
    this.activeMenuBarDropdown?.remove();
    this.activeMenuBarDropdown = null;
    this.activeMenuBarTab?.removeClass('is-active');
    this.activeMenuBarTab = null;
  }

  private showMenuUnder(menu: Menu, anchor: HTMLElement): void {
    const rect = anchor.getBoundingClientRect();
    menu.showAtPosition({ x: rect.left, y: rect.bottom + 4 });
  }

  private createNativeMenu(): Menu {
    const instance = new Menu();
    const dom = (instance as unknown as { dom?: HTMLElement }).dom;
    dom?.addClass('native-powerpoint-light-surface');
    return instance;
  }

  private getFileMenuItems(): MenuDropdownEntry[] {
    return [
      { label: 'Save', icon: 'save', onClick: () => void this.saveCurrentPresentation() },
      { label: 'Duplicate', icon: 'copy', onClick: () => void this.duplicatePresentation() },
      'separator',
      { label: 'Print', icon: 'printer', onClick: () => void this.printPresentation() },
      'separator',
      {
        label: 'Export current slide as PNG',
        icon: 'image',
        onClick: () => void this.exportCurrentSlideAsPng()
      },
      {
        label: 'Export current slide as PDF',
        icon: 'file-output',
        onClick: () => void this.exportDeckAsPdf(true)
      },
      {
        label: 'Export deck as PDF',
        icon: 'file-output',
        onClick: () => void this.exportDeckAsPdf(false)
      },
      {
        label: 'Export deck as PNGs (zip)',
        icon: 'file-archive',
        onClick: () => void this.exportDeckAsPngZip()
      },
      'separator',
      {
        label: 'Present from current slide',
        icon: 'play',
        onClick: () => this.startPresentation()
      }
    ];
  }

  private getEditMenuItems(): MenuDropdownEntry[] {
    const canEdit = this.canEdit();
    const canUseHistory = canEdit && !this.isRestoringHistory;
    const hasSelection = this.selectedShapeIndex !== null || this.selectedShapeIndices.size > 0;
    const hasClipboard = Boolean(this.objectClipboard);

    return [
      {
        label: 'Undo',
        icon: 'undo',
        onClick: () => void this.undo(),
        disabled: !canUseHistory || this.undoStack.length === 0
      },
      {
        label: 'Redo',
        icon: 'redo',
        onClick: () => void this.redo(),
        disabled: !canUseHistory || this.redoStack.length === 0
      },
      'separator',
      {
        label: 'Cut',
        icon: 'scissors',
        onClick: () => void this.cutSelectedShape(),
        disabled: !canEdit || this.selectedShapeIndex === null
      },
      {
        label: 'Copy',
        icon: 'copy',
        onClick: () => void this.copySelectedShape(),
        disabled: this.selectedShapeIndex === null
      },
      {
        label: 'Paste',
        icon: 'clipboard-paste',
        onClick: () => void this.pasteCopiedShape(),
        disabled: !canEdit || !hasClipboard
      },
      {
        label: 'Paste without formatting',
        icon: 'clipboard-type',
        onClick: () => void this.pasteWithoutFormatting(),
        disabled: !canEdit || (!this.activeEditor && !hasClipboard)
      },
      {
        label: 'Delete',
        icon: 'trash-2',
        onClick: () => void this.deleteSelectedShape(),
        disabled: !canEdit || !hasSelection
      },
      'separator',
      {
        label: 'Select all',
        icon: 'box-select',
        onClick: () => this.selectAllShapes()
      },
      'separator',
      {
        label: 'Find and replace',
        icon: 'replace',
        onClick: () => this.openFindPanel({ replace: true })
      }
    ];
  }

  private getInsertMenuItems(): MenuDropdownEntry[] {
    return [
      { label: 'Image from vault', icon: 'image', onClick: () => this.openVaultImagePicker() },
      { label: 'Upload image', icon: 'upload', onClick: () => this.imageFileInput?.click() },
      'separator',
      { label: 'Text box', icon: 'type', onClick: () => void this.addTextBox() },
      { label: 'Rectangle', icon: 'square', onClick: () => void this.insertShape('rect') },
      { label: 'Ellipse', icon: 'circle', onClick: () => void this.insertShape('ellipse') },
      { label: 'Line', icon: 'minus', onClick: () => void this.insertShape('line') },
      { label: 'Arrow', icon: 'move-right', onClick: () => void this.insertShape('rightArrow') },
      'separator',
      { label: 'Table', icon: 'table', onClick: () => this.openTableSizePicker(this.insertTableButton) },
      { label: 'Chart', icon: 'bar-chart-3', onClick: () => void this.insertChart() },
      'separator',
      { label: 'Bulleted list', icon: 'list', onClick: () => void this.applyListStyle('bullet') },
      {
        label: 'Numbered list',
        icon: 'list-ordered',
        onClick: () => void this.applyListStyle('number')
      },
      'separator',
      { label: 'New slide', icon: 'plus', onClick: () => void this.addSlideWithLayout('blank') }
    ];
  }

  private async duplicatePresentation(): Promise<void> {
    const file = this.loadedFile || this.file;
    if (!file) {
      new Notice('Open a presentation to duplicate it.');
      return;
    }

    try {
      if (this.isDirty && this.canEdit()) {
        await this.saveCurrentPresentation();
      }
      const copyPath = this.getAvailableCopyPath(file);
      const data = await this.app.vault.readBinary(file);
      const created = await this.app.vault.createBinary(copyPath, data);
      new Notice(`Duplicated to ${created.name}`);
    } catch (error) {
      new Notice(`Could not duplicate presentation: ${cleanError(error)}`);
    }
  }

  private getAvailableCopyPath(file: TFile): string {
    const folderPath = file.parent?.path;
    const folderPrefix = folderPath && folderPath !== '/' ? `${folderPath}/` : '';
    for (let index = 1; index < 1000; index += 1) {
      const suffix = index === 1 ? 'copy' : `copy ${index}`;
      const candidate = normalizePath(`${folderPrefix}${file.basename} ${suffix}.${file.extension}`);
      if (!this.app.vault.getAbstractFileByPath(candidate)) return candidate;
    }
    return normalizePath(`${folderPrefix}${file.basename} copy ${Date.now()}.${file.extension}`);
  }

  private async printPresentation(): Promise<void> {
    if (!this.engine || this.engine.slideCount === 0) {
      new Notice('Open a presentation with at least one slide to print.');
      return;
    }

    try {
      const indices = Array.from({ length: this.engine.slideCount }, (_, index) => index);
      const elements = this.collectExportSvgElements(indices);
      if (elements.length === 0) {
        throw new Error('No slides could be rendered for printing.');
      }

      new Notice('Preparing slides for printing...');
      const urls: string[] = [];
      for (const element of elements) {
        const bytes = await exportSlideToPng(element, this.contentEl.ownerDocument);
        urls.push(URL.createObjectURL(new Blob([bytes], { type: 'image/png' })));
      }
      this.printSlideImages(urls);
    } catch (error) {
      new Notice(`Could not print: ${cleanError(error)}`);
    }
  }

  private printSlideImages(urls: string[]): void {
    const iframe = activeDocument.body.createEl('iframe', { cls: 'native-powerpoint-print-frame' });
    let cleaned = false;
    const cleanup = () => {
      if (cleaned) return;
      cleaned = true;
      urls.forEach((url) => URL.revokeObjectURL(url));
      iframe.remove();
    };

    const doc = iframe.contentDocument;
    const win = iframe.contentWindow;
    if (!doc || !win) {
      cleanup();
      new Notice('Could not open the print view.');
      return;
    }

    const style = doc.createElement('style');
    style.textContent =
      '@page { size: landscape; margin: 12mm; }' +
      'html, body { margin: 0; padding: 0; background: #ffffff; }' +
      '.native-powerpoint-print-slide { page-break-after: always; text-align: center; }' +
      '.native-powerpoint-print-slide:last-child { page-break-after: auto; }' +
      '.native-powerpoint-print-slide img { width: 100%; height: auto; display: block; }';
    doc.head.appendChild(style);

    let remaining = urls.length;
    const onReady = () => {
      remaining -= 1;
      if (remaining > 0) return;
      win.focus();
      win.print();
    };

    win.addEventListener('afterprint', cleanup, { once: true });
    window.setTimeout(cleanup, 60000);

    for (const url of urls) {
      const wrap = doc.createElement('div');
      wrap.className = 'native-powerpoint-print-slide';
      const img = doc.createElement('img');
      img.addEventListener('load', onReady, { once: true });
      img.addEventListener('error', onReady, { once: true });
      img.src = url;
      wrap.appendChild(img);
      doc.body.appendChild(wrap);
    }
  }

  private openPluginSettings(): void {
    const setting = (
      this.app as unknown as {
        setting?: { open?: () => void; openTabById?: (id: string) => void };
      }
    ).setting;
    if (!setting?.open || !setting.openTabById) {
      new Notice('Unable to open settings from here. Open Obsidian settings manually.');
      return;
    }
    setting.open();
    setting.openTabById('native-powerpoint-doc-editor');
  }

  private updateZoomLabel(): void {
    this.zoomLevelEl?.setText(`${Math.round(this.zoomLevel * 100)}%`);
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

    // Layout mirrors Google Slides' toolbar order: history first, then zoom,
    // then slide operations, then insert/object operations, then find. Slide
    // navigation sits on the right like Slides' top-right controls.
    const historyGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.undoButton = this.createIconButton(historyGroup, 'undo', 'Undo (Ctrl+Z)', () => void this.undo());
    this.redoButton = this.createIconButton(historyGroup, 'redo', 'Redo (Ctrl+Y)', () => void this.redo());

    const zoomGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.createIconButton(zoomGroup, 'zoom-out', 'Zoom out', () => this.setZoom(this.zoomLevel - 0.1));
    this.zoomLevelEl = zoomGroup.createDiv({ cls: 'native-powerpoint-zoom-level', text: '100%' });
    this.createIconButton(zoomGroup, 'zoom-in', 'Zoom in', () => this.setZoom(this.zoomLevel + 0.1));
    this.updateZoomLabel();

    this.createInsertToolbarGroup(toolbar);

    const slideGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    // Primary click adds a blank slide immediately (Google Slides "+" behavior).
    this.createEditIconButton(slideGroup, 'plus', 'New slide', () => void this.addSlideWithLayout('blank'));
    // A caret opens the layout choices without blocking the quick-add action.
    const newSlideLayoutButton = this.createEditIconButton(slideGroup, 'chevron-down', 'New slide layout', () => {
      this.toggleInsertMenu(newSlideLayoutButton, [
        { label: 'Blank', onClick: () => void this.addSlideWithLayout('blank') },
        { label: 'Title', onClick: () => void this.addSlideWithLayout('title') },
        { label: 'Title + Body', onClick: () => void this.addSlideWithLayout('titleBody') }
      ]);
    });
    this.createEditIconButton(slideGroup, 'files', 'Duplicate slide', () => void this.duplicateSlide());
    this.createEditIconButton(slideGroup, 'trash-2', 'Delete slide', () => void this.deleteSlide());
    this.createEditIconButton(slideGroup, 'arrow-left-to-line', 'Move slide left', () => void this.moveSlide(-1));
    this.createEditIconButton(slideGroup, 'arrow-right-to-line', 'Move slide right', () => void this.moveSlide(1));

    const objectGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.copyButton = this.createIconButton(objectGroup, 'copy', 'Copy selected object (Ctrl+C)', () => void this.copySelectedShape());
    this.pasteButton = this.createIconButton(objectGroup, 'clipboard-paste', 'Paste object (Ctrl+V)', () => void this.pasteCopiedShape());
    this.duplicateButton = this.createIconButton(objectGroup, 'copy-plus', 'Duplicate selected object (Ctrl+D)', () => void this.duplicateSelectedShape());
    this.createEditIconButton(objectGroup, 'eraser', 'Delete selected object', () => void this.deleteSelectedShape());

    this.createArrangeToolbarGroups(toolbar);

    const searchGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.findButtonEl = this.createIconButton(searchGroup, 'search', 'Find in presentation (Ctrl+F)', () => this.toggleFindPanel());
    // The find/replace UI is a floating dropdown anchored to the search button
    // rather than an inline element inside the horizontally scrolling toolbar.
    this.createFindPanel();

    const shareGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.createIconButton(shareGroup, 'play', 'Present from current slide', () => this.startPresentation());
    const exportButton = this.createIconButton(shareGroup, 'download', 'Export slides', () =>
      this.openExportMenu(exportButton)
    );

    const navGroup = toolbar.createDiv({
      cls: 'native-powerpoint-toolbar-group native-powerpoint-toolbar-group-end'
    });
    this.createIconButton(navGroup, 'chevron-left', 'Previous slide', () => void this.goToSlide(this.currentSlide - 1));
    this.slideCounterEl = navGroup.createDiv({ cls: 'native-powerpoint-page-counter', text: '0 / 0' });
    this.createIconButton(navGroup, 'chevron-right', 'Next slide', () => void this.goToSlide(this.currentSlide + 1));

    this.updateEditingAvailability();
    this.updateHistoryAvailability();
    this.updateObjectClipboardAvailability();
  }

  private registerInsertMenus(): void {
    const closeMenus = (event: MouseEvent) => {
      const target = isNode(event.target) ? event.target : null;
      if (target && this.activeInsertMenu?.contains(target)) return;
      if (target instanceof Element && target.closest('.native-powerpoint-insert-menu-anchor')) return;
      this.closeInsertMenus();
    };
    this.registerDomEvent(activeDocument, 'pointerdown', closeMenus, true);

    if (!this.layoutEl) return;
    const input = this.layoutEl.createEl('input', {
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

    const replaceInput = this.layoutEl.createEl('input', {
      type: 'file',
      cls: 'native-powerpoint-image-file-input'
    });
    replaceInput.accept = 'image/png,image/jpeg,image/gif,image/webp,image/bmp';
    replaceInput.addEventListener('change', () => {
      const file = replaceInput.files?.[0];
      replaceInput.value = '';
      const shapeIndex = this.pendingReplaceShapeIndex;
      this.pendingReplaceShapeIndex = null;
      if (file && shapeIndex !== null) void this.replaceImageWithLocalFile(shapeIndex, file);
    });
    this.replaceImageFileInput = replaceInput;
  }

  private createInsertToolbarGroup(toolbar: HTMLElement): void {
    const insertGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });

    const imageButton = this.createEditIconButton(insertGroup, 'image', 'Insert image', () => {
      this.toggleInsertMenu(imageButton, [
        { label: 'From vault', onClick: () => this.openVaultImagePicker() },
        { label: 'Upload file', onClick: () => this.imageFileInput?.click() }
      ]);
    });

    const shapeButton = this.createEditIconButton(insertGroup, 'shapes', 'Insert shape', () => {
      this.toggleInsertMenu(shapeButton, [
        { label: 'Rectangle', onClick: () => void this.insertShape('rect') },
        { label: 'Ellipse', onClick: () => void this.insertShape('ellipse') },
        { label: 'Rounded rectangle', onClick: () => void this.insertShape('roundRect') },
        { label: 'Line', onClick: () => void this.insertShape('line') },
        { label: 'Arrow', onClick: () => void this.insertShape('rightArrow') }
      ]);
    });

    this.createEditIconButton(insertGroup, 'type', 'Insert text box', () => void this.addTextBox());
    const tableButton = this.createEditIconButton(insertGroup, 'table', 'Insert table', () =>
      this.openTableSizePicker(tableButton)
    );
    this.insertTableButton = tableButton;
    this.createEditIconButton(insertGroup, 'bar-chart-3', 'Insert chart', () => void this.insertChart());
    this.createEditIconButton(insertGroup, 'list', 'Bulleted list', () => void this.applyListStyle('bullet'));
    this.createEditIconButton(insertGroup, 'list-ordered', 'Numbered list', () => void this.applyListStyle('number'));
  }

  private createArrangeToolbarGroups(toolbar: HTMLElement): void {
    this.alignButtons = [];
    this.distributeButtons = [];
    this.zOrderButtons = [];

    // Object-align buttons were removed from the toolbar to declutter it; object
    // alignment remains available via the right-click "Center on page" menu.
    const alignGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.distributeButtons.push(
      this.createIconButton(alignGroup, 'align-horizontal-distribute-center', 'Distribute horizontally', () => void this.distributeSelectedShapes('horizontal')),
      this.createIconButton(alignGroup, 'align-vertical-distribute-center', 'Distribute vertically', () => void this.distributeSelectedShapes('vertical'))
    );

    const orderGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.zOrderButtons.push(
      this.createIconButton(orderGroup, 'bring-to-front', 'Bring to front', () => void this.reorderSelection('front')),
      this.createIconButton(orderGroup, 'arrow-up', 'Bring forward', () => void this.reorderSelection('forward')),
      this.createIconButton(orderGroup, 'arrow-down', 'Send backward', () => void this.reorderSelection('backward')),
      this.createIconButton(orderGroup, 'send-to-back', 'Send to back', () => void this.reorderSelection('back'))
    );

    const groupGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });
    this.groupButton = this.createIconButton(groupGroup, 'group', 'Group objects', () => void this.groupSelection());
    this.ungroupButton = this.createIconButton(groupGroup, 'ungroup', 'Ungroup objects', () => void this.ungroupSelection());
  }

  private getSelectedIndices(): number[] {
    if (this.selectedShapeIndices.size > 0) return [...this.selectedShapeIndices];
    if (this.selectedShapeIndex !== null) return [this.selectedShapeIndex];
    return [];
  }

  private toggleShapeInSelection(shapeIndex: number): void {
    const next = new Set(this.getSelectedIndices());
    if (next.has(shapeIndex)) {
      next.delete(shapeIndex);
    } else {
      next.add(shapeIndex);
    }
    this.applyMultiSelection([...next]);
  }

  private collectSelectedTransforms(): { index: number; transform: ShapeTransform }[] {
    if (!this.engine || !this.svgEl) return [];
    const result: { index: number; transform: ShapeTransform }[] = [];
    for (const index of this.getSelectedIndices()) {
      const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
      if (isSVGGElement(shape)) {
        result.push({ index, transform: cloneTransform(this.engine.getShapeTransform(shape)) });
      }
    }
    return result;
  }

  private async alignSelectedShapes(mode: AlignMode): Promise<void> {
    if (!this.ensureEditable('align objects')) return;

    const boxes = this.collectSelectedTransforms();
    if (boxes.length < 2) {
      new Notice('Select at least two objects to align.');
      return;
    }

    const minX = Math.min(...boxes.map((box) => box.transform.x));
    const maxX = Math.max(...boxes.map((box) => box.transform.x + box.transform.cx));
    const minY = Math.min(...boxes.map((box) => box.transform.y));
    const maxY = Math.max(...boxes.map((box) => box.transform.y + box.transform.cy));

    const updates = boxes.map(({ index, transform }) => {
      const next = cloneTransform(transform);
      switch (mode) {
        case 'left':
          next.x = minX;
          break;
        case 'center':
          next.x = Math.round((minX + maxX) / 2 - next.cx / 2);
          break;
        case 'right':
          next.x = maxX - next.cx;
          break;
        case 'top':
          next.y = minY;
          break;
        case 'middle':
          next.y = Math.round((minY + maxY) / 2 - next.cy / 2);
          break;
        case 'bottom':
          next.y = maxY - next.cy;
          break;
      }
      return { index, transform: next };
    });

    await this.commitGroupTransforms(updates, 'Align objects');
  }

  private async distributeSelectedShapes(axis: DistributeAxis): Promise<void> {
    if (!this.ensureEditable('distribute objects')) return;

    const boxes = this.collectSelectedTransforms();
    if (boxes.length < 3) {
      new Notice('Select at least three objects to distribute.');
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

    await this.commitGroupTransforms(updates, 'Distribute objects');
  }

  private async reorderSelection(mode: ShapeReorderMode): Promise<void> {
    if (!this.engine || !this.ensureEditable('reorder objects')) return;

    const indices = this.getSelectedIndices().filter((index) => index >= 0);
    if (indices.length === 0) return;

    try {
      const history = await this.captureHistoryEntry('Reorder objects');
      const newIndices = await this.engine.reorderShapes(this.currentSlide, indices, mode);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.applyMultiSelection(newIndices.filter((index) => index >= 0));
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not reorder objects: ${cleanError(error)}`);
    }
  }

  private async groupSelection(): Promise<void> {
    if (!this.engine || !this.ensureEditable('group objects')) return;

    const indices = this.getSelectedIndices().filter((index) => index >= 0);
    if (indices.length < 2) {
      new Notice('Select at least two objects to group.');
      return;
    }

    try {
      const history = await this.captureHistoryEntry('Group objects');
      const groupIndex = await this.engine.groupShapes(this.currentSlide, indices);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(groupIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not group objects: ${cleanError(error)}`);
    }
  }

  private async ungroupSelection(): Promise<void> {
    if (!this.engine || !this.ensureEditable('ungroup objects')) return;

    if (!this.isSingleGroupSelected()) {
      new Notice('Select a single group to ungroup.');
      return;
    }

    const [groupIndex] = this.getSelectedIndices();
    if (groupIndex === undefined) return;

    try {
      const history = await this.captureHistoryEntry('Ungroup objects');
      const newIndices = await this.engine.ungroupShapes(this.currentSlide, groupIndex);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.applyMultiSelection(newIndices.filter((index) => index >= 0));
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not ungroup objects: ${cleanError(error)}`);
    }
  }

  private isSingleGroupSelected(): boolean {
    const indices = this.getSelectedIndices().filter((index) => index >= 0);
    if (indices.length !== 1) return false;
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${indices[0]}"]`);
    return shape?.getAttribute('data-ooxml-shape-type') === 'group';
  }

  private updateArrangeAvailability(): void {
    const canEdit = this.canEdit();
    const count = this.getSelectedIndices().filter((index) => index >= 0).length;
    for (const button of this.alignButtons) {
      this.updateObjectClipboardButton(button, canEdit && count >= 2);
    }
    for (const button of this.distributeButtons) {
      this.updateObjectClipboardButton(button, canEdit && count >= 3);
    }
    for (const button of this.zOrderButtons) {
      this.updateObjectClipboardButton(button, canEdit && count >= 1);
    }
    this.updateObjectClipboardButton(this.groupButton, canEdit && count >= 2);
    this.updateObjectClipboardButton(this.ungroupButton, canEdit && this.isSingleGroupSelected());
  }

  private async nudgeSelection(key: string, large: boolean): Promise<void> {
    if (!this.engine || !this.ensureEditable('move objects')) return;

    const stepEmu = this.engine.pxToEmu(large ? 10 : 1);
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
    await this.commitGroupTransforms(updates, 'Nudge objects');
  }

  private getSnapTargets(excluded: Set<number>): { xs: number[]; ys: number[] } {
    const xs: number[] = [];
    const ys: number[] = [];
    if (!this.engine || !this.svgEl) return { xs, ys };

    this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      if (!isSVGGElement(shape)) return;
      if (shape.parentElement?.closest('g[data-ooxml-shape-idx]')) return;
      const index = getShapeIndex(shape);
      if (index === null || excluded.has(index)) return;
      const transform = this.engine?.getShapeTransform(shape);
      if (!transform) return;
      xs.push(transform.x, transform.x + transform.cx / 2, transform.x + transform.cx);
      ys.push(transform.y, transform.y + transform.cy / 2, transform.y + transform.cy);
    });

    const size = getSvgIntrinsicSize(this.svgEl);
    const scale = this.engine.getSlideScale(this.svgEl);
    if (size && scale) {
      const width = size.width * scale;
      const height = size.height * scale;
      xs.push(0, width / 2, width);
      ys.push(0, height / 2, height);
    }
    return { xs, ys };
  }

  private computeSnap(
    box: { x: number; y: number; cx: number; cy: number },
    excluded: Set<number>
  ): { dx: number; dy: number; guideX: number | null; guideY: number | null } {
    const result = { dx: 0, dy: 0, guideX: null as number | null, guideY: null as number | null };
    if (!this.engine || !this.svgEl) return result;

    const ctm = this.svgEl.getScreenCTM();
    const scale = this.engine.getSlideScale(this.svgEl);
    if (!ctm || !scale || ctm.a === 0 || ctm.d === 0) return result;

    const thresholdX = (SNAP_THRESHOLD_PX * scale) / ctm.a;
    const thresholdY = (SNAP_THRESHOLD_PX * scale) / ctm.d;
    const targets = this.getSnapTargets(excluded);
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

  private emuPointToPane(emuX: number, emuY: number): PointerPoint | null {
    if (!this.engine || !this.svgEl || !this.canvasPane) return null;
    const ctm = this.svgEl.getScreenCTM();
    const scale = this.engine.getSlideScale(this.svgEl);
    if (!ctm || !scale) return null;

    const screenX = (emuX / scale) * ctm.a + ctm.e;
    const screenY = (emuY / scale) * ctm.d + ctm.f;
    const paneRect = this.canvasPane.getBoundingClientRect();
    return {
      x: screenX - paneRect.left + this.canvasPane.scrollLeft,
      y: screenY - paneRect.top + this.canvasPane.scrollTop
    };
  }

  private positionOverlayFromTransform(transform: ShapeTransform): void {
    if (!this.selectionOverlay) return;
    const topLeft = this.emuPointToPane(transform.x, transform.y);
    const bottomRight = this.emuPointToPane(transform.x + transform.cx, transform.y + transform.cy);
    if (!topLeft || !bottomRight) return;
    this.selectionOverlay.setCssProps({
      left: `${topLeft.x}px`,
      top: `${topLeft.y}px`,
      width: `${Math.max(0, bottomRight.x - topLeft.x)}px`,
      height: `${Math.max(0, bottomRight.y - topLeft.y)}px`
    });
  }

  private updateSnapGuides(guideXEmu: number | null, guideYEmu: number | null): void {
    this.clearSnapGuides();
    if (!this.canvasPane || !this.slideSurface) return;

    const surface = this.getElementBox(this.slideSurface);
    if (!surface) return;

    if (guideXEmu !== null) {
      const point = this.emuPointToPane(guideXEmu, 0);
      if (point) {
        const guide = this.canvasPane.createDiv({
          cls: 'native-powerpoint-snap-guide native-powerpoint-snap-guide-vertical'
        });
        guide.setCssProps({
          left: `${point.x}px`,
          top: `${surface.top}px`,
          height: `${surface.height}px`
        });
        this.snapGuides.push(guide);
      }
    }
    if (guideYEmu !== null) {
      const point = this.emuPointToPane(0, guideYEmu);
      if (point) {
        const guide = this.canvasPane.createDiv({
          cls: 'native-powerpoint-snap-guide native-powerpoint-snap-guide-horizontal'
        });
        guide.setCssProps({
          left: `${surface.left}px`,
          top: `${point.y}px`,
          width: `${surface.width}px`
        });
        this.snapGuides.push(guide);
      }
    }
  }

  private clearSnapGuides(): void {
    for (const guide of this.snapGuides) {
      guide.remove();
    }
    this.snapGuides = [];
  }

  private startRotateDrag(event: PointerEvent): void {
    if (!this.engine || this.selectedTransform === null || !this.selectionOverlay) return;
    if (!this.ensureEditable('rotate object')) return;

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
      startTransform: cloneTransform(this.selectedTransform),
      latestTransform: cloneTransform(this.selectedTransform),
      centerClientX,
      centerClientY,
      startAngle: Math.atan2(event.clientY - centerClientY, event.clientX - centerClientX)
    };
  }

  private toggleInsertMenu(
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

  private closeInsertMenus(): void {
    this.activeInsertMenu?.remove();
    this.activeInsertMenu = null;
  }

  private openVaultImagePicker(): void {
    new VaultImageSuggestModal(this.app, (file) => void this.insertImageFromVaultFile(file)).open();
  }

  private openInsertTableModal(): void {
    if (!this.ensureEditable('insert table')) return;
    new InsertTableModal(this.app, (rows, cols) => void this.insertTable(rows, cols)).open();
  }

  // Google Slides-style size picker: a hover grid that matches the look of the
  // other toolbar popovers (color, font) instead of a separate modal dialog.
  private openTableSizePicker(anchor: HTMLElement | null): void {
    if (!this.ensureEditable('insert table')) return;
    if (!anchor) {
      this.openInsertTableModal();
      return;
    }

    const cols = 10;
    const rows = 8;
    this.openToolbarPopover(anchor, (popover) => {
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
          this.bindToolbarButton(cell, () => {
            this.closeToolbarPopover();
            void this.insertTable(r + 1, c + 1);
          });
          cells.push(cell);
        }
      }

      grid.addEventListener('pointerleave', () => highlight(0, 0));
    });
  }

  private async insertImageFromVaultFile(file: TFile): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert image')) return;

    try {
      const bytes = await this.app.vault.readBinary(file);
      const history = await this.captureHistoryEntry('Insert image');
      const shapeIndex = this.engine.addImage(
        this.currentSlide,
        new Uint8Array(bytes),
        getImageMimeType(file.extension)
      );
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not insert image: ${cleanError(error)}`);
    }
  }

  private async insertImageFromLocalFile(file: File): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert image')) return;

    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const history = await this.captureHistoryEntry('Insert image');
      const shapeIndex = this.engine.addImage(
        this.currentSlide,
        bytes,
        file.type || getImageMimeType(file.name.split('.').pop() ?? 'png')
      );
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not insert image: ${cleanError(error)}`);
    }
  }

  private async insertShape(geometry: InsertableShapeGeometry): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert shape')) return;

    try {
      const history = await this.captureHistoryEntry('Insert shape');
      const shapeIndex = this.engine.addShapeGeometry(this.currentSlide, geometry);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not insert shape: ${cleanError(error)}`);
    }
  }

  private async insertTable(rows: number, cols: number): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert table')) return;

    try {
      const history = await this.captureHistoryEntry('Insert table');
      const shapeIndex = await this.engine.addTable(this.currentSlide, rows, cols);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not insert table: ${cleanError(error)}`);
    }
  }

  private async insertChart(): Promise<void> {
    if (!this.engine || !this.ensureEditable('insert chart')) return;

    try {
      const history = await this.captureHistoryEntry('Insert chart');
      const shapeIndex = await this.engine.addChart(this.currentSlide);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not insert chart: ${cleanError(error)}`);
    }
  }

  private async applyListStyle(style: ParagraphListStyle): Promise<void> {
    if (!this.engine || !this.ensureEditable('format text')) return;

    const textTarget = this.getTextEditTarget(this.activeEditorTarget);
    const shapeIndex = textTarget?.shapeIndex ?? this.selectedShapeIndex;
    if (shapeIndex === null) {
      new Notice('Select a text box or place the caret in text first.');
      return;
    }

    const paragraphIndex = textTarget?.kind === 'shape-paragraph' ? textTarget.paragraphIndex : 0;
    try {
      const history = await this.captureHistoryEntry(
        style === 'bullet' ? 'Bulleted list' : style === 'number' ? 'Numbered list' : 'Remove list'
      );
      await this.engine.applyListStyle(this.currentSlide, shapeIndex, paragraphIndex, style);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not update list style: ${cleanError(error)}`);
    }
  }

  private createFindPanel(): void {
    // Mounted on <body> (not inside the editor layout) so the fixed-position
    // dropdown is positioned relative to the viewport. Ancestors in the editor
    // tree use CSS transforms for zoom, which would otherwise make a
    // position:fixed child resolve against the transformed box and render in
    // the wrong place.
    this.findPanelEl?.remove();
    const panel = activeDocument.body.createDiv({ cls: 'native-powerpoint-find-panel native-powerpoint-find-panel-floating native-powerpoint-light-surface' });
    this.findPanelEl = panel;

    const findRow = panel.createDiv({ cls: 'native-powerpoint-find-row' });

    const toggleButton = findRow.createEl('button', {
      cls: 'native-powerpoint-find-btn native-powerpoint-find-replace-toggle',
      attr: { 'aria-label': 'Toggle replace' }
    });
    setIcon(toggleButton, 'chevron-right');
    this.findReplaceToggleEl = toggleButton;

    const input = findRow.createEl('input', {
      cls: 'native-powerpoint-find-input',
      type: 'search',
      attr: {
        'aria-label': 'Find text in presentation',
        placeholder: 'Find text'
      }
    });
    this.findInputEl = input;

    this.findStatusEl = findRow.createDiv({ cls: 'native-powerpoint-find-status', text: 'No search' });

    const previousButton = findRow.createEl('button', {
      cls: 'native-powerpoint-find-btn',
      attr: { 'aria-label': 'Previous match' }
    });
    setIcon(previousButton, 'chevron-up');

    const nextButton = findRow.createEl('button', {
      cls: 'native-powerpoint-find-btn',
      attr: { 'aria-label': 'Next match' }
    });
    setIcon(nextButton, 'chevron-down');

    const closeButton = findRow.createEl('button', {
      cls: 'native-powerpoint-find-btn',
      attr: { 'aria-label': 'Close find' }
    });
    setIcon(closeButton, 'x');

    const replaceRow = panel.createDiv({ cls: 'native-powerpoint-find-replace-row' });

    const replaceInput = replaceRow.createEl('input', {
      cls: 'native-powerpoint-find-input',
      type: 'text',
      attr: {
        'aria-label': 'Replacement text',
        placeholder: 'Replace with'
      }
    });
    this.findReplaceInputEl = replaceInput;

    const replaceButton = replaceRow.createEl('button', {
      cls: 'native-powerpoint-find-replace-btn',
      text: 'Replace',
      attr: { 'aria-label': 'Replace current match' }
    });

    const replaceAllButton = replaceRow.createEl('button', {
      cls: 'native-powerpoint-find-replace-btn',
      text: 'Replace all',
      attr: { 'aria-label': 'Replace all matches' }
    });

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
    replaceInput.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        this.closeFindPanel();
      } else if (event.key === 'Enter') {
        event.preventDefault();
        if (event.shiftKey) {
          void this.replaceAllMatches();
        } else {
          void this.replaceCurrentMatch();
        }
      }
    });
    toggleButton.addEventListener('click', () => this.setFindReplaceMode(!this.isFindReplaceMode));
    previousButton.addEventListener('click', () => void this.moveFindMatch(-1));
    nextButton.addEventListener('click', () => void this.moveFindMatch(1));
    closeButton.addEventListener('click', () => this.closeFindPanel());
    replaceButton.addEventListener('click', () => void this.replaceCurrentMatch());
    replaceAllButton.addEventListener('click', () => void this.replaceAllMatches());

    this.setFindReplaceMode(false);
  }

  private setFindReplaceMode(enabled: boolean): void {
    this.isFindReplaceMode = enabled;
    this.findPanelEl?.toggleClass('is-replace-mode', enabled);
    if (this.findReplaceToggleEl) {
      setIcon(this.findReplaceToggleEl, enabled ? 'chevron-down' : 'chevron-right');
      this.findReplaceToggleEl.setAttribute('aria-expanded', String(enabled));
    }
  }

  private createIconButton(container: HTMLElement, icon: string, label: string, onClick: () => void): HTMLButtonElement {
    const button = container.createEl('button', {
      cls: 'native-powerpoint-toolbar-btn',
      attr: { 'aria-label': label }
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

  private buildSlideSvgElement(index: number): SVGSVGElement | null {
    if (!this.engine || index < 0 || index >= this.engine.slideCount) return null;

    const safeSvg = this.prepareSvgForRender(this.engine.renderSlide(index).svg, true);
    if (!safeSvg.allowed) return null;

    const element = createSvgElementFromString(safeSvg.svg, this.contentEl.ownerDocument);
    if (!element) return null;

    this.engine.applyFontFidelity(element);
    this.engine.formatChartAxisLabels(element, index);
    normalizeSvgForDisplay(element);
    return element;
  }

  private startPresentation(): void {
    if (!this.engine || this.engine.slideCount === 0) {
      new Notice('Open a presentation with at least one slide to present.');
      return;
    }

    this.presentController?.dispose();

    const controller = new PowerPointPresentController({
      ownerDocument: this.contentEl.ownerDocument,
      slideCount: this.engine.slideCount,
      startIndex: this.currentSlide,
      renderSlide: (index) => this.buildSlideSvgElement(index),
      onExit: (lastIndex) => {
        this.presentController = null;
        if (this.engine && lastIndex >= 0 && lastIndex < this.engine.slideCount) {
          void this.goToSlide(lastIndex);
        }
      }
    });

    this.presentController = controller;
    controller.start();
  }

  private openExportMenu(anchor: HTMLElement): void {
    if (!this.engine || this.engine.slideCount === 0) {
      new Notice('Open a presentation with at least one slide to export.');
      return;
    }

    const menu = this.createNativeMenu();
    menu.addItem((item) =>
      item
        .setTitle('Current slide as PNG')
        .setIcon('image')
        .onClick(() => void this.exportCurrentSlideAsPng())
    );
    menu.addItem((item) =>
      item
        .setTitle('Current slide as PDF')
        .setIcon('file-output')
        .onClick(() => void this.exportDeckAsPdf(true))
    );
    menu.addSeparator();
    menu.addItem((item) =>
      item
        .setTitle('Whole deck as PDF')
        .setIcon('file-output')
        .onClick(() => void this.exportDeckAsPdf(false))
    );
    menu.addItem((item) =>
      item
        .setTitle('Whole deck as PNGs (zip)')
        .setIcon('file-archive')
        .onClick(() => void this.exportDeckAsPngZip())
    );

    const rect = anchor.getBoundingClientRect();
    menu.showAtPosition({ x: rect.left, y: rect.bottom + 4 });
  }

  private collectExportSvgElements(indices: number[]): SVGSVGElement[] {
    const elements: SVGSVGElement[] = [];
    for (const index of indices) {
      const element = this.buildSlideSvgElement(index);
      if (element) elements.push(element);
    }
    return elements;
  }

  private async exportCurrentSlideAsPng(): Promise<void> {
    if (!this.engine) return;

    try {
      const element = this.buildSlideSvgElement(this.currentSlide);
      if (!element) {
        throw new Error('This slide could not be rendered for export.');
      }

      const bytes = await exportSlideToPng(element, this.contentEl.ownerDocument);
      await this.saveExportArtifact(
        `${this.getExportBaseName()}-slide-${this.currentSlide + 1}`,
        'png',
        bytes
      );
    } catch (error) {
      new Notice(`Could not export slide as PNG: ${cleanError(error)}`);
    }
  }

  private async exportDeckAsPdf(currentSlideOnly: boolean): Promise<void> {
    if (!this.engine) return;

    try {
      const indices = currentSlideOnly
        ? [this.currentSlide]
        : Array.from({ length: this.engine.slideCount }, (_, index) => index);
      const elements = this.collectExportSvgElements(indices);
      if (elements.length === 0) {
        throw new Error('No slides could be rendered for export.');
      }

      new Notice(currentSlideOnly ? 'Exporting slide to PDF...' : 'Exporting deck to PDF...');
      const bytes = await exportSlidesToPdf(elements, this.contentEl.ownerDocument);
      const baseName = currentSlideOnly
        ? `${this.getExportBaseName()}-slide-${this.currentSlide + 1}`
        : this.getExportBaseName();
      await this.saveExportArtifact(baseName, 'pdf', bytes);
    } catch (error) {
      new Notice(`Could not export PDF: ${cleanError(error)}`);
    }
  }

  private async exportDeckAsPngZip(): Promise<void> {
    if (!this.engine) return;

    try {
      const indices = Array.from({ length: this.engine.slideCount }, (_, index) => index);
      const elements = this.collectExportSvgElements(indices);
      if (elements.length === 0) {
        throw new Error('No slides could be rendered for export.');
      }

      new Notice('Exporting deck to PNG images...');
      const baseName = this.getExportBaseName();
      const bytes = await exportSlidesToPngZip(elements, this.contentEl.ownerDocument, baseName);
      await this.saveExportArtifact(`${baseName}-slides`, 'zip', bytes);
    } catch (error) {
      new Notice(`Could not export PNG images: ${cleanError(error)}`);
    }
  }

  private getExportBaseName(): string {
    const file = this.loadedFile || this.file;
    return file?.basename || 'presentation';
  }

  private getAvailableNumberedPath(path: string): string {
    const lastSlashIndex = path.lastIndexOf('/');
    const folderPrefix = lastSlashIndex >= 0 ? `${path.slice(0, lastSlashIndex)}/` : '';
    const fileName = lastSlashIndex >= 0 ? path.slice(lastSlashIndex + 1) : path;
    const extensionIndex = fileName.lastIndexOf('.');
    const baseName = extensionIndex > 0 ? fileName.slice(0, extensionIndex) : fileName;
    const extension = extensionIndex > 0 ? fileName.slice(extensionIndex) : '';

    for (let index = 2; index < 1000; index += 1) {
      const candidatePath = normalizePath(`${folderPrefix}${baseName} ${index}${extension}`);
      if (!this.app.vault.getAbstractFileByPath(candidatePath)) {
        return candidatePath;
      }
    }

    return normalizePath(`${folderPrefix}${baseName} ${Date.now()}${extension}`);
  }

  private async saveExportArtifact(baseName: string, extension: string, data: ArrayBuffer): Promise<void> {
    const source = this.loadedFile || this.file;
    const folderPath = source?.parent?.path;
    const folderPrefix = folderPath && folderPath !== '/' ? `${folderPath}/` : '';
    const safeBaseName = baseName.replace(/[\\/:*?"<>|]/g, '_') || 'presentation';
    let targetPath = normalizePath(`${folderPrefix}${safeBaseName}.${extension}`);

    const existing = this.app.vault.getAbstractFileByPath(targetPath);
    if (existing) {
      if (!(existing instanceof TFile)) {
        new Notice(`${targetPath} already exists and is not a file.`);
        return;
      }

      const choice = await new Promise<'replace' | 'keep-both' | 'cancel'>((resolve) => {
        const menu = this.createNativeMenu();
        menu.addItem((item) => item.setTitle('Replace existing file').setIcon('refresh-cw').onClick(() => resolve('replace')));
        menu.addItem((item) => item.setTitle('Keep both (numbered copy)').setIcon('copy-plus').onClick(() => resolve('keep-both')));
        menu.addItem((item) => item.setTitle('Cancel export').setIcon('x').onClick(() => resolve('cancel')));
        menu.onHide(() => resolve('cancel'));
        const view = this.contentEl.ownerDocument.defaultView;
        const x = view ? view.innerWidth / 2 : 200;
        const y = view ? view.innerHeight / 3 : 200;
        menu.showAtPosition({ x, y });
      });

      if (choice === 'cancel') return;
      if (choice === 'keep-both') {
        targetPath = this.getAvailableNumberedPath(targetPath);
      }
    }

    const existingTarget = this.app.vault.getAbstractFileByPath(targetPath);
    if (existingTarget instanceof TFile) {
      await this.app.vault.modifyBinary(existingTarget, data);
    } else {
      await this.app.vault.createBinary(targetPath, data);
    }

    new Notice(`Exported to ${targetPath}`);
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

  private openFindPanel(options: { replace?: boolean } = {}): void {
    if (!this.engine || this.isLoading) {
      new Notice('Open a loaded PowerPoint file to search it.');
      return;
    }

    this.findPanelEl?.addClass('is-open');
    if (options.replace) {
      this.setFindReplaceMode(true);
    }
    this.findButtonEl?.addClass('is-active');
    const seedText = this.getSelectedFindSeedText();
    if (seedText && this.findInputEl && !this.findInputEl.value.trim()) {
      this.findInputEl.value = seedText;
    }

    this.attachFindPanelDismissHandlers();
    void this.refreshFindMatches({ reveal: Boolean(this.findInputEl?.value.trim()) });
    window.requestAnimationFrame(() => {
      this.positionFindPanel();
      this.findInputEl?.focus();
      this.findInputEl?.select();
    });
  }

  private toggleFindPanel(options: { replace?: boolean } = {}): void {
    if (this.findPanelEl?.hasClass('is-open')) {
      this.closeFindPanel();
      return;
    }
    this.openFindPanel(options);
  }

  private positionFindPanel(): void {
    const panel = this.findPanelEl;
    const anchor = this.findButtonEl;
    if (!panel || !anchor || !panel.hasClass('is-open')) return;

    const rect = anchor.getBoundingClientRect();
    const panelWidth = panel.offsetWidth || 320;
    const gap = 6;
    const margin = 8;
    let left = rect.right - panelWidth;
    left = Math.max(margin, Math.min(left, window.innerWidth - panelWidth - margin));
    const top = Math.min(rect.bottom + gap, window.innerHeight - panel.offsetHeight - margin);
    panel.style.left = `${Math.round(left)}px`;
    panel.style.top = `${Math.round(Math.max(margin, top))}px`;
  }

  private attachFindPanelDismissHandlers(): void {
    if (!this.findPanelDismissHandler) {
      this.findPanelDismissHandler = (event: Event) => {
        const target = event.target;
        if (!(target instanceof Node)) return;
        if (this.findPanelEl?.contains(target)) return;
        if (this.findButtonEl?.contains(target)) return;
        this.closeFindPanel();
      };
      activeDocument.addEventListener('pointerdown', this.findPanelDismissHandler, true);
    }
    if (!this.findPanelRepositionHandler) {
      this.findPanelRepositionHandler = () => this.positionFindPanel();
      window.addEventListener('resize', this.findPanelRepositionHandler);
      window.addEventListener('scroll', this.findPanelRepositionHandler, true);
    }
  }

  private detachFindPanelDismissHandlers(): void {
    if (this.findPanelDismissHandler) {
      activeDocument.removeEventListener('pointerdown', this.findPanelDismissHandler, true);
      this.findPanelDismissHandler = null;
    }
    if (this.findPanelRepositionHandler) {
      window.removeEventListener('resize', this.findPanelRepositionHandler);
      window.removeEventListener('scroll', this.findPanelRepositionHandler, true);
      this.findPanelRepositionHandler = null;
    }
  }

  private async replaceCurrentMatch(): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('replace text')) return;

    const query = this.findInputEl?.value.trim() ?? '';
    if (!query) {
      this.findInputEl?.focus();
      return;
    }

    if (this.findMatches.length === 0) {
      await this.refreshFindMatches();
      if (this.findMatches.length === 0) {
        new Notice('No matches to replace.');
        return;
      }
    }

    const match = this.findMatches[this.currentFindMatchIndex];
    if (!match || match.shapeIndex === null) {
      new Notice('Select a match to replace, or use Replace all.');
      return;
    }

    const replacement = this.findReplaceInputEl?.value ?? '';
    try {
      const history = await this.captureHistoryEntry('Replace text');
      const count = await this.engine.replaceText(query, replacement, {
        slideIndex: match.slideIndex,
        shapeIndex: match.shapeIndex
      });
      if (count === 0) {
        new Notice('No matches to replace.');
        return;
      }
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
      await this.refreshFindMatches({ reveal: true });
      new Notice(count === 1 ? 'Replaced 1 match.' : `Replaced ${count} matches.`);
    } catch (error) {
      new Notice(`Could not replace text: ${cleanError(error)}`);
    }
  }

  private async replaceAllMatches(): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('replace text')) return;

    const query = this.findInputEl?.value.trim() ?? '';
    if (!query) {
      this.findInputEl?.focus();
      return;
    }

    const replacement = this.findReplaceInputEl?.value ?? '';
    try {
      const history = await this.captureHistoryEntry('Replace all text');
      const count = await this.engine.replaceText(query, replacement);
      if (count === 0) {
        new Notice('No matches to replace.');
        return;
      }
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
      await this.refreshFindMatches({ reveal: true });
      new Notice(count === 1 ? 'Replaced 1 match.' : `Replaced ${count} matches.`);
    } catch (error) {
      new Notice(`Could not replace text: ${cleanError(error)}`);
    }
  }

  private closeFindPanel(): void {
    this.findPanelEl?.removeClass('is-open');
    this.findButtonEl?.removeClass('is-active');
    this.detachFindPanelDismissHandlers();
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

    // Intentionally do not select the matched shape: a selection would draw a
    // box/outline around the whole text frame. Find should only highlight the
    // matched characters themselves.
    this.applyFindHighlight();
    this.updateFindStatus();
  }

  private clearFindHighlight(): void {
    for (const rect of this.findHighlightRects) {
      rect.remove();
    }
    this.findHighlightRects = [];
    this.svgEl?.querySelectorAll('.native-powerpoint-find-current').forEach((element) => {
      element.removeClass('native-powerpoint-find-current');
    });
  }

  private applyFindHighlight(): void {
    this.clearFindHighlight();
    if (!this.svgEl) return;

    const query = this.findInputEl?.value ?? '';
    const trimmed = query.trim();
    if (!trimmed) return;

    const queryLower = trimmed.toLocaleLowerCase();
    const currentMatch = this.findMatches[this.currentFindMatchIndex];
    const currentShapeIndex = currentMatch && currentMatch.slideIndex === this.currentSlide
      ? currentMatch.shapeIndex
      : null;

    for (const shape of Array.from(this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]'))) {
      const shapeIndex = getShapeIndex(shape);
      const isCurrent = shapeIndex !== null && shapeIndex === currentShapeIndex;
      this.highlightFindOccurrencesInShape(shape, queryLower, isCurrent);
    }
  }

  private highlightFindOccurrencesInShape(shape: Element, queryLower: string, isCurrent: boolean): void {
    if (!queryLower) return;
    for (const paragraph of this.getShapeTextParagraphs(shape)) {
      const text = this.getParagraphLeafText(paragraph);
      if (!text) continue;

      const lower = text.toLocaleLowerCase();
      let from = 0;
      while (from <= lower.length) {
        const index = lower.indexOf(queryLower, from);
        if (index === -1) break;
        this.renderFindHighlightRects(paragraph, index, index + queryLower.length, isCurrent);
        from = index + Math.max(1, queryLower.length);
      }
    }
  }

  private renderFindHighlightRects(
    element: SVGTextElement | SVGTSpanElement,
    start: number,
    end: number,
    isCurrent: boolean
  ): void {
    const boxes = this.getSvgInlineSelectionBoxes(element, start, end);
    const textElement = element.closest('text');
    const parent = textElement?.parentNode;
    if (!isSVGTextElement(textElement) || !parent) return;

    for (const box of boxes) {
      const rect = activeDocument.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect.classList.add('native-powerpoint-find-highlight');
      if (isCurrent) rect.classList.add('is-current');
      rect.setAttribute('x', this.formatSvgNumber(box.x));
      rect.setAttribute('y', this.formatSvgNumber(box.y));
      rect.setAttribute('width', this.formatSvgNumber(box.width));
      rect.setAttribute('height', this.formatSvgNumber(box.height));
      rect.setAttribute('rx', '1');
      parent.insertBefore(rect, textElement);
      this.findHighlightRects.push(rect);
    }
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
        const target = isElement(event.target) ? event.target : null;
        if (!target?.closest('.modal')) {
          event.preventDefault();
          event.stopImmediatePropagation();
          this.openFindPanel();
          return;
        }
      }

      if (!this.isActivePowerPointView()) return;
      if (this.activeEditor && activeDocument.activeElement === this.activeEditor) return;

      const target = isElement(event.target) ? event.target : null;
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

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'a') {
        event.preventDefault();
        event.stopImmediatePropagation();
        if (this.lastInteractionRegion === 'thumbnails') {
          this.selectAllSlides();
        } else {
          this.selectAllShapes();
        }
        return;
      }

      const isArrowKey =
        event.key === 'ArrowUp'
        || event.key === 'ArrowDown'
        || event.key === 'ArrowLeft'
        || event.key === 'ArrowRight';
      const hasShapeSelection = this.selectedShapeIndex !== null || this.selectedShapeIndices.size > 0;
      if (isArrowKey && hasShapeSelection) {
        event.preventDefault();
        event.stopImmediatePropagation();
        void this.nudgeSelection(event.key, event.shiftKey);
        return;
      }

      if (event.key === 'Escape' && this.selectedSlideIndices.size > 0) {
        event.preventDefault();
        this.clearSlideSelection();
        return;
      }

      if (
        (event.key === 'Delete' || event.key === 'Backspace')
        && this.lastInteractionRegion === 'thumbnails'
      ) {
        if (this.selectedSlideIndices.size > 0) {
          event.preventDefault();
          void this.deleteSelectedSlides();
          return;
        }
        event.preventDefault();
        void this.deleteSlide();
        return;
      }

      if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
        event.preventDefault();
        void this.goToSlide(this.currentSlide - 1);
      } else if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
        event.preventDefault();
        void this.goToSlide(this.currentSlide + 1);
      } else if (event.key === 'Delete' || event.key === 'Backspace') {
        if (this.selectedShapeIndex !== null || this.selectedShapeIndices.size > 0) {
          event.preventDefault();
          void this.deleteSelectedShape();
        }
      }
    };

    this.registerDomEvent(window, 'keydown', handleKeyDown, true);
    this.registerDomEvent(activeDocument, 'keydown', handleKeyDown, true);

    this.registerDomEvent(window, 'resize', () => this.updateSlideScale());
    this.registerDomEvent(activeDocument, 'pointermove', this.handleDragMove, true);
    this.registerDomEvent(activeDocument, 'pointerup', this.handleDragEnd, true);
    this.registerDomEvent(activeDocument, 'pointerdown', this.handleOutsideSlidePointerDown, true);
  }

  private isActivePowerPointView(): boolean {
    if (this.app.workspace.getActiveViewOfType(NativePowerPointView) === this) {
      return true;
    }

    if (this.contentEl.closest('.workspace-leaf.mod-active')) {
      return true;
    }

    const activeElement = activeDocument.activeElement;
    return Boolean(isNode(activeElement) && this.contentEl.contains(activeElement));
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
    if (this.findReplaceInputEl) this.findReplaceInputEl.value = '';
    this.setFindReplaceMode(false);
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
      if (this.suppressNextClick) {
        this.suppressNextClick = false;
        event.preventDefault();
        event.stopPropagation();
        return;
      }
      const target = isElement(event.target) ? event.target : null;
      if (this.suppressNextTextClick && target?.closest('text')) {
        event.preventDefault();
        event.stopPropagation();
        this.suppressNextTextClick = false;
        return;
      }
      this.suppressNextTextClick = false;

      const shape = target?.closest('g[data-ooxml-shape-idx]') ?? null;
      const shapeIndex = getShapeIndex(shape);
      const additive = event.shiftKey || event.ctrlKey || event.metaKey;
      if (shapeIndex === null) {
        if (!additive) this.clearSelection();
        return;
      }

      if (additive) {
        event.preventDefault();
        event.stopPropagation();
        this.toggleShapeInSelection(shapeIndex);
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
      const target = isElement(event.target) ? event.target : null;
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

      const target = isElement(event.target) ? event.target : null;
      if (target?.closest('text')) {
        this.handleInlineTextPointerDown(event, target);
        return;
      }

      if (this.activeEditor) {
        this.commitActiveTextEditing();
      }

      const shape = target?.closest('g[data-ooxml-shape-idx]') ?? null;
      const shapeIndex = getShapeIndex(shape);
      const additive = event.shiftKey || event.ctrlKey || event.metaKey;

      if (shapeIndex === null) {
        event.preventDefault();
        this.beginMarquee(event, additive);
        return;
      }

      if (additive) {
        // Let the click handler toggle this shape in/out of the selection
        // instead of starting a drag.
        event.preventDefault();
        return;
      }

      if (this.selectedShapeIndices.size > 1 && this.selectedShapeIndices.has(shapeIndex)) {
        event.preventDefault();
        if (this.ensureEditable('move objects')) {
          this.startGroupDrag(event);
        }
        return;
      }

      if (this.selectedShapeIndex === null) return;
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
      if (this.selectedSlideIndices.has(index)) item.addClass('is-selected');

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
      item.addEventListener('click', (event) => {
        this.lastInteractionRegion = 'thumbnails';
        if (event.shiftKey) {
          this.selectSlideRange(this.currentSlide, index);
        } else if (event.metaKey || event.ctrlKey) {
          this.toggleSlideSelection(index);
        } else {
          // A plain click both navigates to and selects the slide so it can be
          // deleted with the keyboard (matching Google Slides' filmstrip).
          this.selectedSlideIndices = new Set([index]);
          void this.goToSlide(index);
        }
      });
      item.addEventListener('contextmenu', (event) => {
        event.preventDefault();
        this.showSlideContextMenu(event, index);
      });
      this.registerThumbnailDrag(item, index);
    }
  }

  private registerThumbnailDrag(item: HTMLElement, index: number): void {
    item.draggable = this.canEdit();
    item.dataset.slideIndex = String(index);

    item.addEventListener('dragstart', (event) => {
      if (!this.canEdit()) {
        event.preventDefault();
        return;
      }
      this.thumbnailDragIndex = index;
      item.addClass('is-dragging');
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('text/plain', String(index));
      }
    });

    item.addEventListener('dragend', () => {
      this.thumbnailDragIndex = null;
      this.clearThumbnailDropIndicators();
      item.removeClass('is-dragging');
    });

    item.addEventListener('dragover', (event) => {
      if (this.thumbnailDragIndex === null) return;
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
      const after = this.isPointerInLowerHalf(event, item);
      this.clearThumbnailDropIndicators();
      item.addClass(after ? 'drop-after' : 'drop-before');
    });

    item.addEventListener('dragleave', () => {
      item.removeClass('drop-before');
      item.removeClass('drop-after');
    });

    item.addEventListener('drop', (event) => {
      event.preventDefault();
      const fromIndex = this.thumbnailDragIndex;
      this.thumbnailDragIndex = null;
      this.clearThumbnailDropIndicators();
      if (fromIndex === null) return;

      const after = this.isPointerInLowerHalf(event, item);
      let toIndex = after ? index + 1 : index;
      if (fromIndex < toIndex) toIndex -= 1;
      void this.reorderSlideByDrag(fromIndex, toIndex);
    });
  }

  private isPointerInLowerHalf(event: DragEvent, item: HTMLElement): boolean {
    const rect = item.getBoundingClientRect();
    return event.clientY > rect.top + rect.height / 2;
  }

  private clearThumbnailDropIndicators(): void {
    this.thumbnailContainer?.querySelectorAll('.drop-before, .drop-after').forEach((element) => {
      element.classList.remove('drop-before', 'drop-after');
    });
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

  private async deleteSelectedSlides(): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('delete slides')) return;

    const targets = Array.from(this.selectedSlideIndices).sort((a, b) => b - a);
    if (targets.length === 0) return;
    if (targets.length >= this.engine.slideCount) {
      new Notice('You cannot delete every slide.');
      return;
    }

    try {
      const history = await this.captureHistoryEntry('Delete slides');
      let resultIndex = this.currentSlide;
      for (const target of targets) {
        const result = await this.engine.deleteSlide(target);
        resultIndex = result.slideIndex;
      }
      this.currentSlide = resultIndex;
      this.clearSlideSelection();
      this.clearSelection();
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not delete slides: ${cleanError(error)}`);
    }
  }

  private async moveSlide(direction: -1 | 1): Promise<void> {
    await this.moveSlideAt(this.currentSlide, direction);
  }

  private async moveSlideAt(index: number, direction: -1 | 1): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('move slide')) return;
    if (index < 0 || index >= this.engine.slideCount) return;

    try {
      const history = await this.captureHistoryEntry('Move slide');
      const result = await this.engine.moveSlide(index, direction);
      if (result.slideIndex === index) return;

      this.currentSlide = result.slideIndex;
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not move slide: ${cleanError(error)}`);
    }
  }

  private async addSlideWithLayout(layout: SlideLayoutKind): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('add slide')) return;

    try {
      const history = await this.captureHistoryEntry('New slide');
      const result = await this.engine.addSlideWithLayout(this.currentSlide, layout);
      this.currentSlide = result.slideIndex;
      this.clearSelection();
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not add slide: ${cleanError(error)}`);
    }
  }

  private async duplicateSlide(targetIndex: number = this.currentSlide): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('duplicate slide')) return;
    if (targetIndex < 0 || targetIndex >= this.engine.slideCount) return;

    try {
      const history = await this.captureHistoryEntry('Duplicate slide');
      const result = await this.engine.duplicateSlide(targetIndex);
      this.currentSlide = result.slideIndex;
      this.clearSelection();
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not duplicate slide: ${cleanError(error)}`);
    }
  }

  private async deleteSlideAt(targetIndex: number): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('delete slide')) return;
    if (targetIndex < 0 || targetIndex >= this.engine.slideCount) return;

    try {
      const history = await this.captureHistoryEntry('Delete slide');
      const result = await this.engine.deleteSlide(targetIndex);
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

  private async reorderSlideByDrag(fromIndex: number, toIndex: number): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('reorder slides')) return;

    const slideCount = this.engine.slideCount;
    if (
      fromIndex === toIndex ||
      fromIndex < 0 ||
      fromIndex >= slideCount ||
      toIndex < 0 ||
      toIndex >= slideCount
    ) {
      return;
    }

    const order = Array.from({ length: slideCount }, (_, index) => index);
    const [moved] = order.splice(fromIndex, 1);
    if (moved === undefined) return;
    order.splice(toIndex, 0, moved);

    try {
      const history = await this.captureHistoryEntry('Reorder slides');
      await this.engine.reorderSlides(order);
      this.currentSlide = toIndex;
      this.clearSelection();
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not reorder slides: ${cleanError(error)}`);
    }
  }

  private showSlideContextMenu(event: MouseEvent, index: number): void {
    if (!this.engine) return;

    const menu = this.createNativeMenu();
    menu.addItem((item) =>
      item
        .setTitle('New slide')
        .setIcon('plus')
        .onClick(() => {
          this.currentSlide = index;
          void this.addSlideWithLayout('blank');
        })
    );
    menu.addItem((item) =>
      item
        .setTitle('Duplicate slide')
        .setIcon('files')
        .onClick(() => void this.duplicateSlide(index))
    );
    menu.addSeparator();
    menu.addItem((item) =>
      item
        .setTitle('Move up')
        .setIcon('arrow-up')
        .setDisabled(index <= 0)
        .onClick(() => void this.moveSlideAt(index, -1))
    );
    menu.addItem((item) =>
      item
        .setTitle('Move down')
        .setIcon('arrow-down')
        .setDisabled(!this.engine || index >= this.engine.slideCount - 1)
        .onClick(() => void this.moveSlideAt(index, 1))
    );
    menu.addSeparator();
    menu.addItem((item) =>
      item
        .setTitle('Delete slide')
        .setIcon('trash-2')
        .setDisabled(!this.engine || this.engine.slideCount <= 1)
        .onClick(() => void this.deleteSlideAt(index))
    );
    menu.showAtMouseEvent(event);
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
    if (!this.engine) return;
    if (this.selectedShapeIndices.size > 1) {
      await this.deleteSelectedShapes();
      return;
    }
    if (this.selectedShapeIndex === null) return;
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

  private async deleteSelectedShapes(): Promise<void> {
    if (!this.engine || this.selectedShapeIndices.size === 0) return;
    if (!this.ensureEditable('delete objects')) return;

    const indices = [...this.selectedShapeIndices].sort((a, b) => b - a);
    try {
      const history = await this.captureHistoryEntry('Delete objects');
      for (const index of indices) {
        this.engine.deleteShape(this.currentSlide, index);
      }
      this.clearSelection();
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
    } catch (error) {
      new Notice(`Could not delete objects: ${cleanError(error)}`);
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

  private async cutSelectedShape(): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) {
      new Notice('Select a slide object to cut.');
      return;
    }
    if (!this.ensureEditable('cut object')) return;

    try {
      this.objectClipboard = await this.engine.copyShape(this.currentSlide, this.selectedShapeIndex);
      this.updateObjectClipboardAvailability();
      await this.deleteSelectedShape();
      new Notice('Cut slide object.');
    } catch (error) {
      new Notice(`Could not cut object: ${cleanError(error)}`);
    }
  }

  /**
   * Paste without formatting. With an active inline text editor this inserts
   * the clipboard's plain text at the caret. Otherwise it falls back to the
   * regular object paste (slide objects carry their own formatting, so there is
   * no plain-text variant for them).
   */
  private async pasteWithoutFormatting(): Promise<void> {
    if (this.activeEditor) {
      if (!this.ensureEditable('paste text')) return;
      try {
        const text = await navigator.clipboard.readText();
        if (text) this.insertPlainTextIntoActiveEditor(text);
      } catch {
        new Notice('Plain text is not available on the clipboard.');
      }
      return;
    }

    await this.pasteCopiedShape();
  }

  private insertPlainTextIntoActiveEditor(text: string): void {
    const editor = this.activeEditor;
    if (!editor) return;

    const start = editor.selectionStart ?? editor.value.length;
    const end = editor.selectionEnd ?? editor.value.length;
    editor.value = `${editor.value.slice(0, start)}${text}${editor.value.slice(end)}`;
    const caret = start + text.length;
    editor.setSelectionRange(caret, caret);
    editor.dispatchEvent(new Event('input', { bubbles: true }));
  }

  private async rotateSelectedShape(deltaDegrees: number): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('rotate object')) return;

    const selected = this.getSelectedShapeElement();
    if (!selected) return;

    const transform = cloneTransform(this.engine.getShapeTransform(selected));
    const degrees = (((this.engine.ooxmlToDegrees(transform.rot) + deltaDegrees) % 360) + 360) % 360;
    transform.rot = this.engine.degreesToOoxml(degrees);
    await this.commitTransform(transform);
  }

  private async centerSelectedOnPage(axis: 'horizontal' | 'vertical'): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('center object')) return;

    const selected = this.getSelectedShapeElement();
    if (!selected) return;

    const transform = cloneTransform(this.engine.getShapeTransform(selected));
    const { cx, cy } = await this.engine.getSlideSizeEmu();
    if (axis === 'horizontal') {
      transform.x = Math.round((cx - transform.cx) / 2);
    } else {
      transform.y = Math.round((cy - transform.cy) / 2);
    }
    await this.commitTransform(transform);
  }

  private async flipSelectedShape(axis: 'horizontal' | 'vertical'): Promise<void> {
    if (this.selectedShapeIndex === null) return;
    await this.applyShapeMutation(
      this.selectedShapeIndex,
      'Flip object',
      'flip object',
      (slideIndex, shapeIndex) => this.engine!.flipShape(slideIndex, shapeIndex, axis)
    );
  }

  private openImageCropDialog(shapeIndex: number): void {
    if (!this.engine || !this.ensureEditable('crop image')) return;

    const current: ImageCrop = this.engine.getImageCrop(this.currentSlide, shapeIndex)
      ?? { left: 0, top: 0, right: 0, bottom: 0 };
    new ImageCropModal(this.app, current, (crop: ImageCropValues) => {
      void this.applyShapeMutation(
        shapeIndex,
        'Crop image',
        'crop image',
        (slideIndex, index) => this.engine!.setImageCrop(slideIndex, index, crop)
      );
    }).open();
  }

  private async resetSelectedImage(shapeIndex: number): Promise<void> {
    await this.applyShapeMutation(
      shapeIndex,
      'Reset image',
      'reset image',
      (slideIndex, index) => this.engine!.resetImage(slideIndex, index)
    );
  }

  private openReplaceImageVaultPicker(shapeIndex: number): void {
    if (!this.ensureEditable('replace image')) return;
    new VaultImageSuggestModal(this.app, (file) => void this.replaceImageWithVaultFile(shapeIndex, file)).open();
  }

  private async replaceImageWithVaultFile(shapeIndex: number, file: TFile): Promise<void> {
    if (!this.engine) return;
    const bytes = new Uint8Array(await this.app.vault.readBinary(file));
    await this.applyShapeMutation(
      shapeIndex,
      'Replace image',
      'replace image',
      (slideIndex, index) => this.engine!.replaceImage(slideIndex, index, bytes, getImageMimeType(file.extension))
    );
  }

  private async replaceImageWithLocalFile(shapeIndex: number, file: File): Promise<void> {
    if (!this.engine) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    const mimeType = file.type || getImageMimeType(file.name.split('.').pop() ?? 'png');
    await this.applyShapeMutation(
      shapeIndex,
      'Replace image',
      'replace image',
      (slideIndex, index) => this.engine!.replaceImage(slideIndex, index, bytes, mimeType)
    );
  }

  private async setSelectedImageAsBackground(shapeIndex: number): Promise<void> {
    if (!this.engine || !this.ensureEditable('set slide background')) return;

    try {
      const image = await this.engine.getShapeImageData(this.currentSlide, shapeIndex);
      if (!image) {
        new Notice('The selected object is not an image.');
        return;
      }

      const history = await this.captureHistoryEntry('Slide background image');
      await this.engine.setSlideBackgroundImage(this.currentSlide, image.bytes, image.mimeType);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) {
        await this.renderThumbnails();
        this.renderInspector();
      }
    } catch (error) {
      new Notice(`Could not set slide background: ${cleanError(error)}`);
    }
  }

  /**
   * Run an engine mutation against a single shape, threading it through the
   * shared dirty/history/re-render flow so undo/redo and persistence behave
   * like the other object operations.
   */
  private async applyShapeMutation(
    shapeIndex: number,
    historyLabel: string,
    action: string,
    mutate: (slideIndex: number, shapeIndex: number) => Promise<void>
  ): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable(action)) return;

    try {
      const history = await this.captureHistoryEntry(historyLabel);
      await mutate(this.currentSlide, shapeIndex);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) {
        this.selectShape(shapeIndex);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not ${action}: ${cleanError(error)}`);
    }
  }

  private selectShape(shapeIndex: number): void {
    if (!this.engine || !this.svgEl) return;

    this.selectedShapeIndex = shapeIndex;
    this.selectedShapeIndices = new Set([shapeIndex]);
    this.removeMultiSelectionBoxes();
    this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      shape.removeClass('native-powerpoint-shape-selected');
    });

    const selected = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!isSVGGElement(selected)) {
      this.selectedShapeIndex = null;
      this.selectedShapeIndices.clear();
      this.selectedTransform = null;
      this.removeSelectionOverlay();
      return;
    }

    selected.addClass('native-powerpoint-shape-selected');
    this.selectedTransform = cloneTransform(this.engine.getShapeTransform(selected));
    this.renderInspector();
    this.updateSelectionOverlay();
    this.updateObjectClipboardAvailability();
    this.updateTextToolbar();
  }

  private getTopLevelShapeIndices(): number[] {
    if (!this.svgEl) return [];
    const indices: number[] = [];
    this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      if (shape.parentElement?.closest('g[data-ooxml-shape-idx]')) return;
      const index = getShapeIndex(shape);
      if (index !== null) indices.push(index);
    });
    return indices;
  }

  private selectAllShapes(): void {
    const indices = this.getTopLevelShapeIndices();
    if (indices.length === 0) return;
    this.applyMultiSelection(indices);
  }

  private selectAllSlides(): void {
    if (!this.engine) return;
    const count = this.engine.slideCount;
    if (count === 0) return;
    this.selectedSlideIndices = new Set(Array.from({ length: count }, (_, index) => index));
    this.applySlideSelectionClasses();
  }

  private clearSlideSelection(): void {
    if (this.selectedSlideIndices.size === 0) return;
    this.selectedSlideIndices.clear();
    this.applySlideSelectionClasses();
  }

  private toggleSlideSelection(index: number): void {
    if (this.selectedSlideIndices.has(index)) {
      this.selectedSlideIndices.delete(index);
    } else {
      this.selectedSlideIndices.add(index);
    }
    this.applySlideSelectionClasses();
  }

  private selectSlideRange(anchor: number, index: number): void {
    const start = Math.min(anchor, index);
    const end = Math.max(anchor, index);
    this.selectedSlideIndices = new Set();
    for (let slide = start; slide <= end; slide += 1) {
      this.selectedSlideIndices.add(slide);
    }
    this.applySlideSelectionClasses();
  }

  private applySlideSelectionClasses(): void {
    this.thumbnailContainer?.querySelectorAll('.native-powerpoint-thumbnail').forEach((thumbnail, index) => {
      thumbnail.classList.toggle('is-selected', this.selectedSlideIndices.has(index));
    });
  }

  private applyMultiSelection(indices: number[]): void {
    if (!this.engine || !this.svgEl) return;

    const valid = indices.filter(
      (index) => this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${index}"]`) !== null
    );
    if (valid.length === 0) {
      this.clearSelection();
      return;
    }
    const [first] = valid;
    if (valid.length === 1 && first !== undefined) {
      this.selectShape(first);
      return;
    }

    this.selectedShapeIndex = null;
    this.selectedShapeIndices = new Set(valid);
    this.selectedTransform = null;
    this.applySelectionClasses();
    this.removeSelectionOverlay();
    this.renderInspector();
    this.updateSelectionOverlay();
    this.updateObjectClipboardAvailability();
  }

  private applySelectionClasses(): void {
    this.svgEl?.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      const index = getShapeIndex(shape);
      if (index !== null && this.selectedShapeIndices.has(index)) {
        shape.addClass('native-powerpoint-shape-selected');
      } else {
        shape.removeClass('native-powerpoint-shape-selected');
      }
    });
  }

  private clearSelection(options: { skipTextCommit?: boolean } = {}): void {
    if (!options.skipTextCommit) {
      this.commitActiveTextEditing();
    }
    this.selectedShapeIndex = null;
    this.selectedShapeIndices.clear();
    this.selectedTransform = null;
    this.svgEl?.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      shape.removeClass('native-powerpoint-shape-selected');
    });
    this.removeSelectionOverlay();
    this.removeMultiSelectionBoxes();
    this.clearSnapGuides();
    this.renderInspector();
    this.updateObjectClipboardAvailability();
    this.updateTextToolbar();
  }

  private renderSlideBackgroundControl(container: HTMLElement): void {
    if (!this.engine || this.engine.slideCount === 0) return;

    const section = container.createDiv({ cls: 'native-powerpoint-slide-background' });
    section.createDiv({ cls: 'native-powerpoint-inspector-subtitle', text: 'Slide background' });
    section.createDiv({
      cls: 'native-powerpoint-inspector-hint',
      text: 'Set the background fill color for the current slide.'
    });

    const currentColor = this.engine.getSlideBackgroundColor(this.currentSlide);
    const row = section.createDiv({ cls: 'native-powerpoint-background-row' });
    const colorInput = row.createEl('input', {
      type: 'color',
      cls: 'native-powerpoint-background-color',
      value: currentColor ? `#${currentColor}` : '#ffffff'
    });
    colorInput.disabled = !this.canEdit();

    const applyButton = row.createEl('button', {
      cls: 'native-powerpoint-inspector-button',
      text: 'Apply'
    });
    applyButton.disabled = !this.canEdit();
    applyButton.addEventListener('click', () => {
      void this.applySlideBackgroundColor(colorInput.value);
    });
  }

  private async applySlideBackgroundColor(hexColor: string): Promise<void> {
    if (!this.engine) return;
    if (!this.ensureEditable('change slide background')) return;

    try {
      const history = await this.captureHistoryEntry('Slide background');
      await this.engine.setSlideBackgroundColor(this.currentSlide, hexColor);
      this.recordHistoryEntry(history);
      this.markDirty();
      const rendered = await this.renderCurrentSlide(true);
      if (rendered) {
        await this.renderThumbnails();
        this.renderInspector();
      }
    } catch (error) {
      new Notice(`Could not change slide background: ${cleanError(error)}`);
    }
  }

  private renderInspector(): void {
    if (!this.inspectorEl) return;

    this.inspectorEl.empty();
    this.inspectorEl.createDiv({ cls: 'native-powerpoint-inspector-title', text: 'Inspector' });
    this.renderViewOnlyWarning(this.inspectorEl);
    this.renderFontFidelity(this.inspectorEl);
    this.renderSlideBackgroundControl(this.inspectorEl);

    if (!this.engine || this.selectedShapeIndex === null || !this.selectedTransform) {
      this.inspectorEl.createDiv({
        cls: 'native-powerpoint-inspector-empty',
        text: this.selectedShapeIndices.size > 1
          ? `${this.selectedShapeIndices.size} objects selected. Drag to move them together, or press Delete to remove them.`
          : 'Select a slide object to adjust its layout. Click text on the slide to edit it directly. Drag on empty space to select multiple objects.'
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
      if (target?.kind === 'shape-paragraph') {
        await this.engine.updateParagraphText(
          this.currentSlide,
          target.shapeIndex,
          target.paragraphIndex,
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
    this.clearWholeShapeInlineSelection();

    // Pointermove fires far faster than we can run the (expensive) SVG glyph
    // measurement, so coalesce to one selection update per animation frame using
    // the most recent pointer position. This keeps dragging smooth and ensures
    // the final position is always processed instead of dropped mid-flood.
    const flushDragFrame = () => {
      const drag = this.inlineSelectionDrag;
      if (!drag) return;
      drag.pendingFrame = null;
      this.clearBrowserTextSelection();
      this.extendInlineSelectionDrag(drag.pendingClientX, drag.pendingClientY);
    };

    const onPointerMove = (moveEvent: PointerEvent) => {
      if (moveEvent.pointerId !== event.pointerId) return;
      moveEvent.preventDefault();
      moveEvent.stopPropagation();
      const drag = this.inlineSelectionDrag;
      if (!drag) return;
      drag.pendingClientX = moveEvent.clientX;
      drag.pendingClientY = moveEvent.clientY;
      if (drag.pendingFrame === null) {
        drag.pendingFrame = window.requestAnimationFrame(flushDragFrame);
      }
    };
    const onPointerUp = (upEvent: PointerEvent) => {
      if (upEvent.pointerId !== event.pointerId) return;
      upEvent.preventDefault();
      upEvent.stopPropagation();
      this.clearBrowserTextSelection();
      this.extendInlineSelectionDrag(upEvent.clientX, upEvent.clientY);
      this.stopInlineSelectionDrag();
    };
    const cleanup = () => {
      if (this.inlineSelectionDrag?.pendingFrame !== null && this.inlineSelectionDrag) {
        window.cancelAnimationFrame(this.inlineSelectionDrag.pendingFrame);
      }
      activeDocument.removeEventListener('pointermove', onPointerMove, true);
      activeDocument.removeEventListener('pointerup', onPointerUp, true);
      activeDocument.removeEventListener('pointercancel', onPointerUp, true);
    };

    this.inlineSelectionDrag = {
      editor,
      element,
      anchorOffset,
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      isSelecting: false,
      pendingFrame: null,
      pendingClientX: event.clientX,
      pendingClientY: event.clientY,
      cleanup
    };
    activeDocument.addEventListener('pointermove', onPointerMove, true);
    activeDocument.addEventListener('pointerup', onPointerUp, true);
    activeDocument.addEventListener('pointercancel', onPointerUp, true);
  }

  private extendInlineSelectionDrag(clientX: number, clientY: number): void {
    const drag = this.inlineSelectionDrag;
    if (!drag || this.activeEditor !== drag.editor || this.activeEditorTarget !== drag.element) return;

    const box = this.getElementBox(drag.element);
    if (!box) return;

    if (!drag.isSelecting) {
      if (!this.hasInlineSelectionDragMoved(drag, clientX, clientY)) {
        drag.editor.setSelectionRange(drag.anchorOffset, drag.anchorOffset);
        this.rememberInlineCaretPlacement(drag.editor, drag.element, drag.anchorOffset);
        this.updateInlineCaret(drag.editor, drag.element);
        return;
      }
      drag.isSelecting = true;
      this.lastInlineCaretPlacement = null;
    }

    // Determine which paragraph in the text box the pointer is currently over so a
    // drag can extend the selection across paragraph boundaries (not just within
    // the paragraph the drag started in).
    const shape = this.getSelectedShapeElement();
    const paragraphs = shape ? this.getShapeTextParagraphs(shape) : [];
    const anchorIndex = paragraphs.indexOf(drag.element);
    const focusParagraph = anchorIndex >= 0
      ? this.getDragFocusParagraph(paragraphs, clientY, drag.element)
      : drag.element;
    const focusIndex = paragraphs.indexOf(focusParagraph);

    if (anchorIndex < 0 || focusIndex < 0 || focusParagraph === drag.element) {
      // Single-paragraph selection: keep using the editor's native selection model.
      this.clearWholeShapeInlineSelection();
      this.activeInlineCaretRow = this.getInlineCaretRowFromClientY(drag.element, clientY, box);
      const focusOffset = this.getInlineTextOffsetAtClientPoint(drag.element, drag.editor, clientX, clientY, box);
      const selectionStart = Math.min(drag.anchorOffset, focusOffset);
      const selectionEnd = Math.max(drag.anchorOffset, focusOffset);
      const direction = focusOffset < drag.anchorOffset ? 'backward' : 'forward';
      this.focusEditorWithoutCanvasScroll(drag.editor);
      drag.editor.setSelectionRange(selectionStart, selectionEnd, direction);
      this.resetInlineEditorScroll(drag.editor);
      this.updateInlineCaret(drag.editor, drag.element);
      return;
    }

    // Cross-paragraph selection.
    const focusBox = this.getElementBox(focusParagraph) ?? box;
    const focusOffset = this.getInlineTextOffsetAtClientPointForElement(focusParagraph, clientX, clientY, focusBox);
    this.renderCrossParagraphSelection(paragraphs, anchorIndex, drag.anchorOffset, focusIndex, focusOffset);
  }

  private getDragFocusParagraph(
    paragraphs: (SVGTextElement | SVGTSpanElement)[],
    clientY: number,
    fallback: SVGTextElement | SVGTSpanElement
  ): SVGTextElement | SVGTSpanElement {
    const paneRect = this.canvasPane?.getBoundingClientRect();
    const localY = paneRect ? clientY - paneRect.top + (this.canvasPane?.scrollTop ?? 0) : clientY;

    let best = fallback;
    let bestDistance = Number.POSITIVE_INFINITY;
    for (const paragraph of paragraphs) {
      const box = this.getElementBox(paragraph);
      if (!box) continue;
      if (localY >= box.top && localY <= box.top + box.height) {
        return paragraph;
      }
      const center = box.top + box.height / 2;
      const distance = Math.abs(localY - center);
      if (distance < bestDistance) {
        bestDistance = distance;
        best = paragraph;
      }
    }
    return best;
  }

  private renderCrossParagraphSelection(
    paragraphs: (SVGTextElement | SVGTSpanElement)[],
    anchorIndex: number,
    anchorOffset: number,
    focusIndex: number,
    focusOffset: number
  ): void {
    let startIndex = anchorIndex;
    let startOffset = anchorOffset;
    let endIndex = focusIndex;
    let endOffset = focusOffset;
    if (focusIndex < anchorIndex || (focusIndex === anchorIndex && focusOffset < anchorOffset)) {
      startIndex = focusIndex;
      startOffset = focusOffset;
      endIndex = anchorIndex;
      endOffset = anchorOffset;
    }

    this.removeInlineSelection();
    this.activeInlineCaret?.addClass('native-powerpoint-inline-caret-hidden');

    const textParts: string[] = [];
    for (let index = startIndex; index <= endIndex; index++) {
      const paragraph = paragraphs[index];
      if (!paragraph) continue;
      const total = this.getLeafCharInfo(paragraph).total;
      const rangeStart = index === startIndex ? Math.max(0, Math.min(total, startOffset)) : 0;
      const rangeEnd = index === endIndex ? Math.max(0, Math.min(total, endOffset)) : total;
      if (rangeEnd > rangeStart) {
        this.renderInlineSelectionRects(paragraph, rangeStart, rangeEnd);
      }
      const paragraphText = paragraph.textContent ?? '';
      textParts.push(paragraphText.slice(rangeStart, rangeEnd));
    }

    // Reuse the compound-selection copy buffer so Ctrl/Cmd+C copies the full range,
    // but render the rects directly here rather than via the whole-shape path.
    this.inlineWholeShapeSelected = false;
    this.inlineWholeShapeSelection = textParts.join('\n');
  }

  private hasInlineSelectionDragMoved(drag: InlineSelectionDrag, clientX: number, clientY: number): boolean {
    const dx = clientX - drag.startClientX;
    const dy = clientY - drag.startClientY;
    return Math.hypot(dx, dy) >= 4;
  }

  private stopInlineSelectionDrag(): void {
    this.inlineSelectionDrag?.cleanup();
    this.inlineSelectionDrag = null;
  }

  private clearBrowserTextSelection(): void {
    activeDocument.getSelection()?.removeAllRanges();
  }

  private commitActiveTextEditing(): void {
    if (!this.activeEditor) return;

    const commit = this.activeEditorCommit;
    if (commit) {
      void commit();
      return;
    }

    this.removeActiveEditor();
  }

  private handleOutsideSlidePointerDown = (event: PointerEvent): void => {
    if (!this.activeEditor) return;

    const target = isNode(event.target) ? event.target : null;
    if (isElement(target) && target.closest('.native-powerpoint-text-toolbar, .native-powerpoint-toolbar-popover')) return;
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

  private handleCanvasContextMenu = (event: MouseEvent): void => {
    event.preventDefault();

    const shapeIndex = this.getTopLevelShapeIndexFromEvent(event);
    if (shapeIndex !== null) {
      this.showObjectContextMenu(event, shapeIndex);
      return;
    }

    this.showCanvasContextMenu(event);
  };

  private showCanvasContextMenu(event: MouseEvent): void {
    const menu = this.createNativeMenu();

    menu.addItem((item) => {
      item
        .setTitle('Paste')
        .setIcon('clipboard-paste')
        .onClick(() => void this.pasteCopiedShape());
      if (!this.objectClipboard) item.setDisabled(true);
    });

    menu.addItem((item) => {
      const shapeCount = this.getTopLevelShapeIndices().length;
      item
        .setTitle('Select all')
        .setIcon('box-select')
        .onClick(() => this.selectAllShapes());
      if (shapeCount === 0) item.setDisabled(true);
    });

    menu.addItem((item) => {
      item
        .setTitle('New text box')
        .setIcon('type')
        .onClick(() => {
          if (this.ensureEditable('add text box')) void this.addTextBox();
        });
    });

    menu.showAtMouseEvent(event);
  }

  /**
   * Resolve the top-level shape index for a right-click, mirroring the
   * top-level ancestor logic in {@link getTopLevelShapeIndices}: walk up out of
   * any nested group shapes to the outermost `g[data-ooxml-shape-idx]`.
   */
  private getTopLevelShapeIndexFromEvent(event: MouseEvent): number | null {
    const target = event.target;
    if (!(target instanceof Element)) return null;

    let shape = target.closest('g[data-ooxml-shape-idx]');
    if (!shape) return null;

    let ancestor = shape.parentElement?.closest('g[data-ooxml-shape-idx]') ?? null;
    while (ancestor) {
      shape = ancestor;
      ancestor = shape.parentElement?.closest('g[data-ooxml-shape-idx]') ?? null;
    }
    return getShapeIndex(shape);
  }

  private getObjectKind(shapeIndex: number, shapeEl: SVGGElement): 'image' | 'text' | 'generic' {
    const type = shapeEl.getAttribute('data-ooxml-shape-type');
    if (type === 'table' || type === 'chart' || type === 'group') {
      return 'generic';
    }
    if (this.engine?.isImageShape(this.currentSlide, shapeIndex) || shapeEl.querySelector('image')) {
      return 'image';
    }
    if (shapeEl.querySelector('text')) {
      return 'text';
    }
    return 'generic';
  }

  /**
   * Build the right-click menu for a slide object. The right-clicked shape is
   * selected first (unless it is already part of the current multi-selection)
   * so every action operates on it.
   */
  private showObjectContextMenu(event: MouseEvent, shapeIndex: number): void {
    if (!this.selectedShapeIndices.has(shapeIndex)) {
      this.selectShape(shapeIndex);
    }

    const shapeEl = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    const kind = isSVGGElement(shapeEl) ? this.getObjectKind(shapeIndex, shapeEl) : 'generic';
    const canEdit = this.canEdit();
    const menu = this.createNativeMenu();

    menu.addItem((item) => {
      item.setTitle('Cut').setIcon('scissors').onClick(() => void this.cutSelectedShape());
      if (!canEdit) item.setDisabled(true);
    });
    menu.addItem((item) =>
      item.setTitle('Copy').setIcon('copy').onClick(() => void this.copySelectedShape())
    );
    menu.addItem((item) => {
      item.setTitle('Paste').setIcon('clipboard-paste').onClick(() => void this.pasteCopiedShape());
      if (!canEdit || !this.objectClipboard) item.setDisabled(true);
    });
    menu.addItem((item) => {
      item
        .setTitle('Paste without formatting')
        .setIcon('clipboard-type')
        .onClick(() => void this.pasteWithoutFormatting());
      if (!canEdit || (!this.activeEditor && !this.objectClipboard)) item.setDisabled(true);
    });
    menu.addItem((item) => {
      item.setTitle('Delete').setIcon('trash-2').onClick(() => void this.deleteSelectedShape());
      if (!canEdit) item.setDisabled(true);
    });

    menu.addSeparator();
    this.addOrderSubsection(menu, canEdit);
    this.addRotateSubsection(menu, canEdit, kind === 'image');
    this.addCenterSubsection(menu, canEdit);

    if (kind === 'image') {
      menu.addSeparator();
      menu.addItem((item) => {
        item.setTitle('Crop image…').setIcon('crop').onClick(() => this.openImageCropDialog(shapeIndex));
        if (!canEdit) item.setDisabled(true);
      });
      this.addReplaceImageSubsection(menu, canEdit, shapeIndex);
      menu.addItem((item) => {
        item
          .setTitle('Reset image')
          .setIcon('rotate-ccw')
          .onClick(() => void this.resetSelectedImage(shapeIndex));
        if (!canEdit) item.setDisabled(true);
      });
      menu.addItem((item) => {
        item
          .setTitle('Set as background')
          .setIcon('image')
          .onClick(() => void this.setSelectedImageAsBackground(shapeIndex));
        if (!canEdit) item.setDisabled(true);
      });
    }

    menu.showAtMouseEvent(event);
  }

  /**
   * Add a labelled group of actions to a menu. Uses a real Obsidian side
   * submenu when {@link MenuItem.setSubmenu} is available (Obsidian 1.4.5+) and
   * falls back to flat, prefixed items with a section label otherwise.
   */
  private addMenuSubsection(
    menu: Menu,
    title: string,
    icon: string,
    populate: (add: (label: string, icon: string, onClick: () => void, disabled?: boolean) => void) => void
  ): void {
    const holder: { submenu: Menu | null } = { submenu: null };
    menu.addItem((item) => {
      item.setTitle(title).setIcon(icon);
      const withSubmenu = item as unknown as { setSubmenu?: () => Menu | undefined };
      const created = typeof withSubmenu.setSubmenu === 'function' ? withSubmenu.setSubmenu() : undefined;
      if (created) {
        holder.submenu = created;
        (created as unknown as { dom?: HTMLElement }).dom?.addClass('native-powerpoint-light-surface');
      } else {
        item.setIsLabel(true);
      }
    });

    const target = holder.submenu;
    if (target) {
      populate((label, itemIcon, onClick, disabled) =>
        target.addItem((item) => {
          item.setTitle(label).setIcon(itemIcon).onClick(onClick);
          if (disabled) item.setDisabled(true);
        })
      );
    } else {
      populate((label, itemIcon, onClick, disabled) =>
        menu.addItem((item) => {
          item.setTitle(`${title}: ${label}`).setIcon(itemIcon).onClick(onClick);
          if (disabled) item.setDisabled(true);
        })
      );
    }
  }

  private addOrderSubsection(menu: Menu, canEdit: boolean): void {
    this.addMenuSubsection(menu, 'Order', 'layers', (add) => {
      add('Bring to front', 'bring-to-front', () => void this.reorderSelection('front'), !canEdit);
      add('Bring forward', 'arrow-up', () => void this.reorderSelection('forward'), !canEdit);
      add('Send backward', 'arrow-down', () => void this.reorderSelection('backward'), !canEdit);
      add('Send to back', 'send-to-back', () => void this.reorderSelection('back'), !canEdit);
    });
  }

  private addRotateSubsection(menu: Menu, canEdit: boolean, allowFlip: boolean): void {
    this.addMenuSubsection(menu, 'Rotate', 'rotate-cw', (add) => {
      add('Rotate right 90°', 'rotate-cw', () => void this.rotateSelectedShape(90), !canEdit);
      add('Rotate left 90°', 'rotate-ccw', () => void this.rotateSelectedShape(-90), !canEdit);
      if (allowFlip) {
        add('Flip horizontal', 'flip-horizontal', () => void this.flipSelectedShape('horizontal'), !canEdit);
        add('Flip vertical', 'flip-vertical', () => void this.flipSelectedShape('vertical'), !canEdit);
      }
    });
  }

  private addCenterSubsection(menu: Menu, canEdit: boolean): void {
    this.addMenuSubsection(menu, 'Center on page', 'align-center-horizontal', (add) => {
      add('Center horizontally', 'align-center-vertical', () => void this.centerSelectedOnPage('horizontal'), !canEdit);
      add('Center vertically', 'align-center-horizontal', () => void this.centerSelectedOnPage('vertical'), !canEdit);
    });
  }

  private addReplaceImageSubsection(menu: Menu, canEdit: boolean, shapeIndex: number): void {
    this.addMenuSubsection(menu, 'Replace image', 'image-plus', (add) => {
      add('From vault…', 'folder', () => this.openReplaceImageVaultPicker(shapeIndex), !canEdit);
      add('Upload file…', 'upload', () => {
        if (!this.ensureEditable('replace image')) return;
        this.pendingReplaceShapeIndex = shapeIndex;
        this.replaceImageFileInput?.click();
      }, !canEdit);
    });
  }

  private handleCanvasPanePointerDown = (event: PointerEvent): void => {
    this.lastInteractionRegion = 'canvas';
    this.clearSlideSelection();
    if (event.button !== 0) return;
    this.suppressNextClick = false;

    const target = isNode(event.target) ? event.target : null;
    if (!this.isCanvasPaneBackgroundTarget(target)) return;

    event.preventDefault();
    event.stopPropagation();
    this.clearBrowserTextSelection();
    if (this.activeEditor) {
      this.commitActiveEditorFromOutside(true);
      return;
    }

    this.removeInlineSelection();
    this.lastInlineCaretPlacement = null;
    this.beginMarquee(event, event.shiftKey || event.ctrlKey || event.metaKey);
  };

  private isCanvasPaneBackgroundTarget(target: Node | null): boolean {
    if (!target || !this.canvasPane?.contains(target)) return false;
    if (this.slideSurface?.contains(target)) return false;
    if (this.activeEditor?.contains(target)) return false;

    const element = isElement(target) ? target : target.parentElement;
    if (element?.closest('.native-powerpoint-selection-box')) return false;
    // The contextual text toolbar lives inside the canvas pane, so clicks on it
    // must not be treated as background clicks (which would commit/clear the
    // editor and hide the toolbar mid-interaction).
    if (element?.closest('.native-powerpoint-text-toolbar, .native-powerpoint-toolbar-popover')) return false;

    return true;
  }

  private commitActiveEditorFromOutside(clearSelectionAfterCommit: boolean): void {
    this.clearBrowserTextSelection();
    const commit = this.activeEditorCommit;
    if (commit) {
      void commit().finally(() => {
        this.clearBrowserTextSelection();
        if (clearSelectionAfterCommit) {
          this.clearSelection({ skipTextCommit: true });
        }
      });
      return;
    }

    this.removeActiveEditor();
    if (clearSelectionAfterCommit) {
      this.clearSelection({ skipTextCommit: true });
    }
  }

  private startTextEditor(target: TextEditTarget | null = null, clientX?: number, clientY?: number): void {
    if (!this.canvasPane || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('edit text')) return;

    const selected = this.getSelectedShapeElement();
    if (!selected) return;
    if (!target) {
      target = this.getTextEditTargetFromSelectedShape();
    }
    if (!target && selected.closest(GENERATED_GRID_SELECTOR)) {
      this.showGeneratedTextNotice();
      return;
    }
    if (!target) return;

    if (this.activeEditor && this.activeEditorTarget === target.element) {
      const currentBox = this.getElementBox(target.element);
      if (currentBox) {
        this.placeInlineCaret(this.activeEditor, target.element, clientX, clientY, currentBox);
        this.focusEditorWithoutCanvasScroll(this.activeEditor);
        this.resetInlineEditorScroll(this.activeEditor);
      }
      return;
    }

    if (this.activeEditor) {
      this.commitActiveTextEditing();
    }
    if (this.activeEditor) {
      this.removeActiveEditor();
    }
    const box = this.getElementBox(target.element);
    if (!box) return;

    const editor = this.canvasPane.createEl('textarea', {
      cls: 'native-powerpoint-inline-editor is-text-run',
      attr: { 'aria-label': 'Edit selected text' }
    });
    const initialText = target.text;
    const initialRunTexts = target.kind === 'shape-paragraph'
      ? target.runElements.map((run) => run.textContent || '')
      : [];
    editor.value = initialText;

    const styleElement = target.kind === 'shape-paragraph' && target.runElements[0]
      ? target.runElements[0]
      : target.element;
    const style = window.getComputedStyle(styleElement);
    target.element.classList.add('native-powerpoint-text-editing');
    this.activeEditorTarget = target.element;
    this.activeShapeTextTarget = target.kind === 'shape-paragraph' ? target : null;
    this.activeTextStyleTarget = target.kind === 'shape-paragraph'
      ? this.getPrimaryStyleRunTarget(target)
      : null;
    this.slideSurface?.addClass('is-inline-text-editing');
    editor.setCssProps({
      color: style.fill,
      fontFamily: style.fontFamily,
      fontSize: `${this.getScreenFontSize(styleElement)}px`,
      fontStyle: style.fontStyle,
      fontWeight: style.fontWeight,
      lineHeight: '1.1',
      textAlign: this.getInlineTextAlignment(style.textAnchor)
    });
    this.positionTextRunEditor(editor, box);
    this.activeEditor = editor;

    this.removeSelectionOverlay();
    this.activeInlineCaret = activeDocument.createElementNS('http://www.w3.org/2000/svg', 'line');
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
      this.clearWholeShapeInlineSelection();
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
        this.clearWholeShapeInlineSelection();
        this.syncShapeParagraphPreview(target, editor.value);
        const nextBox = this.getElementBox(target.element);
        if (nextBox) {
          this.positionTextRunEditor(editor, nextBox);
        }
        updateCaret();
      }
    });
    editor.addEventListener('copy', (event) => {
      if (this.inlineWholeShapeSelection !== null) {
        event.preventDefault();
        event.clipboardData?.setData('text/plain', this.inlineWholeShapeSelection);
      }
    });
    editor.addEventListener('click', updateCaret);
    editor.addEventListener('keyup', updateCaret);
    editor.addEventListener('mouseup', updateCaret);
    editor.addEventListener('select', updateCaret);
    editor.addEventListener('keydown', (event) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'a') {
        event.preventDefault();
        this.selectAllInlineText(editor, target.element);
        return;
      }
      if (
        event.shiftKey
        && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)
      ) {
        this.clearWholeShapeInlineSelection();
        this.lastInlineCaretPlacement = null;
      } else if (
        this.inlineWholeShapeSelection !== null
        && !event.metaKey
        && !event.ctrlKey
        && !event.altKey
        && (event.key === 'Enter' || event.key.length === 1 || ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key))
      ) {
        this.clearWholeShapeInlineSelection();
      }
      if (this.handleInlineDeleteKey(event, editor, target.element)) return;
      queueCaretUpdate();
    });

    const commit = async () => {
      if (this.activeEditor !== editor) return;
      this.removeActiveEditor(editor);
      await this.applyTextValue(editor.value, target);
    };
    this.activeEditorCommit = commit;

    editor.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
        event.preventDefault();
        void commit();
      } else if (event.key === 'Escape') {
        if (target.kind === 'shape-paragraph') {
          target.runElements.forEach((run, index) => {
            run.textContent = initialRunTexts[index] ?? '';
          });
        } else {
          target.element.textContent = initialText;
        }
        editor.value = initialText;
        this.removeActiveEditor(editor);
      }
    });
    editor.addEventListener('blur', (event) => {
      // A bare modifier press (e.g. ⌘/Ctrl arming the app menu) or the window
      // losing key focus blurs the textarea with no related target. Treat that
      // as transient: keep the editor open and restore focus so an in-progress
      // text selection isn't discarded. Genuine clicks outside the editor are
      // already committed by the document-level pointerdown handler before this
      // fires, so by then activeEditor no longer matches and we fall through to
      // commit (which early-returns). Focus moving to another real element
      // (e.g. Tab) still commits.
      const next = event.relatedTarget;
      const stayEditing = next === null
        || (isElement(next)
          && next.closest('.native-powerpoint-text-toolbar, .native-powerpoint-toolbar-popover') !== null);
      if (stayEditing && this.activeEditor === editor && editor.isConnected) {
        this.focusEditorWithoutCanvasScroll(editor);
        updateCaret();
        return;
      }
      void commit();
    });
    this.focusEditorWithoutCanvasScroll(editor);
    this.placeInlineCaret(editor, target.element, clientX, clientY, box);
    this.updateTextToolbar();
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
    this.inlineWholeShapeSelection = null;
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
    this.activeShapeTextTarget = null;
    this.activeTextStyleTarget = null;
    this.slideSurface?.removeClass('is-inline-text-editing');
    if (this.svgEl?.isConnected) {
      this.updateSelectionOverlay();
    } else {
      this.updateTextToolbar();
    }
  }

  // --- Contextual text formatting toolbar (Google Slides–style) ------------

  private shapeHasEditableText(shape: Element): boolean {
    return shape.querySelector('tspan[data-ooxml-run-idx]') !== null
      && shape.closest(GENERATED_GRID_SELECTOR) === null;
  }

  private getTextStyleContext(): TextStyleContext | null {
    if (!this.engine || !this.canEdit() || !this.svgEl) return null;

    if (this.activeTextStyleTarget && this.activeEditor) {
      const target = this.activeTextStyleTarget;
      const anchor = this.getElementBox(this.activeEditorTarget ?? target.element);
      if (!anchor) return null;
      return {
        shapeIndex: target.shapeIndex,
        run: { paragraphIndex: target.paragraphIndex, runIndex: target.runIndex },
        anchor
      };
    }

    // Keep the formatting toolbar visible while a single text shape is selected,
    // even when no inline editor is active (e.g. after the editor is flushed
    // because the user clicked into the toolbar's font-size box).
    if (this.selectedShapeIndex !== null && this.selectedShapeIndices.size <= 1) {
      const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${this.selectedShapeIndex}"]`);
      if (shape && this.shapeHasEditableText(shape)) {
        const anchor = this.getSelectedBox();
        if (anchor) {
          return { shapeIndex: this.selectedShapeIndex, run: null, anchor };
        }
      }
    }

    return null;
  }

  private getFirstRunTarget(shapeIndex: number): RunTarget | null {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    const run = shape?.querySelector('tspan[data-ooxml-run-idx]') ?? null;
    if (!run) return null;

    const paragraph = run.closest('tspan[data-ooxml-para-idx]');
    const paragraphIndex = Number(paragraph?.getAttribute('data-ooxml-para-idx'));
    const runIndex = Number(run.getAttribute('data-ooxml-run-idx'));
    if (!Number.isFinite(paragraphIndex) || !Number.isFinite(runIndex)) return null;
    return { paragraphIndex, runIndex };
  }

  private buildParagraphEditTarget(shapeIndex: number, paragraphIndex: number): ShapeTextEditTarget | null {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!shape) return null;

    // Seed from a run tspan inside the target paragraph so getTextEditTarget
    // resolves the correct paragraph (passing the paragraph tspan alone makes it
    // fall back to the first paragraph in the shape).
    const seed = shape.querySelector(`tspan[data-ooxml-para-idx="${paragraphIndex}"] tspan[data-ooxml-run-idx]`)
      ?? shape.querySelector(`tspan[data-ooxml-para-idx="${paragraphIndex}"]`)
      ?? shape.querySelector('text');
    const target = this.getTextEditTarget(seed);
    return target?.kind === 'shape-paragraph' ? target : null;
  }

  private buildRunTarget(shapeIndex: number, paragraphIndex: number, runIndex: number): ShapeTextEditTarget | null {
    const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    if (!shape) return null;

    const run = Array.from(shape.querySelectorAll('tspan[data-ooxml-run-idx]')).find((candidate) => {
      const paragraph = candidate.closest('tspan[data-ooxml-para-idx]');
      return Number(candidate.getAttribute('data-ooxml-run-idx')) === runIndex
        && Number(paragraph?.getAttribute('data-ooxml-para-idx')) === paragraphIndex;
    });
    if (!isSVGTSpanElement(run)) return null;

    const paragraph = run.closest('tspan[data-ooxml-para-idx]');
    const geometryElement = paragraph && isSVGTSpanElement(paragraph) ? paragraph : run;
    const runElements = paragraph
      ? this.collectParagraphRuns(paragraph)
      : [run];
    const text = runElements.map((candidate) => candidate.textContent || '').join('');

    return {
      kind: 'shape-paragraph',
      shapeIndex,
      paragraphIndex,
      runIndex,
      text,
      element: geometryElement,
      runElements
    };
  }

  private updateTextToolbar(): void {
    const context = this.getTextStyleContext();
    if (!context) {
      this.hideTextToolbar();
      return;
    }

    const controls = this.ensureTextToolbar();
    if (!controls || !this.textToolbarEl) return;

    // Position the toolbar only when it first appears for a shape. Subsequent
    // updates (e.g. flushing the inline editor after clicking the font-size box)
    // change the anchor from the caret line to the whole-shape box, which would
    // make the toolbar jump; keeping the spawn position avoids that.
    const wasVisible = this.textToolbarEl.hasClass('is-visible');
    const shapeChanged = this.textToolbarShapeIndex !== context.shapeIndex;
    this.textToolbarEl.addClass('is-visible');
    if (!wasVisible || shapeChanged) {
      this.positionTextToolbar(context.anchor);
      this.textToolbarShapeIndex = context.shapeIndex;
    }
    this.reflectTextToolbarState(context);
  }

  private hideTextToolbar(): void {
    this.textToolbarEl?.removeClass('is-visible');
    this.textToolbarShapeIndex = null;
    this.closeToolbarPopover();
    this.currentRunStyle = null;
  }

  private positionTextToolbar(anchor: { left: number; top: number; width: number; height: number }): void {
    const toolbar = this.textToolbarEl;
    if (!toolbar || !this.canvasPane) return;

    const toolbarHeight = toolbar.offsetHeight || 40;
    const gap = 8;
    let top = anchor.top - toolbarHeight - gap;
    if (top < this.canvasPane.scrollTop + 4) {
      top = anchor.top + anchor.height + gap;
    }

    const maxLeft = Math.max(0, this.canvasPane.scrollWidth - (toolbar.offsetWidth || 0) - 4);
    const left = Math.min(Math.max(anchor.left, 4), maxLeft);
    toolbar.setCssProps({ left: `${left}px`, top: `${Math.max(0, top)}px` });
  }

  private reflectTextToolbarState(context: TextStyleContext): void {
    const controls = this.textToolbarControls;
    if (!controls || !this.engine) return;

    const runTarget = context.run ?? this.getFirstRunTarget(context.shapeIndex);
    const style = runTarget
      ? this.engine.getRunStyle(this.currentSlide, context.shapeIndex, runTarget.paragraphIndex, runTarget.runIndex)
      : null;
    this.currentRunStyle = style;

    controls.bold.toggleClass('is-active', Boolean(style?.bold));
    controls.italic.toggleClass('is-active', Boolean(style?.italic));
    controls.underline.toggleClass('is-active', Boolean(style?.underline));
    controls.fontLabel.setText(style?.fontFamily ?? this.getEffectiveFontFamily(context) ?? 'Font');

    if (activeDocument.activeElement !== controls.fontSizeInput) {
      const sizePt = style?.fontSizePt ?? this.getEffectiveFontSizePt(context);
      controls.fontSizeInput.value = sizePt ? String(Math.round(sizePt)) : '';
    }

    if (style?.color) {
      this.textColorValue = style.color;
    }
    if (style?.highlight) {
      this.textHighlightValue = style.highlight;
    }
    controls.textColorBar.style.setProperty('--np-swatch-color', `#${style?.color ?? this.textColorValue}`);
    controls.highlightBar.style.setProperty('--np-swatch-color', style?.highlight ? `#${style.highlight}` : 'transparent');

    const alignment = style?.alignment ?? 'l';
    for (const align of ['l', 'ctr', 'r', 'just'] as ParagraphAlignment[]) {
      controls.alignButtons[align].toggleClass('is-active', alignment === align);
    }
  }

  // EMU per SVG user unit, used to convert rendered font sizes back to points.
  private getSvgEmuPerUnit(): number | null {
    const scale = Number(this.svgEl?.getAttribute('data-ooxml-scale'));
    if (Number.isFinite(scale) && scale > 0) return scale;

    const cx = Number(this.svgEl?.getAttribute('data-ooxml-slide-cx'));
    const width = this.svgEl ? Number.parseFloat(this.svgEl.getAttribute('width') || '') : Number.NaN;
    if (Number.isFinite(cx) && Number.isFinite(width) && width > 0) return cx / width;

    return null;
  }

  /**
   * Detect the effective font size (in points) actually rendered for the
   * relevant runs, so the toolbar can show a value even when the size is
   * inherited from the theme/placeholder rather than authored on each run.
   * Returns null when sizes are mixed or cannot be determined.
   */
  private getRelevantTextRuns(context: TextStyleContext): SVGTSpanElement[] {
    if (!this.svgEl) return [];
    const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${context.shapeIndex}"]`);
    if (!shape) return [];

    const allRuns = Array.from(shape.querySelectorAll('tspan[data-ooxml-run-idx]')).filter(isSVGTSpanElement);
    const targetRun = context.run;
    if (!targetRun) return allRuns;

    return allRuns.filter((run) => {
      const para = run.closest('tspan[data-ooxml-para-idx]');
      return Number(run.getAttribute('data-ooxml-run-idx')) === targetRun.runIndex
        && Number(para?.getAttribute('data-ooxml-para-idx')) === targetRun.paragraphIndex;
    });
  }

  private getEffectiveFontSizePt(context: TextStyleContext): number | null {
    const emuPerUnit = this.getSvgEmuPerUnit();
    if (!emuPerUnit) return null;

    const runs = this.getRelevantTextRuns(context);
    if (runs.length === 0) return null;

    const EMU_PER_POINT = 12700;
    let detected: number | null = null;
    for (const run of runs) {
      if ((run.textContent || '').length === 0) continue;
      const userUnits = Number.parseFloat(window.getComputedStyle(run).fontSize);
      if (!Number.isFinite(userUnits) || userUnits <= 0) continue;
      const rounded = Math.round((userUnits * emuPerUnit) / EMU_PER_POINT);
      if (detected === null) {
        detected = rounded;
      } else if (detected !== rounded) {
        return null;
      }
    }
    return detected;
  }

  /**
   * Detect the effective font family actually rendered for the relevant runs,
   * so the toolbar can show a value when the face is inherited from the theme
   * or placeholder rather than authored on each run. Returns null when families
   * are mixed or cannot be determined.
   */
  private getEffectiveFontFamily(context: TextStyleContext): string | null {
    const runs = this.getRelevantTextRuns(context);
    if (runs.length === 0) return null;

    let detected: string | null = null;
    for (const run of runs) {
      if ((run.textContent || '').length === 0) continue;
      const family = parsePrimaryFontFamily(window.getComputedStyle(run).fontFamily);
      if (!family) continue;
      if (detected === null) {
        detected = family;
      } else if (detected !== family) {
        return null;
      }
    }
    return detected;
  }

  private ensureTextToolbar(): TextToolbarControls | null {
    if (this.textToolbarControls && this.textToolbarEl?.isConnected) {
      return this.textToolbarControls;
    }
    if (!this.canvasPane) return null;

    this.textToolbarEl?.remove();
    const toolbar = this.canvasPane.createDiv({ cls: 'native-powerpoint-text-toolbar' });
    toolbar.setAttribute('role', 'toolbar');
    toolbar.setAttribute('aria-label', 'Text formatting');
    toolbar.addEventListener('pointerdown', (event) => event.stopPropagation());

    const styleGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const bold = this.createTextToolbarButton(styleGroup, 'bold', 'Bold', () => this.toggleRunFlag('bold'));
    const italic = this.createTextToolbarButton(styleGroup, 'italic', 'Italic', () => this.toggleRunFlag('italic'));
    const underline = this.createTextToolbarButton(styleGroup, 'underline', 'Underline', () => this.toggleRunFlag('underline'));

    const fontGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const fontButton = fontGroup.createEl('button', {
      cls: 'native-powerpoint-text-toolbar-font',
      attr: { 'aria-label': 'Font family' }
    });
    const fontLabel = fontButton.createSpan({ cls: 'native-powerpoint-text-toolbar-font-label', text: 'Font' });
    setIcon(fontButton.createSpan({ cls: 'native-powerpoint-text-toolbar-caret' }), 'chevron-down');
    this.bindToolbarButton(fontButton, () => this.openFontMenu(fontButton));

    const sizeGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    this.createTextToolbarButton(sizeGroup, 'minus', 'Decrease font size', () => this.stepFontSize(-1));
    const fontSizeInput = sizeGroup.createEl('input', {
      cls: 'native-powerpoint-text-toolbar-size',
      type: 'number',
      attr: {
        'aria-label': 'Font size',
        min: String(TEXT_TOOLBAR_MIN_FONT_SIZE),
        max: String(TEXT_TOOLBAR_MAX_FONT_SIZE)
      }
    });
    fontSizeInput.addEventListener('pointerdown', () => this.flushActiveEditor(), true);
    fontSizeInput.addEventListener('change', () => this.commitFontSizeInput());
    fontSizeInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        this.commitFontSizeInput();
      }
    });
    this.createTextToolbarButton(sizeGroup, 'plus', 'Increase font size', () => this.stepFontSize(1));

    const colorGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const textColorButton = this.createTextToolbarSwatchButton(colorGroup, 'baseline', 'Text color');
    const textColorBar = textColorButton.createDiv({ cls: 'native-powerpoint-text-toolbar-swatch-bar' });
    this.bindToolbarButton(textColorButton, () =>
      this.openColorPopover(textColorButton, this.textColorValue, false, (color) =>
        this.applyRunStyle({ color })));

    const highlightButton = this.createTextToolbarSwatchButton(colorGroup, 'highlighter', 'Highlight color');
    const highlightBar = highlightButton.createDiv({ cls: 'native-powerpoint-text-toolbar-swatch-bar' });
    this.bindToolbarButton(highlightButton, () =>
      this.openColorPopover(highlightButton, this.textHighlightValue, true, (color) =>
        this.applyRunStyle({ highlight: color })));

    const alignGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const alignButtons: Record<ParagraphAlignment, HTMLButtonElement> = {
      l: this.createTextToolbarButton(alignGroup, 'align-left', 'Align left', () => this.applyAlignment('l')),
      ctr: this.createTextToolbarButton(alignGroup, 'align-center', 'Align center', () => this.applyAlignment('ctr')),
      r: this.createTextToolbarButton(alignGroup, 'align-right', 'Align right', () => this.applyAlignment('r')),
      just: this.createTextToolbarButton(alignGroup, 'align-justify', 'Justify', () => this.applyAlignment('just'))
    };

    this.textToolbarEl = toolbar;
    this.textToolbarControls = {
      bold,
      italic,
      underline,
      fontLabel,
      fontSizeInput,
      textColorBar,
      highlightBar,
      alignButtons
    };
    return this.textToolbarControls;
  }

  private createTextToolbarButton(
    container: HTMLElement,
    icon: string,
    label: string,
    action: () => void
  ): HTMLButtonElement {
    const button = container.createEl('button', {
      cls: 'native-powerpoint-toolbar-btn native-powerpoint-text-toolbar-btn',
      attr: { 'aria-label': label }
    });
    setIcon(button, icon);
    this.bindToolbarButton(button, action);
    return button;
  }

  private createTextToolbarSwatchButton(container: HTMLElement, icon: string, label: string): HTMLButtonElement {
    const button = container.createEl('button', {
      cls: 'native-powerpoint-toolbar-btn native-powerpoint-text-toolbar-btn native-powerpoint-text-toolbar-swatch',
      attr: { 'aria-label': label }
    });
    setIcon(button.createSpan({ cls: 'native-powerpoint-text-toolbar-swatch-icon' }), icon);
    return button;
  }

  private bindToolbarButton(button: HTMLElement, action: () => void): void {
    button.addEventListener('pointerdown', (event) => {
      event.preventDefault();
      event.stopPropagation();
      action();
    });
  }

  private toggleRunFlag(flag: 'bold' | 'italic' | 'underline'): void {
    const editor = this.activeEditor;
    const target = this.activeTextStyleTarget;
    if (editor && target && this.engine) {
      const start = Math.min(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      const end = Math.max(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      if (start < end) {
        const next = !this.engine.isRangeStyled(
          this.currentSlide,
          target.shapeIndex,
          target.paragraphIndex,
          start,
          end,
          flag
        );
        this.applyRunStyle({ [flag]: next });
        return;
      }
    }

    const current = this.currentRunStyle?.[flag] ?? false;
    this.applyRunStyle({ [flag]: !current });
  }

  private stepFontSize(delta: number): void {
    const inputValue = Number(this.textToolbarControls?.fontSizeInput?.value);
    const current = this.currentRunStyle?.fontSizePt
      ?? (Number.isFinite(inputValue) && inputValue > 0 ? inputValue : 18);
    const next = Math.min(TEXT_TOOLBAR_MAX_FONT_SIZE, Math.max(TEXT_TOOLBAR_MIN_FONT_SIZE, Math.round(current) + delta));
    this.applyRunStyle({ fontSizePt: next });
  }

  private commitFontSizeInput(): void {
    const input = this.textToolbarControls?.fontSizeInput;
    if (!input) return;

    const value = Number(input.value);
    if (!Number.isFinite(value) || value < TEXT_TOOLBAR_MIN_FONT_SIZE) return;

    const clamped = Math.min(TEXT_TOOLBAR_MAX_FONT_SIZE, Math.round(value));
    this.applyRunStyle({ fontSizePt: clamped });
  }

  private flushActiveEditor(): void {
    const editor = this.activeEditor;
    const target = this.activeTextStyleTarget;
    if (!editor) return;

    this.removeActiveEditor(editor);
    if (target && editor.value !== target.text) {
      void this.applyTextValue(editor.value, target);
    }
  }

  private applyRunStyle(change: RunStyleChange): void {
    const engine = this.engine;
    if (!engine) return;
    void this.runTextFormatting('Format text', (shapeIndex, run, selection) => {
      if (selection) {
        return engine.setRunStyleForRange(
          this.currentSlide,
          shapeIndex,
          selection.paragraphIndex,
          selection.start,
          selection.end,
          change
        );
      }
      return engine.setRunStyle(this.currentSlide, shapeIndex, run, change);
    });
  }

  private applyAlignment(align: ParagraphAlignment): void {
    const engine = this.engine;
    if (!engine) return;
    void this.runTextFormatting('Align text', (shapeIndex, run, _selection) =>
      engine.setParagraphAlignment(this.currentSlide, shapeIndex, run ? run.paragraphIndex : null, align));
  }

  private async runTextFormatting(
    label: string,
    apply: (
      shapeIndex: number,
      run: RunTarget | null,
      selection: { paragraphIndex: number; start: number; end: number } | null
    ) => Promise<void>
  ): Promise<void> {
    const engine = this.engine;
    if (!engine || !this.ensureEditable('format text')) return;

    const context = this.getTextStyleContext();
    if (!context) return;

    // Capture the live inline-editor selection up front. The textarea itself
    // lives in the canvas pane (not the slide SVG), so it survives a slide
    // re-render — we keep it open and refresh its element references in place
    // afterwards instead of tearing it down and reopening, which is what used to
    // drop the selection and make a follow-up Bold apply to the whole shape.
    const editor = this.activeEditor;
    const styleTarget = this.activeTextStyleTarget;
    let pendingText: string | null = null;
    let selectionRange: { paragraphIndex: number; start: number; end: number } | null = null;
    let savedStart = 0;
    let savedEnd = 0;
    if (editor && styleTarget) {
      savedStart = Math.min(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      savedEnd = Math.max(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      selectionRange = { paragraphIndex: styleTarget.paragraphIndex, start: savedStart, end: savedEnd };
      pendingText = editor.value !== styleTarget.text ? editor.value : null;
    }

    const scrollPosition = this.captureCanvasScroll();
    try {
      const history = await this.captureHistoryEntry(label);
      if (pendingText !== null && selectionRange) {
        await engine.updateParagraphText(this.currentSlide, context.shapeIndex, selectionRange.paragraphIndex, pendingText);
      }
      await apply(context.shapeIndex, context.run, selectionRange);
      this.recordHistoryEntry(history);
      this.markDirty();

      const rendered = await this.renderCurrentSlide(true);
      if (rendered) {
        this.restoreCanvasScrollSoon(scrollPosition);
        await this.renderThumbnails();
        if (editor && this.activeEditor === editor && selectionRange) {
          if (this.refreshActiveShapeEditorAfterRender()) {
            const length = editor.value.length;
            this.clearWholeShapeInlineSelection();
            editor.setSelectionRange(Math.min(savedStart, length), Math.min(savedEnd, length));
            this.refreshInlineEditorGeometry();
          } else {
            this.removeActiveEditor(editor);
          }
        }
        this.updateTextToolbar();
      }
    } catch (error) {
      new Notice(`Could not format text: ${cleanError(error)}`);
    }
  }

  /**
   * After a slide re-render, re-point the still-open inline editor at the freshly
   * rendered paragraph/run nodes (the old ones are detached) and rebuild the
   * SVG-side caret. Returns false when the paragraph can no longer be found.
   */
  private refreshActiveShapeEditorAfterRender(): boolean {
    const target = this.activeShapeTextTarget;
    if (!target || !this.activeEditor) return false;

    const fresh = this.buildParagraphEditTarget(target.shapeIndex, target.paragraphIndex);
    if (!fresh) return false;

    this.activeEditorTarget?.classList.remove('native-powerpoint-text-editing');
    target.element = fresh.element;
    target.runElements = fresh.runElements;
    target.runIndex = fresh.runIndex;
    target.text = fresh.text;

    this.activeEditorTarget = fresh.element;
    this.activeEditorTarget.classList.add('native-powerpoint-text-editing');
    this.activeTextStyleTarget = this.getPrimaryStyleRunTarget(target);
    this.slideSurface?.addClass('is-inline-text-editing');

    // The previous caret/selection rects lived inside the SVG that was just
    // replaced, so drop the stale references and re-create the caret line.
    this.removeInlineSelection();
    this.activeInlineCaret = activeDocument.createElementNS('http://www.w3.org/2000/svg', 'line');
    this.activeInlineCaret.classList.add('native-powerpoint-svg-caret');
    this.activeInlineCaret.setAttribute('aria-hidden', 'true');
    this.svgEl?.appendChild(this.activeInlineCaret);
    this.removeSelectionOverlay();
    return true;
  }

  private refreshInlineEditorGeometry(): void {
    const editor = this.activeEditor;
    const target = this.activeShapeTextTarget;
    if (!editor || !target) return;

    const box = this.getElementBox(target.element);
    if (box) {
      this.positionTextRunEditor(editor, box);
    }
    this.focusEditorWithoutCanvasScroll(editor);
    this.updateInlineCaret(editor, target.element);
  }

  private openFontMenu(anchor: HTMLElement): void {
    const fonts = [...TEXT_TOOLBAR_FONTS];
    const context = this.getTextStyleContext();
    const current = this.currentRunStyle?.fontFamily
      ?? (context ? this.getEffectiveFontFamily(context) : null);
    if (current && !fonts.includes(current)) {
      fonts.unshift(current);
    }

    this.openToolbarPopover(anchor, (popover) => {
      popover.addClass('native-powerpoint-font-menu');
      for (const font of fonts) {
        const item = popover.createEl('button', {
          cls: 'native-powerpoint-color-popover-item native-powerpoint-font-menu-item',
          text: font
        });
        item.style.setProperty('--np-font-family', font);
        if (current === font) {
          item.addClass('is-active');
        }
        this.bindToolbarButton(item, () => {
          this.closeToolbarPopover();
          this.applyRunStyle({ fontFamily: font });
        });
      }
    });
  }

  private openColorPopover(
    anchor: HTMLElement,
    currentColor: string,
    allowNone: boolean,
    onPick: (color: string | null) => void
  ): void {
    this.openToolbarPopover(anchor, (popover) => {
      popover.addClass('native-powerpoint-color-popover');

      if (allowNone) {
        const noneButton = popover.createEl('button', {
          cls: 'native-powerpoint-color-popover-none',
          text: 'No color'
        });
        this.bindToolbarButton(noneButton, () => {
          this.closeToolbarPopover();
          onPick(null);
        });
      }

      const grid = popover.createDiv({ cls: 'native-powerpoint-color-popover-grid' });
      for (const swatch of TEXT_TOOLBAR_SWATCHES) {
        const cell = grid.createEl('button', {
          cls: 'native-powerpoint-color-popover-swatch',
          attr: { 'aria-label': `#${swatch}` }
        });
        cell.style.setProperty('--np-swatch-color', `#${swatch}`);
        if (swatch.toUpperCase() === currentColor.toUpperCase()) {
          cell.addClass('is-active');
        }
        this.bindToolbarButton(cell, () => {
          this.closeToolbarPopover();
          onPick(swatch);
        });
      }

      const customRow = popover.createDiv({ cls: 'native-powerpoint-color-popover-custom' });
      customRow.createSpan({ text: 'Custom' });
      const customInput = customRow.createEl('input', {
        type: 'color',
        attr: { 'aria-label': 'Custom color', value: `#${currentColor}` }
      });
      customInput.value = `#${currentColor}`;
      customInput.addEventListener('pointerdown', () => this.flushActiveEditor(), true);
      customInput.addEventListener('change', () => {
        const picked = customInput.value.replace(/^#/, '').toUpperCase();
        this.closeToolbarPopover();
        onPick(picked);
      });
    });
  }

  private openToolbarPopover(anchor: HTMLElement, build: (popover: HTMLElement) => void): void {
    this.closeToolbarPopover();

    const popover = activeDocument.body.createDiv({
      cls: 'native-powerpoint-toolbar-popover native-powerpoint-light-surface'
    });
    popover.addEventListener('pointerdown', (event) => event.stopPropagation());
    build(popover);

    const anchorRect = anchor.getBoundingClientRect();
    popover.setCssProps({ left: `${anchorRect.left}px`, top: `${anchorRect.bottom + 4}px` });

    const onOutsidePointerDown = (event: PointerEvent): void => {
      const target = isNode(event.target) ? event.target : null;
      if (target && (popover.contains(target) || anchor.contains(target))) return;
      this.closeToolbarPopover();
    };
    activeDocument.addEventListener('pointerdown', onOutsidePointerDown, true);

    this.activeToolbarPopover = popover;
    this.toolbarPopoverCleanup = () => {
      activeDocument.removeEventListener('pointerdown', onOutsidePointerDown, true);
    };
  }

  private closeToolbarPopover(): void {
    this.toolbarPopoverCleanup?.();
    this.toolbarPopoverCleanup = null;
    this.activeToolbarPopover?.remove();
    this.activeToolbarPopover = null;
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
    if (this.activeShapeTextTarget) {
      this.syncShapeParagraphPreview(this.activeShapeTextTarget, nextText);
    } else if (isSVGTextElement(element)) {
      element.textContent = nextText;
    }
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

    if (this.inlineWholeShapeSelected) {
      this.renderWholeShapeInlineSelection();
      return;
    }

    const selectionStart = Math.max(0, Math.min(editor.selectionStart ?? editor.value.length, editor.value.length));
    const selectionEnd = Math.max(0, Math.min(editor.selectionEnd ?? editor.value.length, editor.value.length));
    const start = Math.min(selectionStart, selectionEnd);
    const end = Math.max(selectionStart, selectionEnd);
    if (start === end) return;

    this.renderInlineSelectionRects(element, start, end);
  }

  private renderInlineSelectionRects(
    element: SVGTextElement | SVGTSpanElement,
    start: number,
    end: number
  ): void {
    const boxes = this.getSvgInlineSelectionBoxes(element, start, end);
    const textElement = element.closest('text');
    const parent = textElement?.parentNode;
    if (!isSVGTextElement(textElement) || !parent) return;

    for (const box of boxes) {
      const rect = activeDocument.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect.classList.add('native-powerpoint-svg-selection');
      rect.setAttribute('x', this.formatSvgNumber(box.x));
      rect.setAttribute('y', this.formatSvgNumber(box.y));
      rect.setAttribute('width', this.formatSvgNumber(box.width));
      rect.setAttribute('height', this.formatSvgNumber(box.height));
      parent.insertBefore(rect, textElement);
      this.activeInlineSelectionRects.push(rect);
    }
  }

  private getShapeTextParagraphs(shape: Element): (SVGTextElement | SVGTSpanElement)[] {
    const result: (SVGTextElement | SVGTSpanElement)[] = [];
    for (const text of Array.from(shape.querySelectorAll('text'))) {
      if (text.closest(GENERATED_GRID_SELECTOR)) continue;
      const paragraphs = Array.from(text.querySelectorAll('tspan[data-ooxml-para-idx]')).filter(isSVGTSpanElement);
      if (paragraphs.length > 0) {
        result.push(...paragraphs);
      } else if (isSVGTextElement(text) && (text.textContent ?? '').length > 0) {
        result.push(text);
      }
    }
    return result;
  }

  private renderWholeShapeInlineSelection(): void {
    const shape = this.getSelectedShapeElement();
    if (!shape) return;
    for (const paragraph of this.getShapeTextParagraphs(shape)) {
      const total = this.getLeafCharInfo(paragraph).total;
      if (total <= 0) continue;
      this.renderInlineSelectionRects(paragraph, 0, total);
    }
  }

  private selectAllInlineText(editor: HTMLTextAreaElement, element: SVGTextElement | SVGTSpanElement): void {
    const shape = this.getSelectedShapeElement();
    if (!shape) return;

    const paragraphs = this.getShapeTextParagraphs(shape);
    const combined = paragraphs.map((paragraph) => paragraph.textContent ?? '').join('\n');
    this.inlineWholeShapeSelection = combined;
    this.inlineWholeShapeSelected = true;
    this.lastInlineCaretPlacement = null;
    editor.setSelectionRange(0, editor.value.length);
    this.updateInlineCaret(editor, element);
  }

  private clearWholeShapeInlineSelection(): void {
    this.inlineWholeShapeSelection = null;
    this.inlineWholeShapeSelected = false;
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
    const top = screenGeometry.top;
    const height = screenGeometry.height;
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
    const rootMatrix = this.svgEl?.getScreenCTM();
    if (!rootMatrix) return [];

    let rootInverse: DOMMatrix;
    try {
      rootInverse = rootMatrix.inverse();
    } catch {
      return [];
    }

    const { entries, total } = this.getLeafCharInfo(element);
    if (total <= 0) return [];

    const normalizedStart = Math.max(0, Math.min(total, start));
    const normalizedEnd = Math.max(normalizedStart, Math.min(total, end));
    const rows: SvgInlineSelectionBox[] = [];
    for (let index = normalizedStart; index < normalizedEnd; index++) {
      const entry = entries.find((candidate) => index >= candidate.start && index < candidate.start + candidate.count);
      const elementMatrix = entry?.span.getScreenCTM();
      if (!entry || !elementMatrix) continue;

      let charBox: SvgInlineSelectionBox | null = null;
      try {
        charBox = this.transformSvgRectToSvgRoot(
          entry.span.getExtentOfChar(index - entry.start),
          elementMatrix,
          rootInverse
        );
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

  private getInlineTextOffsetAtClientPointForElement(
    element: SVGTextElement | SVGTSpanElement,
    clientX: number,
    clientY: number | undefined,
    box: { left: number; top: number; width: number; height: number }
  ): number {
    const text = element.textContent ?? '';
    if (text.length === 0) return 0;

    const paneRect = this.canvasPane?.getBoundingClientRect();
    const localClientX = paneRect
      ? clientX - paneRect.left + (this.canvasPane?.scrollLeft ?? 0)
      : clientX;
    const localClientY = paneRect && clientY !== undefined
      ? clientY - paneRect.top + (this.canvasPane?.scrollTop ?? 0)
      : box.top + box.height / 2;

    const geometryOffset = this.getInlineTextOffsetFromSvgGeometry(element, localClientX, localClientY, text.length);
    return geometryOffset ?? Math.max(0, Math.min(text.length, Math.round(text.length * ((localClientX - box.left) / Math.max(1, box.width)))));
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
      return { left: svgGeometry.left, top: svgGeometry.top, height: svgGeometry.height };
    }

    const row = this.activeInlineCaretRow ?? this.getDefaultInlineCaretRow(element, box, fallbackHeight);

    const text = editor.value;
    if (text.length === 0) return { left: box.left, ...row };

    const style = window.getComputedStyle(editor);
    const canvas = activeDocument.createElement('canvas');
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

    const paneRect = this.canvasPane?.getBoundingClientRect();
    if (!paneRect) return 1;

    try {
      const rows: number[] = [];
      for (const { span, count } of this.getLeafCharInfo(element).entries) {
        const matrix = span.getScreenCTM();
        if (!matrix) continue;
        for (let index = 0; index < count; index++) {
          const position = span.getStartPositionOfChar(index);
          const point = new DOMPoint(position.x, position.y).matrixTransform(matrix);
          const localY = point.y - paneRect.top + (this.canvasPane?.scrollTop ?? 0);
          if (!rows.some((row) => Math.abs(row - localY) < 4)) {
            rows.push(localY);
          }
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
    const canvas = activeDocument.createElement('canvas');
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

  private getLeafTextSpans(element: SVGTextElement | SVGTSpanElement): SVGTextContentElement[] {
    // Include every innermost tspan (bullet markers, runs, etc.), not only
    // data-ooxml-run-idx nodes. Paragraph textContent also contains bullet
    // prefixes, so limiting to runs desynchronizes string indices from glyph
    // geometry and misplaces find/selection highlights.
    const leafTspans = Array.from(element.querySelectorAll('tspan'))
      .filter(isSVGTSpanElement)
      .filter((span) => !span.querySelector('tspan'));
    if (leafTspans.length > 0) {
      return leafTspans as SVGTextContentElement[];
    }
    return [element as SVGTextContentElement];
  }

  private getParagraphLeafText(element: SVGTextElement | SVGTSpanElement): string {
    return this.getLeafCharInfo(element).entries
      .map((entry) => entry.span.textContent || '')
      .join('');
  }

  private getLeafCharInfo(
    element: SVGTextElement | SVGTSpanElement
  ): { entries: { span: SVGTextContentElement; count: number; start: number }[]; total: number } {
    const entries: { span: SVGTextContentElement; count: number; start: number }[] = [];
    let total = 0;
    for (const span of this.getLeafTextSpans(element)) {
      let count = 0;
      try {
        count = span.getNumberOfChars();
      } catch {
        count = (span.textContent || '').length;
      }
      if (count <= 0) continue;
      entries.push({ span, count, start: total });
      total += count;
    }
    return { entries, total };
  }

  private getSvgTextCaretGeometry(
    element: SVGTextElement | SVGTSpanElement,
    offset: number,
    preferredHeight = this.getScreenFontSize(element) * 1.08
  ): { left: number; top: number; height: number } | null {
    const text = element.textContent || '';
    if (!text) return this.getFallbackInlineCaretGeometry(element, preferredHeight);

    const paneRect = this.canvasPane?.getBoundingClientRect();
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

    const localIndex = Math.max(0, Math.min(entry.count - 1, globalCharIndex - entry.start));
    const matrix = entry.span.getScreenCTM();
    if (!matrix) return null;

    let position: DOMPoint;
    let extent: SvgRectLike | null = null;
    try {
      position = useStart
        ? entry.span.getStartPositionOfChar(localIndex)
        : entry.span.getEndPositionOfChar(localIndex);
      extent = entry.span.getExtentOfChar(localIndex);
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
      if (bounds && bounds.height > 0) {
        // Snap to the real glyph row so the caret fills the actual line height
        // and aligns vertically to the text, instead of floating at the click Y.
        height = Math.max(6, bounds.height);
        top = bounds.top;
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
    rect: SvgRectLike,
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
    rect: SvgRectLike,
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
    return isSVGGElement(shape) ? shape : null;
  }

  private getTextEditTargetFromSelectedShape(): ShapeTextEditTarget | null {
    const shape = this.getSelectedShapeElement();
    if (!shape) return null;
    const textElement = shape.querySelector('text');
    const target = this.getTextEditTarget(textElement);
    return target?.kind === 'shape-paragraph' ? target : null;
  }

  private getPrimaryStyleRunTarget(target: ShapeTextEditTarget): ShapeTextEditTarget {
    const run = target.runElements[0];
    if (!run) return target;
    return { ...target, element: run };
  }

  private syncShapeParagraphPreview(target: TextEditTarget, text: string): void {
    if (target.kind === 'shape-paragraph') {
      const firstRun = target.runElements[0];
      if (firstRun) {
        firstRun.textContent = text;
        for (let index = 1; index < target.runElements.length; index++) {
          const run = target.runElements[index];
          if (run) run.textContent = '';
        }
        return;
      }
      if (isSVGTextElement(target.element)) {
        target.element.textContent = text;
      }
      return;
    }

    target.element.textContent = text;
  }

  private collectParagraphRuns(paraContainer: Element): SVGTSpanElement[] {
    const runs = paraContainer.matches('tspan[data-ooxml-para-idx]')
      ? Array.from(paraContainer.querySelectorAll(':scope > tspan[data-ooxml-run-idx]'))
      : Array.from(paraContainer.querySelectorAll('tspan[data-ooxml-run-idx]'));
    return runs.filter(isSVGTSpanElement);
  }

  private collectParagraphLineContainers(textEl: SVGTextElement, paragraphIndex: number): SVGTSpanElement[] {
    const direct = Array.from(textEl.children).filter(
      (child): child is SVGTSpanElement =>
        isSVGTSpanElement(child) && child.getAttribute('data-ooxml-para-idx') === String(paragraphIndex)
    );
    if (direct.length > 0) return direct;

    const nested = textEl.querySelector(`tspan[data-ooxml-para-idx="${paragraphIndex}"]`);
    return nested && isSVGTSpanElement(nested) ? [nested] : [];
  }

  private getParagraphPlainText(lineContainers: SVGTSpanElement[]): string {
    return lineContainers
      .map((container) => this.collectParagraphRuns(container).map((run) => run.textContent || '').join(''))
      .join('\n');
  }

  private getTextEditTarget(element: Element | null): TextEditTarget | null {
    const textEl = element?.closest('text');
    if (!isSVGTextElement(textEl) || textEl.closest(GENERATED_GRID_SELECTOR)) return null;

    const shape = textEl.closest('g[data-ooxml-shape-idx]');
    const shapeIndex = getShapeIndex(shape);
    if (shapeIndex === null) return null;

    const clickedRun = element?.closest('tspan[data-ooxml-run-idx]');
    let paraContainer: Element | null = clickedRun?.closest('tspan[data-ooxml-para-idx]') ?? null;
    if (!paraContainer) {
      paraContainer = textEl.querySelector('tspan[data-ooxml-para-idx]') ?? textEl;
    }

    const seedRuns = this.collectParagraphRuns(paraContainer);
    if (seedRuns.length === 0) {
      const text = textEl.textContent || '';
      if (!text) return null;
      return {
        kind: 'shape-paragraph',
        shapeIndex,
        paragraphIndex: 0,
        runIndex: 0,
        text,
        element: textEl,
        runElements: []
      };
    }

    const firstRun = seedRuns[0];
    if (!firstRun) return null;

    const paragraph = firstRun.closest('tspan[data-ooxml-para-idx]');
    const paragraphIndex = Number(paragraph?.getAttribute('data-ooxml-para-idx') ?? 0);
    const resolvedParagraphIndex = Number.isFinite(paragraphIndex) ? paragraphIndex : 0;
    const lineContainers = this.collectParagraphLineContainers(textEl, resolvedParagraphIndex);
    const geometryElement = lineContainers[0] ?? (paragraph && isSVGTSpanElement(paragraph) ? paragraph : textEl);
    const runElements = lineContainers.flatMap((container) => this.collectParagraphRuns(container));
    const firstRunIndex = Number(firstRun.getAttribute('data-ooxml-run-idx') ?? 0);
    const text = this.getParagraphPlainText(lineContainers.length > 0 ? lineContainers : [paraContainer].filter(isSVGTSpanElement));

    return {
      kind: 'shape-paragraph',
      shapeIndex,
      paragraphIndex: resolvedParagraphIndex,
      runIndex: Number.isFinite(firstRunIndex) ? firstRunIndex : 0,
      text,
      element: geometryElement,
      runElements
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
      !isSVGTextElement(textElement)
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
    this.updateMultiSelectionBoxes();
    if (!this.canvasPane || this.selectedShapeIndex === null) {
      this.removeSelectionOverlay();
      this.updateTextToolbar();
      return;
    }

    const box = this.getSelectedBox();
    if (!box) {
      this.removeSelectionOverlay();
      this.updateTextToolbar();
      return;
    }

    if (!this.selectionOverlay) {
      this.selectionOverlay = this.canvasPane.createDiv({ cls: 'native-powerpoint-selection-box' });
      if (this.canEdit()) {
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
        rotateHandle.setAttribute('aria-label', 'Rotate object');
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

    this.updateTextToolbar();
  }

  private removeSelectionOverlay(): void {
    this.selectionOverlay?.remove();
    this.selectionOverlay = null;
  }

  private updateMultiSelectionBoxes(): void {
    this.removeMultiSelectionBoxes();
    if (!this.canvasPane || !this.svgEl || this.selectedShapeIndices.size <= 1) return;

    for (const index of this.selectedShapeIndices) {
      const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
      if (!isSVGGElement(shape)) continue;

      const box = this.getElementBox(shape);
      if (!box) continue;

      const boxEl = this.canvasPane.createDiv({
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

  private removeMultiSelectionBoxes(): void {
    for (const box of this.multiSelectionBoxes) {
      box.remove();
    }
    this.multiSelectionBoxes = [];
  }

  private collectShapesInClientRect(left: number, top: number, right: number, bottom: number): number[] {
    if (!this.svgEl) return [];

    const indices: number[] = [];
    this.svgEl.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
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

  private previewSelectionClasses(indices: Set<number>): void {
    this.svgEl?.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
      const index = getShapeIndex(shape);
      if (index !== null && indices.has(index)) {
        shape.addClass('native-powerpoint-shape-selected');
      } else {
        shape.removeClass('native-powerpoint-shape-selected');
      }
    });
  }

  private beginMarquee(event: PointerEvent, additive: boolean): void {
    if (!this.canvasPane) return;

    this.cancelMarquee();
    this.marquee = {
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      additive,
      base: [...this.selectedShapeIndices],
      moved: false
    };
  }

  private updateMarquee(event: PointerEvent): void {
    if (!this.marquee || event.pointerId !== this.marquee.pointerId || !this.canvasPane) return;

    const deltaX = event.clientX - this.marquee.startClientX;
    const deltaY = event.clientY - this.marquee.startClientY;
    if (!this.marquee.moved && Math.hypot(deltaX, deltaY) < 4) return;
    if (!this.marquee.moved) {
      this.removeSelectionOverlay();
      this.removeMultiSelectionBoxes();
    }
    this.marquee.moved = true;

    if (!this.marqueeEl) {
      this.marqueeEl = this.canvasPane.createDiv({ cls: 'native-powerpoint-marquee-box' });
    }

    const paneRect = this.canvasPane.getBoundingClientRect();
    const left = Math.min(event.clientX, this.marquee.startClientX);
    const top = Math.min(event.clientY, this.marquee.startClientY);
    const width = Math.abs(deltaX);
    const height = Math.abs(deltaY);
    this.marqueeEl.setCssProps({
      left: `${left - paneRect.left + this.canvasPane.scrollLeft}px`,
      top: `${top - paneRect.top + this.canvasPane.scrollTop}px`,
      width: `${width}px`,
      height: `${height}px`
    });

    const hits = this.collectShapesInClientRect(left, top, left + width, top + height);
    const preview = new Set<number>(this.marquee.additive ? this.marquee.base : []);
    hits.forEach((index) => preview.add(index));
    this.previewSelectionClasses(preview);
  }

  private finishMarquee(event: PointerEvent): void {
    if (!this.marquee || event.pointerId !== this.marquee.pointerId) return;

    const marquee = this.marquee;
    this.marquee = null;
    this.marqueeEl?.remove();
    this.marqueeEl = null;

    if (!marquee.moved) {
      if (marquee.additive) {
        this.suppressNextClick = true;
        this.applyMultiSelection(marquee.base);
      } else {
        this.clearSelection();
      }
      return;
    }

    this.suppressNextClick = true;
    const left = Math.min(event.clientX, marquee.startClientX);
    const top = Math.min(event.clientY, marquee.startClientY);
    const right = Math.max(event.clientX, marquee.startClientX);
    const bottom = Math.max(event.clientY, marquee.startClientY);
    const hits = this.collectShapesInClientRect(left, top, right, bottom);
    const finalSet = new Set<number>(marquee.additive ? marquee.base : []);
    hits.forEach((index) => finalSet.add(index));
    this.applyMultiSelection([...finalSet]);
  }

  private cancelMarquee(): void {
    this.marquee = null;
    this.marqueeEl?.remove();
    this.marqueeEl = null;
  }

  private startGroupDrag(event: PointerEvent): void {
    if (!this.engine || !this.svgEl) return;

    const startPoint = this.getSvgPoint(event);
    if (!startPoint) return;

    const start = new Map<number, ShapeTransform>();
    for (const index of this.selectedShapeIndices) {
      const shape = this.svgEl.querySelector(`g[data-ooxml-shape-idx="${index}"]`);
      if (isSVGGElement(shape)) {
        start.set(index, cloneTransform(this.engine.getShapeTransform(shape)));
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

  private updateGroupDrag(event: PointerEvent): void {
    if (!this.groupDrag || event.pointerId !== this.groupDrag.pointerId || !this.engine || !this.svgEl) {
      return;
    }

    const deltaClientX = event.clientX - this.groupDrag.startClientX;
    const deltaClientY = event.clientY - this.groupDrag.startClientY;
    if (!this.groupDrag.moved && Math.hypot(deltaClientX, deltaClientY) < 3) return;
    this.groupDrag.moved = true;

    const point = this.getSvgPoint(event);
    if (!point) return;

    const scale = this.engine.getSlideScale(this.svgEl);
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
      ? this.computeSnap(
          { x: minX, y: minY, cx: maxX - minX, cy: maxY - minY },
          new Set(this.selectedShapeIndices)
        )
      : { dx: 0, dy: 0, guideX: null, guideY: null };
    if (snap.dx !== 0 || snap.dy !== 0) {
      this.groupDrag.latest.forEach((transform) => {
        transform.x += snap.dx;
        transform.y += snap.dy;
      });
    }
    this.updateSnapGuides(snap.guideX, snap.guideY);

    const ctm = this.svgEl.getScreenCTM();
    const snapClientX = ctm && ctm.a !== 0 ? (snap.dx * ctm.a) / scale : 0;
    const snapClientY = ctm && ctm.d !== 0 ? (snap.dy * ctm.d) / scale : 0;
    const cssTransform = `translate(${deltaClientX + snapClientX}px, ${deltaClientY + snapClientY}px)`;
    for (const box of this.multiSelectionBoxes) {
      box.style.transform = cssTransform;
    }
  }

  private finishGroupDrag(event: PointerEvent): void {
    if (!this.groupDrag || event.pointerId !== this.groupDrag.pointerId) return;

    const groupDrag = this.groupDrag;
    this.groupDrag = null;
    this.clearSnapGuides();
    for (const box of this.multiSelectionBoxes) {
      box.style.removeProperty('transform');
    }

    if (!groupDrag.moved) return;

    this.suppressNextClick = true;
    const updates = [...groupDrag.latest.entries()].map(([index, transform]) => ({ index, transform }));
    void this.commitGroupTransforms(updates);
  }

  private async commitGroupTransforms(
    updates: { index: number; transform: ShapeTransform }[],
    label = 'Move objects'
  ): Promise<void> {
    if (!this.engine || updates.length === 0) return;
    if (!this.ensureEditable('move objects')) return;

    try {
      const changed = updates.filter((update) => {
        const shape = this.svgEl?.querySelector(`g[data-ooxml-shape-idx="${update.index}"]`);
        return !(
          isSVGGElement(shape)
          && this.engine !== null
          && transformsMatch(this.engine.getShapeTransform(shape), update.transform)
        );
      });
      if (changed.length === 0) return;

      const history = await this.captureHistoryEntry(label);
      for (const update of changed) {
        await this.engine.updateShapeTransform(this.currentSlide, update.index, update.transform);
      }
      this.recordHistoryEntry(history);
      this.markDirty();
      const indices = updates.map((update) => update.index);
      const rendered = await this.renderCurrentSlide();
      if (rendered) {
        this.applyMultiSelection(indices);
        await this.renderThumbnails();
      }
    } catch (error) {
      new Notice(`Could not move objects: ${cleanError(error)}`);
    }
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
    if (this.marquee) {
      this.updateMarquee(event);
      return;
    }
    if (this.groupDrag) {
      this.updateGroupDrag(event);
      return;
    }
    if (!this.dragState || !this.engine || !this.svgEl) return;
    if (event.pointerId !== this.dragState.pointerId) return;

    if (this.dragState.mode === 'rotate') {
      this.updateRotateDrag(event);
      return;
    }

    const point = this.getSvgPoint(event);
    if (!point) return;

    const scale = this.engine.getSlideScale(this.svgEl);
    const dx = (point.x - this.dragState.startPoint.x) * scale;
    const dy = (point.y - this.dragState.startPoint.y) * scale;
    const next = cloneTransform(this.dragState.startTransform);

    if (this.dragState.mode === 'move') {
      next.x += dx;
      next.y += dy;
      const snap = this.computeSnap(
        { x: next.x, y: next.y, cx: next.cx, cy: next.cy },
        new Set(this.selectedShapeIndex === null ? [] : [this.selectedShapeIndex])
      );
      next.x += snap.dx;
      next.y += snap.dy;
      this.updateSnapGuides(snap.guideX, snap.guideY);
      this.dragState.latestTransform = next;
      this.selectedTransform = cloneTransform(next);
      this.updateInspectorValues();
      this.positionOverlayFromTransform(next);
      return;
    }

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

    this.dragState.latestTransform = next;
    this.selectedTransform = cloneTransform(next);
    this.updateInspectorValues();
    this.updateSelectionOverlayDuringDrag(event);
  };

  private updateRotateDrag(event: PointerEvent): void {
    if (!this.engine || !this.dragState) return;

    const centerX = this.dragState.centerClientX ?? 0;
    const centerY = this.dragState.centerClientY ?? 0;
    const angle = Math.atan2(event.clientY - centerY, event.clientX - centerX);
    const deltaDegrees = ((angle - (this.dragState.startAngle ?? 0)) * 180) / Math.PI;
    let degrees = this.engine.ooxmlToDegrees(this.dragState.startTransform.rot) + deltaDegrees;
    if (event.shiftKey) degrees = Math.round(degrees / 15) * 15;
    degrees = ((degrees % 360) + 360) % 360;

    const next = cloneTransform(this.dragState.startTransform);
    next.rot = this.engine.degreesToOoxml(degrees);
    this.dragState.latestTransform = next;
    this.selectedTransform = cloneTransform(next);
    this.updateInspectorValues();
    if (this.selectionOverlay) {
      this.selectionOverlay.style.transform = `rotate(${degrees}deg)`;
    }
  }

  private handleDragEnd = (event: PointerEvent): void => {
    if (this.marquee) {
      this.finishMarquee(event);
      return;
    }
    if (this.groupDrag) {
      this.finishGroupDrag(event);
      return;
    }
    if (!this.dragState || event.pointerId !== this.dragState.pointerId) return;

    this.clearSnapGuides();
    const transform = cloneTransform(this.dragState.latestTransform);
    this.dragState = null;
    void this.commitTransform(transform);
  };

  private async commitTransform(transform: ShapeTransform): Promise<void> {
    if (!this.engine || this.selectedShapeIndex === null) return;
    if (!this.ensureEditable('edit object')) return;

    try {
      const selected = this.getSelectedShapeElement();
      const shapeIndex = selected ? getShapeIndex(selected) : this.selectedShapeIndex;
      if (shapeIndex === null) return;
      if (selected && transformsMatch(this.engine.getShapeTransform(selected), transform)) return;

      const history = await this.captureHistoryEntry('Edit layout');
      await this.engine.updateShapeTransform(this.currentSlide, shapeIndex, transform);
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
    this.updateZoomLabel();
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
    if (this.findReplaceInputEl) this.findReplaceInputEl.value = '';
    this.setFindReplaceMode(false);
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
      button.setAttribute('aria-label', canEdit ? baseTitle : `${baseTitle}: ${disabledReason}`);
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
    this.updateArrangeAvailability();
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
    button.setAttribute('aria-label', nextEntry ? `${label} ${nextEntry.label} (${shortcut})` : `${label} (${shortcut})`);
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
    this.updateHeaderTitle();
    const count = this.engine?.slideCount || 0;
    this.slideCounterEl?.setText(count ? `${this.currentSlide + 1} / ${count}` : '0 / 0');

    if (this.thumbnailContainer) {
      this.thumbnailContainer.querySelectorAll('.native-powerpoint-thumbnail').forEach((thumbnail, index) => {
        thumbnail.toggleClass('active', index === this.currentSlide);
      });
    }
  }
}
