import type { TFile } from 'obsidian';
import { forwardRef, useCallback, useEffect, useImperativeHandle, useMemo, useRef, useState, type ChangeEvent, type ComponentProps } from 'react';
import { DocxEditor, type DocxEditorRef, type EditorMode } from '@eigenpal/docx-editor-react';
import type { RenderedDomContext } from '@eigenpal/docx-editor-core/plugin-api';
import { insertTable, setFontFamily, setFontSize } from '@eigenpal/docx-editor-core/prosemirror/commands';
import { loadFontFromBuffer } from '@eigenpal/docx-editor-core/utils';
import type { FontOption } from '@eigenpal/docx-editor-core/utils/fontOptions';
import type { Translations } from '@eigenpal/docx-editor-i18n';
import { Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import editorStyles from '@eigenpal/docx-editor-react/styles.css';
import { createEditorTranslator } from './editorTranslations';
import { debugLog, errorLog, warnLog } from './logger';
import { Notice, Platform, setIcon } from './obsidianRuntime';
import { exportRenderedPagesToPdf } from './renderedPdfExport';

let stylesInjected = false;
let editorInstanceCounter = 0;

interface DocxSectionProperties {
	pageHeight?: number;
	marginTop?: number;
	marginBottom?: number;
}

interface DocxDocumentWithSectionProperties {
	package?: {
		document?: {
			finalSectionProperties?: DocxSectionProperties;
			sections?: Array<{
				properties?: DocxSectionProperties;
			}>;
		};
	};
}

const DEFAULT_PAGE_HEIGHT_TWIPS = 15840;
const DEFAULT_MARGIN_TWIPS = 1440;
const MIN_TOUCH_ZOOM = 0.25;
const MAX_TOUCH_ZOOM = 4;
const TOUCH_ZOOM_SENSITIVITY = 0.55;
const TOUCH_ZOOM_MIN_DELTA = 0.006;
const MAX_INSERTED_IMAGE_WIDTH = 612;
const IMPORT_FONT_MENU_LABEL = 'Import font...';
const CUSTOM_TABLE_MENU_LABEL = 'Custom...';
const FONT_FILE_ACCEPT = '.ttf,.otf,.woff,.woff2,font/ttf,font/otf,font/woff,font/woff2,application/font-woff,application/font-woff2';
const FONT_FILE_EXTENSION_PATTERN = /\.(?:ttf|otf|woff2?)$/i;
const MIN_CUSTOM_TABLE_SIZE = 1;
const MAX_CUSTOM_TABLE_SIZE = 50;
const DEFAULT_FONT_SIZE_POINTS = 11;
const MIN_FONT_SIZE_POINTS = 1;
const MAX_FONT_SIZE_POINTS = 400;
const FONT_SIZE_HOLD_INITIAL_DELAY_MS = 420;
const FONT_SIZE_HOLD_INITIAL_INTERVAL_MS = 180;
const FONT_SIZE_HOLD_INTERVAL_DECAY = 0.82;
const FONT_SIZE_HOLD_MIN_INTERVAL_MS = 36;
const TOOLBAR_TOOLTIP_DELAY_MS = 450;
const SELECTED_LIST_MARKER_CLASS = 'docxidian-list-marker-selected';
const LIST_PARAGRAPH_SELECTOR = '.layout-paragraph[data-pm-start]';
const LIST_MARKER_SELECTOR = '.layout-list-marker, .docx-list-marker';
const DEFAULT_EDITOR_FONT_FAMILIES: FontOption[] = [
	{ name: 'Arial', fontFamily: 'Arial, Helvetica, sans-serif', category: 'sans-serif' },
	{ name: 'Calibri', fontFamily: '"Calibri", Arial, sans-serif', category: 'sans-serif' },
	{ name: 'Helvetica', fontFamily: 'Helvetica, Arial, sans-serif', category: 'sans-serif' },
	{ name: 'Verdana', fontFamily: 'Verdana, Geneva, sans-serif', category: 'sans-serif' },
	{ name: 'Open Sans', fontFamily: '"Open Sans", sans-serif', category: 'sans-serif' },
	{ name: 'Roboto', fontFamily: 'Roboto, sans-serif', category: 'sans-serif' },
	{ name: 'Times New Roman', fontFamily: '"Times New Roman", Times, serif', category: 'serif' },
	{ name: 'Georgia', fontFamily: 'Georgia, serif', category: 'serif' },
	{ name: 'Cambria', fontFamily: 'Cambria, Georgia, serif', category: 'serif' },
	{ name: 'Garamond', fontFamily: 'Garamond, serif', category: 'serif' },
	{ name: 'Courier New', fontFamily: '"Courier New", Courier, monospace', category: 'monospace' },
	{ name: 'Consolas', fontFamily: 'Consolas, monospace', category: 'monospace' },
];

function isPrimaryShortcut(evt: KeyboardEvent, key: string): boolean {
	const normalizedKey = evt.key.toLowerCase();
	const isMacShortcut = evt.metaKey && !evt.ctrlKey;
	const isNonMacShortcut = evt.ctrlKey && !evt.metaKey && !Platform.isMacOS;
	const hasPrimaryModifier = isMacShortcut || isNonMacShortcut;
	return normalizedKey === key && hasPrimaryModifier && !evt.altKey && !evt.shiftKey;
}

type FindReplaceMode = 'find' | 'replace';
type SaveStatus = 'saved' | 'saving' | 'unsaved' | 'failed';
type FontSizeStepDirection = -1 | 1;

interface FontSizeHoldState {
	control: HTMLElement | null;
	currentSize: number;
	direction: FontSizeStepDirection;
	repeatCount: number;
	repeatTimer: number | null;
	startTimer: number | null;
}

interface SaveDocumentOptions {
	silent?: boolean;
	dirtyVersion?: number;
}

interface ExportDocumentBufferOptions {
	preserveAutosave?: boolean;
}

interface FindMatch {
	from: number;
	to: number;
	text: string;
}

interface RefreshFindOptions {
	select?: boolean;
}

export interface PasteClipboardOptions {
	preserveFormatting: boolean;
}

interface FindHighlightState {
	matches: FindMatch[];
	currentIndex: number;
}

interface PinchZoomState {
	source: 'touch' | 'gesture' | 'pointer';
	startDistance: number;
	lastDistance: number;
	startZoom: number;
	lastZoom: number;
}

interface PointerPoint {
	x: number;
	y: number;
}

type WebKitGestureEvent = Event & {
	clientX?: number;
	clientY?: number;
	scale?: number;
};

interface FindReplaceLabels {
	find: string;
	findAndReplace: string;
	findText: string;
	replaceWith: string;
	replace: string;
	replaceAll: string;
	matchCase: string;
	wholeWords: string;
	showReplace: string;
	close: string;
	previous: string;
	next: string;
	noMatches: string;
	resultCount: (current: number, total: number) => string;
}

function createFindReplaceLabels(i18n: Translations | undefined): FindReplaceLabels {
	const translate = createEditorTranslator(i18n);

	return {
		find: translate('dialogs.findReplace.titleFind', undefined, 'Find'),
		findAndReplace: translate('dialogs.findReplace.titleFindReplace', undefined, 'Find and Replace'),
		findText: translate('dialogs.findReplace.findPlaceholder', undefined, 'Find text'),
		replaceWith: translate('dialogs.findReplace.replacePlaceholder', undefined, 'Replace with'),
		replace: translate('dialogs.findReplace.replaceButton', undefined, 'Replace'),
		replaceAll: translate('dialogs.findReplace.replaceAllButton', undefined, 'Replace all'),
		matchCase: translate('dialogs.findReplace.matchCase', undefined, 'Match case'),
		wholeWords: translate('dialogs.findReplace.wholeWords', undefined, 'Whole words'),
		showReplace: translate('dialogs.findReplace.toggleReplace', undefined, 'Show replace'),
		close: translate('common.close', undefined, 'Close'),
		previous: translate('dialogs.findReplace.findPrevious', undefined, 'Previous match'),
		next: translate('dialogs.findReplace.findNext', undefined, 'Next match'),
		noMatches: translate('dialogs.findReplace.noResults', undefined, 'No matches'),
		resultCount: (current, total) => translate(
			'dialogs.findReplace.matchCount',
			{ current, total },
			`${current} of ${total}`,
		),
	};
}

const findHighlightPluginKey = new PluginKey<FindHighlightState>('docxidian-find-highlight');
const preserveTypedSpacePluginKey = new PluginKey('docxidian-preserve-typed-space');

function insertPlainTypedText(view: EditorView, text: string, from = view.state.selection.from, to = view.state.selection.to) {
	view.dispatch(view.state.tr.insertText(text, from, to).scrollIntoView());
	return true;
}

function readFileAsDataUrl(file: File) {
	return new Promise<string>((resolve, reject) => {
		const reader = new FileReader();
		reader.onerror = () => reject(new Error(`Could not read ${file.name}.`));
		reader.onload = () => {
			if (typeof reader.result === 'string') {
				resolve(reader.result);
				return;
			}

			reject(new Error('The selected image could not be read.'));
		};
		reader.readAsDataURL(file);
	});
}

function scaleImageDimensions(width: number, height: number) {
	if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
		return { width: MAX_INSERTED_IMAGE_WIDTH, height: Math.round(MAX_INSERTED_IMAGE_WIDTH * 0.75) };
	}

	if (width <= MAX_INSERTED_IMAGE_WIDTH) {
		return { width, height };
	}

	const scale = MAX_INSERTED_IMAGE_WIDTH / width;
	return { width: MAX_INSERTED_IMAGE_WIDTH, height: Math.round(height * scale) };
}

function loadImageDimensions(src: string) {
	return new Promise<{ width: number; height: number }>((resolve, reject) => {
		const image = new Image();
		image.onerror = () => reject(new Error('The selected file is not a readable image.'));
		image.onload = () => resolve(scaleImageDimensions(image.naturalWidth, image.naturalHeight));
		image.src = src;
	});
}

function isSupportedFontFile(file: File) {
	return FONT_FILE_EXTENSION_PATTERN.test(file.name) || file.type.startsWith('font/');
}

function getImportedFontBaseName(file: File) {
	return file.name
		.replace(FONT_FILE_EXTENSION_PATTERN, '')
		.replace(/[-_]+/g, ' ')
		.replace(/\s+/g, ' ')
		.trim() || 'Imported Font';
}

function getUniqueImportedFontName(file: File, fonts: FontOption[]) {
	const baseName = getImportedFontBaseName(file);
	const usedNames = new Set(fonts.map((font) => font.name.toLowerCase()));
	let fontName = baseName;
	let suffix = 2;

	while (usedNames.has(fontName.toLowerCase())) {
		fontName = `${baseName} ${suffix}`;
		suffix += 1;
	}

	return fontName;
}

function cssFontFamilyName(fontName: string) {
	return `"${fontName.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

function clampFontSizePoints(value: number) {
	if (!Number.isFinite(value)) {
		return DEFAULT_FONT_SIZE_POINTS;
	}

	return Math.min(MAX_FONT_SIZE_POINTS, Math.max(MIN_FONT_SIZE_POINTS, Math.round(value * 2) / 2));
}

function fontSizePointsToHalfPoints(value: number) {
	return Math.round(clampFontSizePoints(value) * 2);
}

function parseFontSizePoints(value: string | null | undefined) {
	const match = value?.replace(',', '.').match(/\d+(?:\.\d+)?/);
	if (!match) {
		return DEFAULT_FONT_SIZE_POINTS;
	}

	return clampFontSizePoints(Number(match[0]));
}

function getFontSizeControl(button: HTMLElement) {
	return button.parentElement;
}

function readFontSizeControlPoints(control: HTMLElement | null) {
	const input = control?.querySelector<HTMLInputElement>('[data-testid="font-size-input"]');
	if (input) {
		return parseFontSizePoints(input.value);
	}

	const display = control?.querySelector<HTMLElement>('[data-testid="font-size-display"]');
	return parseFontSizePoints(display?.textContent);
}

function updateFontSizeControlDisplay(control: HTMLElement | null, value: number) {
	const nextText = String(clampFontSizePoints(value));
	const input = control?.querySelector<HTMLInputElement>('[data-testid="font-size-input"]');
	if (input) {
		input.value = nextText;
	}

	const display = control?.querySelector<HTMLElement>('[data-testid="font-size-display"]');
	if (display) {
		display.textContent = nextText;
	}
}

function getFontSizeStepTarget(target: EventTarget | null) {
	if (!(target instanceof Element)) {
		return null;
	}

	const button = target.closest<HTMLButtonElement>('[data-testid="font-size-decrease"], [data-testid="font-size-increase"]');
	if (!button) {
		return null;
	}

	const direction: FontSizeStepDirection = button.dataset.testid === 'font-size-decrease' ? -1 : 1;
	return { button, direction };
}

function getToolbarTooltipTarget(target: EventTarget | null, editorRoot: HTMLElement | null) {
	if (!(target instanceof Element) || !editorRoot) {
		return null;
	}

	const candidate = target.closest<HTMLElement>('button, [role="button"]');
	if (!candidate || !editorRoot.contains(candidate)) {
		return null;
	}

	const excludedToolbarRow = candidate.closest(
		'[data-docxidian-no-toolbar-tooltip], [data-testid="title-bar"], [role="menubar"]'
	);
	if (excludedToolbarRow && editorRoot.contains(excludedToolbarRow)) {
		return null;
	}

	const toolbar = candidate.closest('[data-testid="editor-toolbar"], [data-testid="formatting-bar"], .docx-table-toolbar');
	if (!toolbar || !editorRoot.contains(toolbar)) {
		return null;
	}

	return candidate;
}

function getToolbarTooltipText(target: HTMLElement) {
	return (
		target.getAttribute('aria-label')
		|| target.dataset.docxidianNativeTitle
		|| target.getAttribute('title')
		|| target.textContent
		|| ''
	).replace(/\s+/g, ' ').trim();
}

function suspendNativeTitle(target: HTMLElement) {
	const title = target.getAttribute('title');
	if (!title || target.dataset.docxidianTooltipTitle !== undefined) {
		return;
	}

	target.dataset.docxidianTooltipTitle = title;
	target.removeAttribute('title');
}

function restoreNativeTitle(target: HTMLElement | null) {
	if (!target || target.dataset.docxidianTooltipTitle === undefined) {
		return;
	}

	target.setAttribute('title', target.dataset.docxidianTooltipTitle);
	delete target.dataset.docxidianTooltipTitle;
}

function markLightMenuSurface(surface: HTMLElement, className: string) {
	surface.classList.add('docxidian-light-menu-surface', className);

	const shell = surface.closest<HTMLElement>('[data-radix-popper-content-wrapper], [style*="position: fixed"], [style*="position: absolute"]');
	if (shell) {
		shell.classList.add('docxidian-light-menu-shell');
	}
}

function normalizeEditorFloatingLayers(editorRoot: HTMLElement) {
	activeDocument.querySelectorAll<HTMLElement>('div[style*="position: fixed"]').forEach((layer) => {
		layer.classList.toggle('docxidian-fixed-dialog-layer', Boolean(layer.querySelector(':scope > [role="dialog"]')));
	});

	const hasFloatingMenu = Boolean(
		editorRoot.querySelector('[role="menubar"] [style*="position: fixed"], [role="menubar"] [style*="position: absolute"]'),
	);
	editorRoot.classList.toggle('docxidian-has-floating-menu', hasFloatingMenu);
}

function isFontDropdownListbox(listbox: HTMLElement) {
	const optionLabels = Array.from(listbox.querySelectorAll<HTMLElement>('[role="option"]'))
		.map((option) => option.textContent?.replace(/\s+/g, ' ').trim().toLowerCase() ?? '')
		.filter(Boolean);

	return optionLabels.includes('arial')
		&& optionLabels.includes('times new roman')
		&& optionLabels.includes('courier new');
}

function appendImportFontOption(listbox: HTMLElement, onImportFont: () => void) {
	markLightMenuSurface(listbox, 'docxidian-font-listbox');

	if (listbox.querySelector('[data-docxidian-import-font-option]')) {
		return;
	}

	const separator = activeDocument.createElement('div');
	separator.className = 'docxidian-import-font-separator';
	separator.setAttribute('role', 'separator');
	separator.dataset.docxidianImportFontOption = 'true';

	const button = activeDocument.createElement('button');
	button.type = 'button';
	button.className = 'docxidian-import-font-option';
	button.dataset.docxidianImportFontOption = 'true';
	button.setAttribute('role', 'option');
	button.setAttribute('aria-selected', 'false');
	button.textContent = IMPORT_FONT_MENU_LABEL;

	const openImporter = (evt: Event) => {
		evt.preventDefault();
		evt.stopImmediatePropagation();
		evt.stopPropagation();
		activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		onImportFont();
	};

	button.addEventListener('pointerdown', (evt) => {
		evt.preventDefault();
		evt.stopImmediatePropagation();
		evt.stopPropagation();
	});
	button.addEventListener('mousedown', (evt) => {
		evt.preventDefault();
		evt.stopImmediatePropagation();
		evt.stopPropagation();
	});
	button.addEventListener('click', openImporter);
	button.addEventListener('keydown', (evt) => {
		if (evt.key === 'Enter' || evt.key === ' ') {
			openImporter(evt);
		}
	});

	listbox.append(separator, button);
}

function clampCustomTableSize(value: number) {
	if (!Number.isFinite(value)) {
		return MIN_CUSTOM_TABLE_SIZE;
	}

	return Math.min(MAX_CUSTOM_TABLE_SIZE, Math.max(MIN_CUSTOM_TABLE_SIZE, Math.round(value)));
}

function isTableSizeGrid(grid: HTMLElement) {
	const gridCells = grid.querySelectorAll('[role="gridcell"]');
	if (gridCells.length < 9) {
		return false;
	}

	const label = grid.getAttribute('aria-label')?.toLowerCase() ?? '';
	return label.includes('table') || gridCells.length === 36;
}

function appendCustomTableOption(grid: HTMLElement, onCustomTable: () => void) {
	const container = grid.parentElement;
	if (!container || container.querySelector('[data-docxidian-custom-table-option]')) {
		return;
	}

	markLightMenuSurface(container, 'docxidian-table-size-menu');

	const separator = activeDocument.createElement('div');
	separator.className = 'docxidian-custom-table-separator';
	separator.setAttribute('role', 'separator');
	separator.dataset.docxidianCustomTableOption = 'true';

	const button = activeDocument.createElement('button');
	button.type = 'button';
	button.className = 'docxidian-custom-table-option';
	button.dataset.docxidianCustomTableOption = 'true';
	button.textContent = CUSTOM_TABLE_MENU_LABEL;

	const openCustomTable = (evt: Event) => {
		evt.preventDefault();
		evt.stopImmediatePropagation();
		evt.stopPropagation();
		activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		onCustomTable();
	};

	button.addEventListener('pointerdown', (evt) => {
		evt.preventDefault();
		evt.stopImmediatePropagation();
		evt.stopPropagation();
	});
	button.addEventListener('mousedown', (evt) => {
		evt.preventDefault();
		evt.stopImmediatePropagation();
		evt.stopPropagation();
	});
	button.addEventListener('click', openCustomTable);
	button.addEventListener('keydown', (evt) => {
		if (evt.key === 'Enter' || evt.key === ' ') {
			openCustomTable(evt);
		}
	});

	container.append(separator, button);
}

function getPlainTextFromKeyboardEvent(event: KeyboardEvent) {
	if (event.metaKey || event.ctrlKey || event.altKey || event.isComposing) {
		return null;
	}

	if (event.key === ' ' || event.key === 'Spacebar') {
		return ' ';
	}

	return event.key.length === 1 ? event.key : null;
}

function getPlainTextFromInputEvent(event: InputEvent) {
	if (event.isComposing || event.inputType !== 'insertText' || !event.data || /[\r\n]/.test(event.data)) {
		return null;
	}

	return event.data;
}

const preserveTypedSpacePlugin = new Plugin({
	key: preserveTypedSpacePluginKey,
	props: {
		handleDOMEvents: {
			beforeinput(view, event) {
				if (!(event instanceof InputEvent)) {
					return false;
				}

				const text = getPlainTextFromInputEvent(event);
				if (!text) {
					return false;
				}

				event.preventDefault();
				return insertPlainTypedText(view, text);
			},
			paste(view, event) {
				if (!(event instanceof ClipboardEvent)) {
					return false;
				}

				const marker = getCurrentParagraphListMarker(view);
				const strippedText = stripMatchingListMarkerPrefixFromPastedText(
					getPlainTextFromClipboardEvent(event),
					marker,
				);
				if (strippedText === null) {
					return false;
				}

				event.preventDefault();
				return insertPlainTypedText(view, strippedText);
			},
		},
		handleKeyDown(view, event) {
			const text = getPlainTextFromKeyboardEvent(event);
			if (!text) {
				return false;
			}

			event.preventDefault();
			return insertPlainTypedText(view, text);
		},
		handleTextInput(view, from, to, text) {
			if (!text || /[\r\n]/.test(text)) {
				return false;
			}

			return insertPlainTypedText(view, text, from, to);
		},
	},
});

function clampZoom(zoom: number) {
	return Math.max(MIN_TOUCH_ZOOM, Math.min(MAX_TOUCH_ZOOM, Math.round(zoom * 100) / 100));
}

function scaleTouchZoom(startZoom: number, rawScale: number) {
	if (!Number.isFinite(rawScale) || rawScale <= 0) {
		return startZoom;
	}

	return clampZoom(startZoom * Math.pow(rawScale, TOUCH_ZOOM_SENSITIVITY));
}

function getTouchDistance(first: Touch, second: Touch) {
	return Math.hypot(first.clientX - second.clientX, first.clientY - second.clientY);
}

function getTouchCenter(first: Touch, second: Touch) {
	return {
		x: (first.clientX + second.clientX) / 2,
		y: (first.clientY + second.clientY) / 2,
	};
}

function getPointDistance(first: PointerPoint, second: PointerPoint) {
	return Math.hypot(first.x - second.x, first.y - second.y);
}

function getPointCenter(first: PointerPoint, second: PointerPoint) {
	return {
		x: (first.x + second.x) / 2,
		y: (first.y + second.y) / 2,
	};
}

function getScrollableEditorElement(root: HTMLElement) {
	const pages = root.querySelector<HTMLElement>('.paged-editor__pages');
	let candidate: HTMLElement | null = pages?.parentElement ?? root;

	while (candidate && candidate !== root) {
		const style = window.getComputedStyle(candidate);
		const canScroll = /(auto|scroll)/.test(`${style.overflow} ${style.overflowX} ${style.overflowY}`);
		if (canScroll && (candidate.scrollHeight > candidate.clientHeight || candidate.scrollWidth > candidate.clientWidth)) {
			return candidate;
		}
		candidate = candidate.parentElement;
	}

	return root;
}

function centerEditorViewport(root: HTMLElement) {
	const pages = root.querySelector<HTMLElement>('.paged-editor__pages');
	if (!pages) {
		return false;
	}

	const scrollContainer = getScrollableEditorElement(root);
	const scrollRect = scrollContainer.getBoundingClientRect();
	const pagesRect = pages.getBoundingClientRect();
	if (scrollRect.width <= 0 || pagesRect.width <= 0) {
		return false;
	}

	const pagesLeftInScrollSpace = pagesRect.left - scrollRect.left + scrollContainer.scrollLeft;
	const pagesCenter = pagesLeftInScrollSpace + (pagesRect.width / 2);
	const nextScrollLeft = Math.max(0, pagesCenter - (scrollContainer.clientWidth / 2));
	if (!Number.isFinite(nextScrollLeft)) {
		return false;
	}

	scrollContainer.scrollLeft = nextScrollLeft;
	return true;
}

function shouldEnableTouchPinchZoom() {
	if (typeof window === 'undefined' || typeof navigator === 'undefined') {
		return false;
	}

	return Platform.isMobile || Platform.isMobileApp || (navigator.maxTouchPoints >= 2 && window.matchMedia('(hover: none)').matches);
}

function getEditorModeFromButton(button: HTMLButtonElement): EditorMode | null {
	const label = button.textContent?.replace(/\s+/g, ' ').trim().toLowerCase() ?? '';

	if (label.startsWith('bearbeiten') || label.startsWith('edit')) {
		return 'editing';
	}
	if (label.startsWith('vorschlagen') || label.startsWith('suggest')) {
		return 'suggesting';
	}
	if (label.startsWith('anzeigen') || label.startsWith('view')) {
		return 'viewing';
	}

	return null;
}

function escapeRegExp(value: string) {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function createFindPattern(searchText: string, matchCase: boolean, wholeWord: boolean) {
	if (!searchText.trim()) {
		return null;
	}

	const source = wholeWord ? `\\b${escapeRegExp(searchText)}\\b` : escapeRegExp(searchText);
	return new RegExp(source, matchCase ? 'g' : 'gi');
}

function findMatchesInView(editor: DocxEditorRef | null, searchText: string, matchCase: boolean, wholeWord: boolean) {
	const view = editor?.getEditorRef()?.getView();
	const pattern = createFindPattern(searchText, matchCase, wholeWord);
	const matches: FindMatch[] = [];

	if (!view || !pattern) {
		return matches;
	}

	view.state.doc.descendants((node, pos) => {
		if (!node.isTextblock) {
			return true;
		}

		const text = node.textContent;
		for (const match of text.matchAll(pattern)) {
			const index = match.index ?? -1;
			if (index < 0) {
				continue;
			}

			matches.push({
				from: pos + 1 + index,
				to: pos + 1 + index + match[0].length,
				text: match[0],
			});
		}

		return false;
	});

	return matches;
}

function getVisibleListMarker(attrs: Record<string, unknown>) {
	if (attrs.listMarkerHidden) {
		return '';
	}

	const marker = attrs.listMarker;
	return typeof marker === 'string' ? marker.trimEnd() : '';
}

function getCurrentParagraphListMarker(view: EditorView) {
	const { $from } = view.state.selection;
	for (let depth = $from.depth; depth >= 0; depth -= 1) {
		const node = $from.node(depth);
		if (node.isTextblock) {
			return getVisibleListMarker(node.attrs);
		}
	}

	return '';
}

function stripMatchingListMarkerPrefixFromPastedText(text: string, marker: string) {
	if (!marker || !text) {
		return null;
	}

	const singleLineMatch = text.match(/^([^\r\n]*)(?:\r\n|\n|\r)?$/);
	if (!singleLineMatch) {
		return null;
	}

	const line = singleLineMatch[1] ?? '';
	const markerPrefixPattern = new RegExp(`^${escapeRegExp(marker)}(?:\\t|\\s)+`);
	if (!markerPrefixPattern.test(line)) {
		return null;
	}

	return line.replace(markerPrefixPattern, '');
}

function getPlainTextFromClipboardEvent(event: ClipboardEvent) {
	return event.clipboardData?.getData('text/plain') ?? event.clipboardData?.getData('Text') ?? '';
}

function clearListMarkerSelectionHighlights(root: HTMLElement) {
	root.querySelectorAll<HTMLElement>(`.${SELECTED_LIST_MARKER_CLASS}`).forEach((marker) => {
		marker.classList.remove(SELECTED_LIST_MARKER_CLASS);
	});
}

function getCssGeneratedContentText(element: HTMLElement) {
	const generatedContent = window.getComputedStyle(element, '::before').content;
	if (!generatedContent || generatedContent === 'none' || generatedContent === 'normal') {
		return '';
	}

	if (
		(generatedContent.startsWith('"') && generatedContent.endsWith('"'))
		|| (generatedContent.startsWith("'") && generatedContent.endsWith("'"))
	) {
		return generatedContent.slice(1, -1).replace(/\\A/g, '\n').replace(/\\"/g, '"').replace(/\\'/g, "'");
	}

	return generatedContent;
}

function getRenderedListMarkerText(marker: HTMLElement) {
	const text = marker.textContent?.trimEnd() ?? '';
	if (text.trim().length > 0) {
		return text;
	}

	const generatedContentText = getCssGeneratedContentText(marker).trimEnd();
	return generatedContentText.trim().length > 0 ? generatedContentText : '';
}

function rangeIntersectsNode(range: Range, node: Node) {
	try {
		return range.intersectsNode(node);
	} catch {
		return false;
	}
}

function getSelectedRenderedListMarkers(root: HTMLElement) {
	const selectedMarkers = new Map<number, string>();
	const browserSelection = window.getSelection();
	if (!browserSelection || browserSelection.rangeCount === 0) {
		return selectedMarkers;
	}

	const ranges: Range[] = [];
	for (let index = 0; index < browserSelection.rangeCount; index += 1) {
		ranges.push(browserSelection.getRangeAt(index));
	}

	root.querySelectorAll<HTMLElement>(LIST_PARAGRAPH_SELECTOR).forEach((paragraph) => {
		const marker = paragraph.querySelector<HTMLElement>(LIST_MARKER_SELECTOR);
		if (!marker) {
			return;
		}

		const markerText = getRenderedListMarkerText(marker);
		if (markerText.length === 0) {
			return;
		}

		const markerSelected = marker.classList.contains(SELECTED_LIST_MARKER_CLASS)
			|| ranges.some((range) => rangeIntersectsNode(range, marker));
		if (!markerSelected) {
			return;
		}

		const paragraphStart = Number(paragraph.dataset.pmStart);
		if (!Number.isFinite(paragraphStart)) {
			return;
		}

		selectedMarkers.set(paragraphStart, markerText);
		selectedMarkers.set(paragraphStart + 1, markerText);
	});

	return selectedMarkers;
}

function updateListMarkerSelectionHighlights(root: HTMLElement, view: EditorView) {
	clearListMarkerSelectionHighlights(root);

	const { selection } = view.state;
	if (selection.empty) {
		return;
	}

	const selectionStart = Math.min(selection.from, selection.to);
	const selectionEnd = Math.max(selection.from, selection.to);
	root.querySelectorAll<HTMLElement>(LIST_PARAGRAPH_SELECTOR).forEach((paragraph) => {
		const marker = paragraph.querySelector<HTMLElement>(LIST_MARKER_SELECTOR);
		if (!marker) {
			return;
		}

		const paragraphStart = Number(paragraph.dataset.pmStart);
		if (!Number.isFinite(paragraphStart)) {
			return;
		}

		if (selectionStart <= paragraphStart + 1 && selectionEnd > paragraphStart) {
			marker.classList.add(SELECTED_LIST_MARKER_CLASS);
		}
	});
}

function splitClipboardLines(text: string) {
	return text.split(/\r\n|\n|\r/);
}

function addListMarkersToPlainText(view: EditorView, text: string) {
	if (!text.trim()) {
		return null;
	}

	const selectedLines = splitClipboardLines(text);
	const textToMarker = new Map<string, string>();
	view.state.doc.descendants((node) => {
		if (!node.isTextblock) {
			return true;
		}

		const marker = getVisibleListMarker(node.attrs);
		const paragraphText = node.textContent;
		if (marker.length > 0 && paragraphText.length > 0 && !textToMarker.has(paragraphText)) {
			textToMarker.set(paragraphText, marker);
		}

		return false;
	});

	let includedMarker = false;
	const lines = selectedLines.map((line) => {
		const marker = textToMarker.get(line);
		if (!marker || line.startsWith(marker)) {
			return line;
		}

		includedMarker = true;
		return `${marker}\t${line}`;
	});

	return includedMarker ? lines.join('\n') : null;
}

interface ElectronClipboard {
	readHTML?: () => string;
	readRTF?: () => string;
	readText: () => string;
	write?: (data: { html?: string; rtf?: string; text?: string }) => void;
	writeText: (text: string) => void;
}

function getElectronClipboard() {
	try {
		if (typeof require !== 'function') {
			return null;
		}

		const electron = require('electron') as {
			clipboard?: ElectronClipboard;
		};
		return electron.clipboard ?? null;
	} catch {
		return null;
	}
}

async function readPlainTextClipboard() {
	const electronClipboard = getElectronClipboard();
	if (electronClipboard) {
		return electronClipboard.readText();
	}

	return await navigator.clipboard.readText();
}

function readHtmlClipboard() {
	return getElectronClipboard()?.readHTML?.() ?? '';
}

async function writePlainTextClipboard(text: string) {
	const electronClipboard = getElectronClipboard();
	if (electronClipboard) {
		const html = electronClipboard.readHTML?.() ?? '';
		const rtf = electronClipboard.readRTF?.() ?? '';
		if (electronClipboard.write && (html.length > 0 || rtf.length > 0)) {
			electronClipboard.write({
				text,
				...(html.length > 0 ? { html } : {}),
				...(rtf.length > 0 ? { rtf } : {}),
			});
		} else {
			electronClipboard.writeText(text);
		}
		return;
	}

	await navigator.clipboard.writeText(text);
}

async function rewritePlainTextClipboardWithListMarkers(view: EditorView) {
	try {
		const clipboardText = await readPlainTextClipboard();
		const listAwareText = addListMarkersToPlainText(view, clipboardText);
		if (!listAwareText || listAwareText === clipboardText) {
			return false;
		}

		await writePlainTextClipboard(listAwareText);
		debugLog('clipboard', 'Rewrote clipboard text with list markers', {
			originalTextLength: clipboardText.length,
			listAwareTextLength: listAwareText.length,
		});
		return true;
	} catch (error) {
		debugLog('clipboard', 'Could not rewrite clipboard text with list markers', error);
		return false;
	}
}

function getPasteTextWithListMarkerGuard(view: EditorView, text: string) {
	const marker = getCurrentParagraphListMarker(view);
	return stripMatchingListMarkerPrefixFromPastedText(text, marker) ?? text;
}

async function pasteClipboardIntoEditor(view: EditorView, options: PasteClipboardOptions) {
	view.focus();
	const text = await readPlainTextClipboard();
	if (!options.preserveFormatting) {
		return text ? insertPlainTypedText(view, getPasteTextWithListMarkerGuard(view, text)) : false;
	}

	const html = readHtmlClipboard();
	if (html.trim()) {
		try {
			const clipboardData = new DataTransfer();
			clipboardData.setData('text/html', html);
			clipboardData.setData('text/plain', text);
			const pasteEvent = new ClipboardEvent('paste', {
				bubbles: true,
				cancelable: true,
				clipboardData,
			});
			const defaultAllowed = view.dom.dispatchEvent(pasteEvent);
			if (pasteEvent.defaultPrevented || !defaultAllowed) {
				return true;
			}
		} catch (error) {
			debugLog('clipboard', 'Could not dispatch formatted paste event; falling back to plain text paste', error);
		}
	}

	return text ? insertPlainTypedText(view, getPasteTextWithListMarkerGuard(view, text)) : false;
}

export function ensureEditorStyles() {
	if (stylesInjected) {
		return;
	}

	const styleSheet = new CSSStyleSheet();
	styleSheet.replaceSync(editorStyles);
	activeDocument.adoptedStyleSheets = [...activeDocument.adoptedStyleSheets, styleSheet];
	stylesInjected = true;
}

const SaveButton = ({ onClick }: { onClick: () => void }) => {
	const ref = useRef<HTMLButtonElement>(null);
	useEffect(() => {
		if (ref.current) {
			ref.current.replaceChildren();
			setIcon(ref.current, 'save');
		}
	}, []);

	return (
		<button
			ref={ref}
			type="button"
			className="clickable-icon docxidian-logo-save-button"
			onClick={onClick}
			aria-label="Save"
			style={{
				background: 'transparent',
				border: 'none',
				boxShadow: 'none',
				padding: '4px 8px',
				cursor: 'pointer',
				display: 'flex',
				alignItems: 'center',
				justifyContent: 'center',
				color: 'inherit'
			}}
		/>
	);
};

const SAVE_STATUS_LABELS: Record<SaveStatus, string> = {
	saved: 'Saved',
	saving: 'Saving...',
	unsaved: 'Unsaved',
	failed: 'Save failed',
};

const SaveStatusIndicator = ({ status }: { status: SaveStatus }) => (
	<span
		className={`docxidian-save-status docxidian-save-status-${status}`}
		role="status"
		aria-live="polite"
		title={SAVE_STATUS_LABELS[status]}
	>
		<span className="docxidian-save-status-dot" aria-hidden="true" />
		{SAVE_STATUS_LABELS[status]}
	</span>
);

interface FindReplaceDialogProps {
	isOpen: boolean;
	labels: FindReplaceLabels;
	mode: FindReplaceMode;
	searchText: string;
	replaceText: string;
	matchCase: boolean;
	wholeWord: boolean;
	matchCount: number;
	currentIndex: number;
	onSearchTextChange: (value: string) => void;
	onReplaceTextChange: (value: string) => void;
	onMatchCaseChange: (value: boolean) => void;
	onWholeWordChange: (value: boolean) => void;
	onModeChange: (mode: FindReplaceMode) => void;
	onNext: () => void;
	onPrevious: () => void;
	onReplace: () => void;
	onReplaceAll: () => void;
	onClose: () => void;
}

const FindReplaceDialog = ({
	isOpen,
	labels,
	mode,
	searchText,
	replaceText,
	matchCase,
	wholeWord,
	matchCount,
	currentIndex,
	onSearchTextChange,
	onReplaceTextChange,
	onMatchCaseChange,
	onWholeWordChange,
	onModeChange,
	onNext,
	onPrevious,
	onReplace,
	onReplaceAll,
	onClose,
}: FindReplaceDialogProps) => {
	if (!isOpen) {
		return null;
	}

	const resultText = searchText.trim()
		? (matchCount > 0 ? labels.resultCount(currentIndex + 1, matchCount) : labels.noMatches)
		: '';

	return (
		<div
			className="docxidian-find-dialog"
			style={{
				position: 'fixed',
				right: '24px',
				top: '92px',
				zIndex: 100050,
				width: '360px',
				background: 'white',
				border: '1px solid var(--background-modifier-border, #d1d5db)',
				borderRadius: '8px',
				boxShadow: '0 10px 30px rgba(0, 0, 0, 0.18)',
				padding: '12px',
				color: 'var(--text-normal, #202124)',
			}}
		>
			<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '10px' }}>
				<strong>{mode === 'replace' ? labels.findAndReplace : labels.find}</strong>
				<button type="button" aria-label={labels.close} onClick={onClose} style={{ border: 'none', background: 'transparent', cursor: 'pointer' }}>×</button>
			</div>
			<div style={{ display: 'flex', gap: '6px', marginBottom: '8px' }}>
				<input
					value={searchText}
					onChange={(evt) => onSearchTextChange(evt.currentTarget.value)}
					placeholder={labels.findText}
					autoFocus
					style={{ flex: 1, height: '30px' }}
					onKeyDown={(evt) => {
						if (evt.key === 'Enter') {
							evt.preventDefault();
							evt.shiftKey ? onPrevious() : onNext();
						}
					}}
				/>
				<button type="button" aria-label={labels.previous} title={labels.previous} onClick={onPrevious} disabled={matchCount === 0}>↑</button>
				<button type="button" aria-label={labels.next} title={labels.next} onClick={onNext} disabled={matchCount === 0}>↓</button>
			</div>
			{mode === 'replace' && (
				<div style={{ display: 'flex', gap: '6px', marginBottom: '8px' }}>
					<input
						value={replaceText}
						onChange={(evt) => onReplaceTextChange(evt.currentTarget.value)}
						placeholder={labels.replaceWith}
						style={{ flex: 1, height: '30px' }}
					/>
					<button type="button" onClick={onReplace} disabled={matchCount === 0}>{labels.replace}</button>
					<button type="button" onClick={onReplaceAll} disabled={matchCount === 0}>{labels.replaceAll}</button>
				</div>
			)}
			<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px' }}>
				<div style={{ display: 'flex', gap: '10px', fontSize: '12px' }}>
					<label><input type="checkbox" checked={matchCase} onChange={(evt) => onMatchCaseChange(evt.currentTarget.checked)} /> {labels.matchCase}</label>
					<label><input type="checkbox" checked={wholeWord} onChange={(evt) => onWholeWordChange(evt.currentTarget.checked)} /> {labels.wholeWords}</label>
				</div>
				<div style={{ fontSize: '12px', color: 'var(--text-muted, #6b7280)', whiteSpace: 'nowrap' }}>{resultText}</div>
			</div>
			{mode === 'find' && (
				<button type="button" onClick={() => onModeChange('replace')} style={{ marginTop: '10px' }}>{labels.showReplace}</button>
			)}
		</div>
	);
};

interface CustomTableDialogProps {
	isOpen: boolean;
	onClose: () => void;
	onInsert: (rows: number, columns: number) => void;
}

const CustomTableDialog = ({ isOpen, onClose, onInsert }: CustomTableDialogProps) => {
	const [rows, setRows] = useState(3);
	const [columns, setColumns] = useState(3);

	useEffect(() => {
		if (isOpen) {
			setRows(3);
			setColumns(3);
		}
	}, [isOpen]);

	const submit = useCallback(() => {
		onInsert(clampCustomTableSize(rows), clampCustomTableSize(columns));
	}, [columns, onInsert, rows]);

	if (!isOpen) {
		return null;
	}

	return (
		<div
			className="docxidian-custom-table-backdrop"
			onMouseDown={(evt) => {
				if (evt.target === evt.currentTarget) {
					onClose();
				}
			}}
			onKeyDown={(evt) => {
				if (evt.key === 'Escape') {
					evt.preventDefault();
					onClose();
				}
				if (evt.key === 'Enter') {
					evt.preventDefault();
					submit();
				}
			}}
		>
			<div
				className="docxidian-custom-table-dialog"
				role="dialog"
				aria-modal="true"
				aria-label="Custom table"
				onMouseDown={(evt) => evt.stopPropagation()}
			>
				<div className="docxidian-custom-table-title">Custom table</div>
				<label className="docxidian-custom-table-field">
					<span>Rows</span>
					<input
						type="number"
						min={MIN_CUSTOM_TABLE_SIZE}
						max={MAX_CUSTOM_TABLE_SIZE}
						value={rows}
						autoFocus
						onChange={(evt) => setRows(clampCustomTableSize(Number(evt.currentTarget.value)))}
					/>
				</label>
				<label className="docxidian-custom-table-field">
					<span>Columns</span>
					<input
						type="number"
						min={MIN_CUSTOM_TABLE_SIZE}
						max={MAX_CUSTOM_TABLE_SIZE}
						value={columns}
						onChange={(evt) => setColumns(clampCustomTableSize(Number(evt.currentTarget.value)))}
					/>
				</label>
				<div className="docxidian-custom-table-actions">
					<button type="button" onClick={onClose}>Cancel</button>
					<button type="button" className="docxidian-custom-table-primary" onClick={submit}>Insert</button>
				</div>
			</div>
		</div>
	);
};

export interface DocxReactViewProps {
	file: TFile | null;
	buffer: ArrayBuffer | null;
	documentKey: string;
	error: string | null;
	isLoading: boolean;
	authorName: string;
	i18n: Translations | undefined;
	showRuler: boolean;
	autosave: boolean;
	defaultZoom: number;
	reserveReviewSidebar: boolean;
	onDirtyChange: (isDirty: boolean) => void;
	onSave: (buffer: ArrayBuffer) => Promise<void>;
	onDocumentNameChange: (name: string) => Promise<void>;
}

export interface DocxReactViewHandle {
	save: () => Promise<boolean>;
	exportBuffer: (options?: ExportDocumentBufferOptions) => Promise<ArrayBuffer | null>;
	exportRenderedPdf: () => Promise<ArrayBuffer | null>;
	pasteFromClipboard: (options: PasteClipboardOptions) => Promise<boolean>;
	rewriteClipboardTextWithListMarkers: () => Promise<boolean>;
	openFind: () => void;
	openFindReplace: () => void;
	openImagePicker: () => void;
	openCustomTableDialog: () => void;
	openFontPicker: () => void;
	setMode: (mode: EditorMode) => void;
	setZoom: (zoom: number) => void;
}

export const DocxReactView = forwardRef<DocxReactViewHandle, DocxReactViewProps>(function DocxReactView(
	{ file, buffer, documentKey, error, isLoading, authorName, i18n, showRuler, autosave, defaultZoom, reserveReviewSidebar, onDirtyChange, onSave, onDocumentNameChange },
	ref,
) {
	const editorRef = useRef<DocxEditorRef>(null);
	const renderedDomContextRef = useRef<RenderedDomContext | null>(null);
	const imageInputRef = useRef<HTMLInputElement>(null);
	const fontInputRef = useRef<HTMLInputElement>(null);
	const editorClassNameRef = useRef(`docxidian-editor-${++editorInstanceCounter}`);
	const rulerSyncFrameRef = useRef<number | null>(null);
	const rulerSyncTimeoutRef = useRef<number | null>(null);
	const initialCenterFrameRef = useRef<number | null>(null);
	const initialCenterTimeoutsRef = useRef<number[]>([]);
	const centeredDocumentKeyRef = useRef<string | null>(null);
	const pinchZoomStateRef = useRef<PinchZoomState | null>(null);
	const pinchZoomScrollFrameRef = useRef<number | null>(null);
	const activeTouchPointersRef = useRef<Map<number, PointerPoint>>(new Map());
	const fontSizeHoldRef = useRef<FontSizeHoldState | null>(null);
	const listMarkerSelectionFrameRef = useRef<number | null>(null);
	const dirtyTrackingEnabledRef = useRef(false);
	const dirtyVersionRef = useRef(0);
	const isSavingRef = useRef(false);
	const autosaveTimeoutRef = useRef<number | null>(null);
	const renameTimeoutRef = useRef<number | null>(null);
	const pendingSaveModeRef = useRef<'save' | 'export'>('save');
	const pendingSaveOptionsRef = useRef<SaveDocumentOptions | undefined>(undefined);
	const pendingSavePromiseRef = useRef<Promise<boolean> | null>(null);
	const [documentName, setDocumentName] = useState(file?.name ?? '');
	const [editorMode, setEditorMode] = useState<EditorMode>('editing');
	const [saveStatus, setSaveStatus] = useState<SaveStatus>('saved');
	const [findDialogMode, setFindDialogMode] = useState<FindReplaceMode | null>(null);
	const [findSearchText, setFindSearchText] = useState('');
	const [findReplaceText, setFindReplaceText] = useState('');
	const [findMatchCase, setFindMatchCase] = useState(false);
	const [findWholeWord, setFindWholeWord] = useState(false);
	const [findMatches, setFindMatches] = useState<FindMatch[]>([]);
	const [currentFindIndex, setCurrentFindIndex] = useState(0);
	const [importedFonts, setImportedFonts] = useState<FontOption[]>([]);
	const [customTableDialogOpen, setCustomTableDialogOpen] = useState(false);
	const filePath = file?.path ?? null;
	const findReplaceLabels = useMemo(() => createFindReplaceLabels(i18n), [i18n]);
	const fontFamilies = useMemo<FontOption[]>(() => [
		...DEFAULT_EDITOR_FONT_FAMILIES,
		...importedFonts,
	], [importedFonts]);
	const findHighlightPlugin = useMemo(() => new Plugin<FindHighlightState>({
		key: findHighlightPluginKey,
		state: {
			init: () => ({ matches: [], currentIndex: 0 }),
			apply: (transaction, previous) => transaction.getMeta(findHighlightPluginKey) ?? previous,
		},
		props: {
			decorations: (state) => {
				const pluginState = findHighlightPluginKey.getState(state);
				if (!pluginState || pluginState.matches.length === 0) {
					return DecorationSet.empty;
				}

				return DecorationSet.create(
					state.doc,
					pluginState.matches.map((match, index) => Decoration.inline(
						match.from,
						match.to,
						{ class: index === pluginState.currentIndex ? 'docxidian-find-current' : 'docxidian-find-match' },
					)),
				);
			},
		},
	}), []);
	const externalPlugins = useMemo(() => [preserveTypedSpacePlugin, findHighlightPlugin], [findHighlightPlugin]);
	const pluginSidebarItems = useMemo<NonNullable<ComponentProps<typeof DocxEditor>['pluginSidebarItems']>>(() => {
		if (!reserveReviewSidebar) {
			return [];
		}

		return [{
			id: 'docxidian-review-sidebar-reservation',
			anchorPos: 1,
			estimatedHeight: 1,
			priority: Number.MAX_SAFE_INTEGER,
			render: () => null,
		}];
	}, [reserveReviewSidebar]);

	useEffect(() => {
		ensureEditorStyles();
	}, []);

	const syncListMarkerSelectionHighlights = useCallback(() => {
		const root = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		const view = editorRef.current?.getEditorRef()?.getView();
		if (!root || !view) {
			return;
		}

		updateListMarkerSelectionHighlights(root, view);
	}, []);

	const scheduleListMarkerSelectionHighlightSync = useCallback(() => {
		if (listMarkerSelectionFrameRef.current !== null) {
			window.cancelAnimationFrame(listMarkerSelectionFrameRef.current);
		}

		listMarkerSelectionFrameRef.current = window.requestAnimationFrame(() => {
			listMarkerSelectionFrameRef.current = null;
			syncListMarkerSelectionHighlights();
		});
	}, [syncListMarkerSelectionHighlights]);

	useEffect(() => () => {
		if (listMarkerSelectionFrameRef.current !== null) {
			window.cancelAnimationFrame(listMarkerSelectionFrameRef.current);
			listMarkerSelectionFrameRef.current = null;
		}
	}, []);

	useEffect(() => {
		dirtyTrackingEnabledRef.current = false;
		const timeout = window.setTimeout(() => {
			dirtyTrackingEnabledRef.current = true;
		}, 500);

		return () => {
			window.clearTimeout(timeout);
			dirtyTrackingEnabledRef.current = false;
		};
	}, [file, buffer]);

	useEffect(() => {
		setDocumentName(file?.name ?? '');
		setEditorMode('editing');
		setFindDialogMode(null);
		setFindSearchText('');
		setFindReplaceText('');
		setFindMatches([]);
		setCurrentFindIndex(0);
		dirtyVersionRef.current = 0;
		setSaveStatus('saved');
	}, [filePath]);

	const setMode = useCallback((mode: EditorMode) => {
		setEditorMode(mode);
	}, []);

	const publishFindHighlights = useCallback((matches: FindMatch[], currentIndex: number) => {
		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view) {
			return;
		}

		view.dispatch(view.state.tr.setMeta(findHighlightPluginKey, { matches, currentIndex }));
	}, []);

	const selectFindMatch = useCallback((matches: FindMatch[], index: number) => {
		const view = editorRef.current?.getEditorRef()?.getView();
		const match = matches[index];
		if (!view || !match) {
			return;
		}

		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, match.from, match.to)).scrollIntoView());
		editorRef.current?.scrollToPosition(match.from);
	}, []);

	const refreshFindMatches = useCallback((searchText: string, matchCase = findMatchCase, wholeWord = findWholeWord, preferredIndex = 0, options: RefreshFindOptions = {}) => {
		const matches = findMatchesInView(editorRef.current, searchText, matchCase, wholeWord);
		const nextIndex = matches.length > 0 ? Math.max(0, Math.min(preferredIndex, matches.length - 1)) : 0;

		setFindMatches(matches);
		setCurrentFindIndex(nextIndex);
		publishFindHighlights(matches, nextIndex);
		if (options.select && matches.length > 0) {
			selectFindMatch(matches, nextIndex);
		}

		return matches;
	}, [findMatchCase, findWholeWord, publishFindHighlights, selectFindMatch]);

	const openFindReplacePanel = useCallback((mode: FindReplaceMode) => {
		const selectedText = editorRef.current?.getSelectionInfo()?.selectedText?.trim();
		const nextSearchText = selectedText || findSearchText;

		setFindDialogMode(mode);
		if (selectedText) {
			setFindSearchText(selectedText);
		}
		refreshFindMatches(nextSearchText);
	}, [findSearchText, refreshFindMatches]);

	const openFindReplaceDialog = useCallback((mode: FindReplaceMode) => {
		openFindReplacePanel(mode);
	}, [openFindReplacePanel]);

	useEffect(() => {
		const handleFindShortcut = (evt: KeyboardEvent) => {
			const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			const key = evt.key.toLowerCase();
			const isFindShortcut = isPrimaryShortcut(evt, 'f');
			const isReplaceShortcut = !Platform.isMacOS && key === 'h' && evt.ctrlKey && !evt.metaKey && !evt.altKey && !evt.shiftKey;
			if (
				!editorRoot
				|| (!isFindShortcut && !isReplaceShortcut)
				|| !(evt.target instanceof Node)
				|| (!editorRoot.contains(evt.target) && !activeDocument.querySelector('.docxidian-find-dialog')?.contains(evt.target))
			) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			openFindReplaceDialog(isFindShortcut ? 'find' : 'replace');
		};

		activeDocument.addEventListener('keydown', handleFindShortcut, true);
		return () => activeDocument.removeEventListener('keydown', handleFindShortcut, true);
	}, [openFindReplaceDialog]);

	const moveFindMatch = useCallback((direction: 1 | -1) => {
		if (findMatches.length === 0) {
			return;
		}

		const nextIndex = (currentFindIndex + direction + findMatches.length) % findMatches.length;
		setCurrentFindIndex(nextIndex);
		publishFindHighlights(findMatches, nextIndex);
		selectFindMatch(findMatches, nextIndex);
	}, [currentFindIndex, findMatches, publishFindHighlights, selectFindMatch]);

	const replaceCurrentMatch = useCallback(() => {
		const view = editorRef.current?.getEditorRef()?.getView();
		const match = findMatches[currentFindIndex];
		if (!view || !match) {
			return;
		}

		const textNode = findReplaceText ? view.state.schema.text(findReplaceText) : null;
		view.dispatch(view.state.tr.replaceWith(match.from, match.to, textNode ? [textNode] : []).scrollIntoView());
		refreshFindMatches(findSearchText, findMatchCase, findWholeWord, currentFindIndex);
	}, [currentFindIndex, findMatchCase, findMatches, findReplaceText, findSearchText, findWholeWord, refreshFindMatches]);

	const replaceAllMatches = useCallback(() => {
		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view || findMatches.length === 0) {
			return;
		}

		let transaction = view.state.tr;
		for (const match of [...findMatches].sort((a, b) => b.from - a.from)) {
			const textNode = findReplaceText ? view.state.schema.text(findReplaceText) : null;
			transaction = transaction.replaceWith(match.from, match.to, textNode ? [textNode] : []);
		}
		view.dispatch(transaction.scrollIntoView());
		refreshFindMatches(findSearchText, findMatchCase, findWholeWord, 0);
	}, [findMatchCase, findMatches, findReplaceText, findSearchText, findWholeWord, refreshFindMatches]);

	const normalizeEditorModeDropdown = useCallback(() => {
		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		normalizeEditorFloatingLayers(editorRoot);
		const modeMenus = Array.from(activeDocument.querySelectorAll<HTMLElement>('div[style*="position: fixed"]'))
			.map((menu) => ({
				menu,
				buttons: Array.from(menu.querySelectorAll<HTMLButtonElement>(':scope > button'))
					.filter((button) => getEditorModeFromButton(button) !== null && button.querySelector(':scope span span')),
			}))
			.filter(({ buttons }) => {
				const modes = new Set(buttons.map((button) => getEditorModeFromButton(button)));
				return buttons.length === 3 && modes.has('editing') && modes.has('suggesting') && modes.has('viewing');
			});

		modeMenus.forEach(({ menu, buttons }) => {
			menu.dataset.docxidianModeMenu = 'true';
			markLightMenuSurface(menu, 'docxidian-mode-menu');
			menu.addClass('docxidian-mode-menu-normalized');

			buttons.forEach((button) => {
				const mode = getEditorModeFromButton(button);
				if (mode) {
					button.dataset.docxidianModeMenuItem = mode;
				}

				button.addClass('docxidian-mode-menu-item');

				const icon = button.querySelector<HTMLElement>(':scope > svg:first-child');
				if (icon) {
					icon.addClass('docxidian-mode-menu-icon');
				}

				const labelColumn = button.querySelector<HTMLElement>(':scope > span');
				if (labelColumn) {
					labelColumn.addClass('docxidian-mode-menu-label');
				}

				const checkIcon = button.querySelector<HTMLElement>(':scope > svg:last-child:not(:first-child)');
				if (checkIcon) {
					checkIcon.addClass('docxidian-mode-menu-check');
				}
			});
		});
	}, []);

	useEffect(() => {
		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		normalizeEditorModeDropdown();
		const observer = new MutationObserver(normalizeEditorModeDropdown);
		observer.observe(activeDocument.body, {
			childList: true,
			subtree: true,
		});

		return () => observer.disconnect();
	}, [buffer, filePath, isLoading, normalizeEditorModeDropdown]);

	useEffect(() => {
		const handleModePointerDown = (evt: PointerEvent) => {
			const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			if (!editorRoot || !(evt.target instanceof Element)) {
				return;
			}

			const button = evt.target.closest('button');
			if (!(button instanceof HTMLButtonElement)) {
				return;
			}

			const isEditorModeButton = editorRoot.contains(button) || button.closest('[data-docxidian-mode-menu]');
			if (!isEditorModeButton) {
				return;
			}
			const mode = getEditorModeFromButton(button);
			if (mode) {
				window.setTimeout(() => setMode(mode), 0);
			}
		};

		activeDocument.addEventListener('pointerdown', handleModePointerDown, true);
		return () => activeDocument.removeEventListener('pointerdown', handleModePointerDown, true);
	}, [setMode]);

	const clearAutosaveTimeout = useCallback(() => {
		if (autosaveTimeoutRef.current !== null) {
			window.clearTimeout(autosaveTimeoutRef.current);
			autosaveTimeoutRef.current = null;
		}
	}, []);

	const clearRenameTimeout = useCallback(() => {
		if (renameTimeoutRef.current !== null) {
			window.clearTimeout(renameTimeoutRef.current);
			renameTimeoutRef.current = null;
		}
	}, []);

	const syncVerticalRulerMarkers = useCallback((docxDocument: DocxDocumentWithSectionProperties | null | undefined) => {
		if (!showRuler || !docxDocument) {
			return;
		}

		const documentProperties = docxDocument.package?.document;
		const sectionProperties = {
			...documentProperties?.sections?.[0]?.properties,
			...documentProperties?.finalSectionProperties,
		};
		const pageHeight = sectionProperties.pageHeight ?? DEFAULT_PAGE_HEIGHT_TWIPS;
		const topMargin = sectionProperties.marginTop ?? DEFAULT_MARGIN_TWIPS;
		const bottomMargin = sectionProperties.marginBottom ?? DEFAULT_MARGIN_TWIPS;
		const ruler = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current} .docx-vertical-ruler`);

		if (!ruler || pageHeight <= 0) {
			return;
		}

		const pxPerTwip = ruler.getBoundingClientRect().height / pageHeight;
		const topMarker = ruler.querySelector<HTMLElement>('.docx-ruler-marker-topMargin');
		const bottomMarker = ruler.querySelector<HTMLElement>('.docx-ruler-marker-bottomMargin');

		if (topMarker) {
			topMarker.style.top = `${Math.round(topMargin * pxPerTwip - 5)}px`;
		}
		if (bottomMarker) {
			bottomMarker.style.top = `${Math.round((pageHeight - bottomMargin) * pxPerTwip - 5)}px`;
		}
	}, [showRuler]);

	const scheduleVerticalRulerMarkerSync = useCallback((document: DocxDocumentWithSectionProperties | null | undefined) => {
		if (rulerSyncFrameRef.current !== null) {
			window.cancelAnimationFrame(rulerSyncFrameRef.current);
		}
		if (rulerSyncTimeoutRef.current !== null) {
			window.clearTimeout(rulerSyncTimeoutRef.current);
		}

		rulerSyncFrameRef.current = window.requestAnimationFrame(() => {
			rulerSyncFrameRef.current = null;
			syncVerticalRulerMarkers(document);
			window.requestAnimationFrame(() => syncVerticalRulerMarkers(document));
		});
		rulerSyncTimeoutRef.current = window.setTimeout(() => {
			rulerSyncTimeoutRef.current = null;
			syncVerticalRulerMarkers(document);
		}, 50);
	}, [syncVerticalRulerMarkers]);

	const clearInitialDocumentCenter = useCallback(() => {
		if (initialCenterFrameRef.current !== null) {
			window.cancelAnimationFrame(initialCenterFrameRef.current);
			initialCenterFrameRef.current = null;
		}

		for (const timeout of initialCenterTimeoutsRef.current) {
			window.clearTimeout(timeout);
		}
		initialCenterTimeoutsRef.current = [];
	}, []);

	const centerInitialDocumentViewport = useCallback(() => {
		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return false;
		}

		return centerEditorViewport(editorRoot);
	}, []);

	const scheduleInitialDocumentCenter = useCallback(() => {
		clearInitialDocumentCenter();

		const runCenter = () => {
			centerInitialDocumentViewport();
		};

		initialCenterFrameRef.current = window.requestAnimationFrame(() => {
			initialCenterFrameRef.current = null;
			runCenter();
			window.requestAnimationFrame(runCenter);
		});
		initialCenterTimeoutsRef.current = [80, 240, 600].map((delay) => window.setTimeout(runCenter, delay));
	}, [centerInitialDocumentViewport, clearInitialDocumentCenter]);

	useEffect(() => {
		if (showRuler) {
			scheduleVerticalRulerMarkerSync(editorRef.current?.getDocument());
		}
	}, [showRuler, file, buffer, scheduleVerticalRulerMarkerSync]);

	useEffect(() => {
		if (isLoading || error || !file || !buffer || centeredDocumentKeyRef.current === documentKey) {
			return;
		}

		centeredDocumentKeyRef.current = documentKey;
		scheduleInitialDocumentCenter();
		return clearInitialDocumentCenter;
	}, [buffer, clearInitialDocumentCenter, documentKey, error, file, isLoading, scheduleInitialDocumentCenter]);

	useEffect(() => {
		if (!shouldEnableTouchPinchZoom()) {
			return;
		}

		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		const hostRoot = editorRoot.closest<HTMLElement>('.docxidian-host') ?? editorRoot;
		editorRoot.addClass('docxidian-touch-pinch-root');
		hostRoot.addClass('docxidian-touch-pinch-root');

		const isEditorTarget = (target: EventTarget | null) => target instanceof Node && hostRoot.contains(target);

		const shouldIgnoreGestureSource = (source: PinchZoomState['source']) => {
			const activeSource = pinchZoomStateRef.current?.source;
			return activeSource !== undefined && activeSource !== source;
		};

		const zoomAroundViewportPoint = (nextZoom: number, viewportPoint: PointerPoint, source: PinchZoomState['source']) => {
			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || pinchState.source !== source || Math.abs(nextZoom - pinchState.lastZoom) < TOUCH_ZOOM_MIN_DELTA) {
				return false;
			}

			const scrollContainer = getScrollableEditorElement(editorRoot);
			const rect = scrollContainer.getBoundingClientRect();
			const localX = viewportPoint.x - rect.left;
			const localY = viewportPoint.y - rect.top;
			const documentX = (scrollContainer.scrollLeft + localX) / pinchState.lastZoom;
			const documentY = (scrollContainer.scrollTop + localY) / pinchState.lastZoom;

			editorRef.current?.setZoom(nextZoom);
			pinchState.lastZoom = nextZoom;

			if (pinchZoomScrollFrameRef.current !== null) {
				window.cancelAnimationFrame(pinchZoomScrollFrameRef.current);
			}
			pinchZoomScrollFrameRef.current = window.requestAnimationFrame(() => {
				pinchZoomScrollFrameRef.current = null;
				scrollContainer.scrollLeft = Math.max(0, documentX * nextZoom - localX);
				scrollContainer.scrollTop = Math.max(0, documentY * nextZoom - localY);
				scheduleVerticalRulerMarkerSync(editorRef.current?.getDocument());
			});
			return true;
		};

		const handleTouchStart = (evt: TouchEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}
			if (shouldIgnoreGestureSource('touch')) {
				return;
			}
			if (evt.touches.length !== 2) {
				if (pinchZoomStateRef.current?.source === 'touch') {
					pinchZoomStateRef.current = null;
				}
				return;
			}

			const first = evt.touches.item(0);
			const second = evt.touches.item(1);
			if (!first || !second) {
				return;
			}
			const startZoom = editorRef.current?.getZoom() ?? 1;
			const startDistance = getTouchDistance(first, second);
			if (startDistance <= 0) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			pinchZoomStateRef.current = {
				source: 'touch',
				startDistance,
				lastDistance: startDistance,
				startZoom,
				lastZoom: startZoom,
			};
		};

		const handleTouchMove = (evt: TouchEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}
			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || pinchState.source !== 'touch' || evt.touches.length !== 2) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();

			const first = evt.touches.item(0);
			const second = evt.touches.item(1);
			if (!first || !second) {
				return;
			}
			const distance = getTouchDistance(first, second);
			if (distance <= 0) {
				return;
			}

			const center = getTouchCenter(first, second);
			const didZoom = zoomAroundViewportPoint(scaleTouchZoom(pinchState.lastZoom, distance / pinchState.lastDistance), center, 'touch');
			if (didZoom) {
				pinchState.lastDistance = distance;
			}
		};

		const handleTouchEnd = (evt: TouchEvent) => {
			if (evt.touches.length < 2 && pinchZoomStateRef.current?.source === 'touch') {
				pinchZoomStateRef.current = null;
			}
		};

		const handleGestureEnd = () => {
			if (pinchZoomStateRef.current?.source === 'gesture') {
				pinchZoomStateRef.current = null;
			}
		};

		const handleGestureStart = (evt: WebKitGestureEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}
			if (shouldIgnoreGestureSource('gesture')) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			const startZoom = editorRef.current?.getZoom() ?? 1;
			pinchZoomStateRef.current = {
				source: 'gesture',
				startDistance: 1,
				lastDistance: 1,
				startZoom,
				lastZoom: startZoom,
			};
		};

		const handleGestureChange = (evt: WebKitGestureEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}

			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || typeof evt.scale !== 'number' || evt.scale <= 0) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();

			const scrollContainer = getScrollableEditorElement(editorRoot);
			const rect = scrollContainer.getBoundingClientRect();
			zoomAroundViewportPoint(scaleTouchZoom(pinchState.startZoom, evt.scale), {
				x: evt.clientX ?? rect.left + rect.width / 2,
				y: evt.clientY ?? rect.top + rect.height / 2,
			}, 'gesture');
		};

		const handlePointerDown = (evt: PointerEvent) => {
			if (evt.pointerType !== 'touch' || !isEditorTarget(evt.target)) {
				return;
			}
			if (shouldIgnoreGestureSource('pointer')) {
				return;
			}

			activeTouchPointersRef.current.set(evt.pointerId, { x: evt.clientX, y: evt.clientY });
			if (activeTouchPointersRef.current.size !== 2) {
				return;
			}

			const [first, second] = Array.from(activeTouchPointersRef.current.values());
			if (!first || !second) {
				return;
			}
			const startDistance = getPointDistance(first, second);
			if (startDistance <= 0) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			const startZoom = editorRef.current?.getZoom() ?? 1;
			pinchZoomStateRef.current = {
				source: 'pointer',
				startDistance,
				lastDistance: startDistance,
				startZoom,
				lastZoom: startZoom,
			};
		};

		const handlePointerMove = (evt: PointerEvent) => {
			if (evt.pointerType !== 'touch' || !activeTouchPointersRef.current.has(evt.pointerId)) {
				return;
			}

			activeTouchPointersRef.current.set(evt.pointerId, { x: evt.clientX, y: evt.clientY });
			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || pinchState.source !== 'pointer' || activeTouchPointersRef.current.size !== 2) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			const [first, second] = Array.from(activeTouchPointersRef.current.values());
			if (!first || !second) {
				return;
			}
			const distance = getPointDistance(first, second);
			if (distance <= 0) {
				return;
			}

			const didZoom = zoomAroundViewportPoint(
				scaleTouchZoom(pinchState.lastZoom, distance / pinchState.lastDistance),
				getPointCenter(first, second),
				'pointer',
			);
			if (didZoom) {
				pinchState.lastDistance = distance;
			}
		};

		const handlePointerEnd = (evt: PointerEvent) => {
			if (evt.pointerType !== 'touch') {
				return;
			}

			activeTouchPointersRef.current.delete(evt.pointerId);
			if (activeTouchPointersRef.current.size < 2 && pinchZoomStateRef.current?.source === 'pointer') {
				pinchZoomStateRef.current = null;
			}
		};

		activeDocument.addEventListener('touchstart', handleTouchStart, { passive: false, capture: true });
		activeDocument.addEventListener('touchmove', handleTouchMove, { passive: false, capture: true });
		activeDocument.addEventListener('touchend', handleTouchEnd, { passive: true, capture: true });
		activeDocument.addEventListener('touchcancel', handleTouchEnd, { passive: true, capture: true });
		activeDocument.addEventListener('gesturestart', handleGestureStart, { passive: false, capture: true });
		activeDocument.addEventListener('gesturechange', handleGestureChange, { passive: false, capture: true });
		activeDocument.addEventListener('gestureend', handleGestureEnd, { passive: true, capture: true });
		activeDocument.addEventListener('pointerdown', handlePointerDown, { passive: false, capture: true });
		activeDocument.addEventListener('pointermove', handlePointerMove, { passive: false, capture: true });
		activeDocument.addEventListener('pointerup', handlePointerEnd, { passive: true, capture: true });
		activeDocument.addEventListener('pointercancel', handlePointerEnd, { passive: true, capture: true });

		return () => {
			editorRoot.removeClass('docxidian-touch-pinch-root');
			hostRoot.removeClass('docxidian-touch-pinch-root');
			activeDocument.removeEventListener('touchstart', handleTouchStart, true);
			activeDocument.removeEventListener('touchmove', handleTouchMove, true);
			activeDocument.removeEventListener('touchend', handleTouchEnd, true);
			activeDocument.removeEventListener('touchcancel', handleTouchEnd, true);
			activeDocument.removeEventListener('gesturestart', handleGestureStart, true);
			activeDocument.removeEventListener('gesturechange', handleGestureChange, true);
			activeDocument.removeEventListener('gestureend', handleGestureEnd, true);
			activeDocument.removeEventListener('pointerdown', handlePointerDown, true);
			activeDocument.removeEventListener('pointermove', handlePointerMove, true);
			activeDocument.removeEventListener('pointerup', handlePointerEnd, true);
			activeDocument.removeEventListener('pointercancel', handlePointerEnd, true);
			pinchZoomStateRef.current = null;
			activeTouchPointersRef.current.clear();
			if (pinchZoomScrollFrameRef.current !== null) {
				window.cancelAnimationFrame(pinchZoomScrollFrameRef.current);
				pinchZoomScrollFrameRef.current = null;
			}
		};
	}, [buffer, filePath, isLoading, scheduleVerticalRulerMarkerSync]);

	useEffect(() => () => {
		clearAutosaveTimeout();
		clearInitialDocumentCenter();
		clearRenameTimeout();
		if (rulerSyncFrameRef.current !== null) {
			window.cancelAnimationFrame(rulerSyncFrameRef.current);
			rulerSyncFrameRef.current = null;
		}
		if (rulerSyncTimeoutRef.current !== null) {
			window.clearTimeout(rulerSyncTimeoutRef.current);
			rulerSyncTimeoutRef.current = null;
		}
	}, [clearAutosaveTimeout, clearInitialDocumentCenter, clearRenameTimeout]);

	const persistDocument = useCallback(async (output: ArrayBuffer, options?: SaveDocumentOptions) => {
		if (!file) {
			return false;
		}

		if (isSavingRef.current) {
			return false;
		}
		isSavingRef.current = true;
		setSaveStatus('saving');
		const saveVersion = options?.dirtyVersion ?? dirtyVersionRef.current;

		try {
			await onSave(output);
			if (dirtyVersionRef.current === saveVersion) {
				onDirtyChange(false);
				setSaveStatus('saved');
			} else {
				setSaveStatus('unsaved');
			}
			if (!options?.silent) {
				new Notice(`Saved ${file.name}`);
			}
			return true;
		} catch (saveError) {
			const message = saveError instanceof Error ? saveError.message : 'Unknown save error';
			setSaveStatus('failed');
			new Notice(`Could not save ${file.name}: ${message}`);
			return false;
		} finally {
			setTimeout(() => {
				isSavingRef.current = false;
			}, 300);
		}
	}, [file, onDirtyChange, onSave]);

	const saveDocument = useCallback(async (options?: SaveDocumentOptions) => {
		clearAutosaveTimeout();

		if (!file) {
			new Notice('No docx file is open.');
			return false;
		}

		const saveOptions = { ...options, dirtyVersion: dirtyVersionRef.current };
		setSaveStatus('saving');
		pendingSaveOptionsRef.current = saveOptions;
		pendingSavePromiseRef.current = null;
		const output = await editorRef.current?.save({ selective: false });
		const pendingSavePromise = pendingSavePromiseRef.current;
		pendingSaveOptionsRef.current = undefined;
		pendingSavePromiseRef.current = null;

		if (!output) {
			setSaveStatus('failed');
			new Notice(`Could not save ${file.name}: the editor did not return a document.`);
			return false;
		}

		if (pendingSavePromise) {
			return pendingSavePromise;
		}

		return persistDocument(output, saveOptions);
	}, [clearAutosaveTimeout, file, persistDocument]);

	const exportDocumentBuffer = useCallback(async (options?: ExportDocumentBufferOptions) => {
		if (!options?.preserveAutosave) {
			clearAutosaveTimeout();
		}

		if (!file) {
			new Notice('No docx file is open.');
			return null;
		}

		pendingSaveModeRef.current = 'export';
		pendingSaveOptionsRef.current = undefined;
		pendingSavePromiseRef.current = null;

		try {
			const output = await editorRef.current?.save({ selective: false });
			if (!output) {
				new Notice(`Could not export ${file.name}: the editor did not return a document.`);
				return null;
			}

			return output;
		} finally {
			pendingSaveModeRef.current = 'save';
			pendingSaveOptionsRef.current = undefined;
			pendingSavePromiseRef.current = null;
		}
	}, [clearAutosaveTimeout, file]);

	const exportRenderedPdfBuffer = useCallback(async () => {
		const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			warnLog('export', 'Could not export rendered PDF because the editor root is missing', {
				editorClassName: editorClassNameRef.current,
				hasRenderedDomContext: Boolean(renderedDomContextRef.current?.pagesContainer?.isConnected),
			});
			new Notice('The docx editor is not ready yet.');
			return null;
		}

		try {
			const renderedPagesContainer = renderedDomContextRef.current?.pagesContainer ?? null;
			const pdfBuffer = await exportRenderedPagesToPdf(editorRoot, renderedPagesContainer);
			if (!pdfBuffer) {
				new Notice('Could not find rendered DOCX pages to export.');
				return null;
			}
			debugLog('export', 'Exported rendered DOCX pages to PDF', {
				bytes: pdfBuffer.byteLength,
			});
			return pdfBuffer;
		} catch (renderError) {
			errorLog('export', 'Could not export rendered DOCX pages to PDF', renderError);
			const message = renderError instanceof Error ? renderError.message : 'Unknown PDF render error';
			new Notice(`Could not render formatted PDF: ${message}`);
			return null;
		}
	}, []);

	const handleRenderedDomContextReady = useCallback((context: RenderedDomContext) => {
		renderedDomContextRef.current = context;
		debugLog('export', 'Rendered DOCX DOM context ready', {
			pageCount: context.pagesContainer.querySelectorAll('.layout-page').length,
			zoom: context.zoom,
		});
	}, []);

	useEffect(() => {
		if (!autosave) {
			clearAutosaveTimeout();
		}
	}, [autosave, clearAutosaveTimeout]);

	const scheduleAutosave = useCallback(() => {
		if (!autosave) {
			clearAutosaveTimeout();
			return;
		}

		clearAutosaveTimeout();
		autosaveTimeoutRef.current = window.setTimeout(() => {
			autosaveTimeoutRef.current = null;
			void saveDocument({ silent: true });
		}, 1500);
	}, [autosave, clearAutosaveTimeout, saveDocument]);

	const scheduleRename = useCallback((name: string) => {
		clearRenameTimeout();
		renameTimeoutRef.current = window.setTimeout(async () => {
			renameTimeoutRef.current = null;
			try {
				await onDocumentNameChange(name);
			} catch (renameError) {
				const message = renameError instanceof Error ? renameError.message : 'Unknown rename error';
				new Notice(`Could not rename ${file?.name ?? 'document'}: ${message}`);
				setDocumentName(file?.name ?? '');
			}
		}, 700);
	}, [clearRenameTimeout, file, onDocumentNameChange]);

	const insertCustomTable = useCallback((rows: number, columns: number) => {
		if (editorMode === 'viewing') {
			new Notice('Switch to editing mode to insert a table.');
			return;
		}

		if (!file || !buffer) {
			new Notice('Open a loaded docx file to insert a table.');
			return;
		}

		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view) {
			new Notice('The DOCX editor is not ready yet.');
			return;
		}

		const inserted = insertTable(
			clampCustomTableSize(rows),
			clampCustomTableSize(columns),
		)(view.state, view.dispatch);

		if (!inserted) {
			new Notice('Could not insert a table here.');
			return;
		}

		view.focus();
		setCustomTableDialogOpen(false);
	}, [buffer, editorMode, file]);

	const openCustomTableDialog = useCallback(() => {
		setCustomTableDialogOpen(true);
	}, []);

	const decorateTableSizeDropdown = useCallback(() => {
		activeDocument.querySelectorAll<HTMLElement>('[role="grid"]').forEach((grid) => {
			if (isTableSizeGrid(grid)) {
				appendCustomTableOption(grid, openCustomTableDialog);
			}
		});
	}, [openCustomTableDialog]);

	useEffect(() => {
		decorateTableSizeDropdown();
		const observer = new MutationObserver(decorateTableSizeDropdown);
		observer.observe(activeDocument.body, {
			childList: true,
			subtree: true,
		});

		return () => observer.disconnect();
	}, [buffer, decorateTableSizeDropdown, filePath, isLoading]);

	const applyFontFamilyToSelection = useCallback((fontFamily: string) => {
		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view || editorMode === 'viewing') {
			return false;
		}

		const applied = setFontFamily(fontFamily)(view.state, view.dispatch);
		view.focus();
		return applied;
	}, [editorMode]);

	const stopFontSizeHold = useCallback(() => {
		const hold = fontSizeHoldRef.current;
		if (!hold) {
			return;
		}

		if (hold.startTimer !== null) {
			window.clearTimeout(hold.startTimer);
		}
		if (hold.repeatTimer !== null) {
			window.clearTimeout(hold.repeatTimer);
		}

		fontSizeHoldRef.current = null;
	}, []);

	const applyFontSizeStepToSelection = useCallback((direction: FontSizeStepDirection, control: HTMLElement | null) => {
		if (editorMode === 'viewing') {
			return false;
		}

		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view) {
			return false;
		}

		const hold = fontSizeHoldRef.current;
		const currentSize = hold && hold.control === control
			? hold.currentSize
			: readFontSizeControlPoints(control);
		const nextSize = clampFontSizePoints(currentSize + direction);
		if (nextSize === currentSize) {
			return false;
		}

		const applied = setFontSize(fontSizePointsToHalfPoints(nextSize))(view.state, view.dispatch);
		if (!applied) {
			return false;
		}

		if (hold && hold.control === control) {
			hold.currentSize = nextSize;
		}
		updateFontSizeControlDisplay(control, nextSize);
		view.focus();
		return true;
	}, [editorMode]);

	const startFontSizeHold = useCallback((button: HTMLButtonElement, direction: FontSizeStepDirection) => {
		if (button.disabled || editorMode === 'viewing') {
			return;
		}

		stopFontSizeHold();

		const control = getFontSizeControl(button);
		const currentSize = readFontSizeControlPoints(control);
		const nextSize = clampFontSizePoints(currentSize + direction);
		if (nextSize === currentSize) {
			return;
		}

		const hold: FontSizeHoldState = {
			control,
			currentSize,
			direction,
			repeatCount: 0,
			repeatTimer: null,
			startTimer: null,
		};
		fontSizeHoldRef.current = hold;

		const scheduleRepeat = (delay: number) => {
			hold.repeatTimer = window.setTimeout(() => {
				if (fontSizeHoldRef.current !== hold) {
					return;
				}

				hold.repeatCount += 1;
				if (!applyFontSizeStepToSelection(hold.direction, hold.control)) {
					stopFontSizeHold();
					return;
				}

				const nextDelay = Math.max(
					FONT_SIZE_HOLD_MIN_INTERVAL_MS,
					FONT_SIZE_HOLD_INITIAL_INTERVAL_MS * Math.pow(FONT_SIZE_HOLD_INTERVAL_DECAY, hold.repeatCount),
				);
				scheduleRepeat(nextDelay);
			}, delay);
		};

		if (!applyFontSizeStepToSelection(direction, control)) {
			stopFontSizeHold();
			return;
		}

		hold.startTimer = window.setTimeout(() => {
			if (fontSizeHoldRef.current === hold) {
				scheduleRepeat(FONT_SIZE_HOLD_INITIAL_INTERVAL_MS);
			}
		}, FONT_SIZE_HOLD_INITIAL_DELAY_MS);
	}, [applyFontSizeStepToSelection, editorMode, stopFontSizeHold]);

	const openFontPicker = useCallback(() => {
		if (!fontInputRef.current) {
			new Notice('The font picker is not ready yet.');
			return;
		}

		fontInputRef.current.click();
	}, []);

	const decorateFontDropdown = useCallback(() => {
		activeDocument.querySelectorAll<HTMLElement>('[role="listbox"]').forEach((listbox) => {
			if (isFontDropdownListbox(listbox)) {
				appendImportFontOption(listbox, openFontPicker);
			}
		});
	}, [openFontPicker]);

	useEffect(() => {
		decorateFontDropdown();
		const observer = new MutationObserver(decorateFontDropdown);
		observer.observe(activeDocument.body, {
			childList: true,
			subtree: true,
		});

		return () => observer.disconnect();
	}, [buffer, decorateFontDropdown, filePath, isLoading]);

	useEffect(() => {
		let lastPointerHandledAt = 0;

		const suppressEvent = (evt: Event) => {
			evt.preventDefault();
			evt.stopImmediatePropagation();
			evt.stopPropagation();
		};

		const getScopedTarget = (target: EventTarget | null) => {
			const stepTarget = getFontSizeStepTarget(target);
			if (!stepTarget) {
				return null;
			}

			const editorRoot = activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			if (!editorRoot?.contains(stepTarget.button)) {
				return null;
			}

			return stepTarget;
		};

		const handlePressStart = (evt: MouseEvent | PointerEvent) => {
			if ('button' in evt && evt.button !== 0) {
				return;
			}

			const stepTarget = getScopedTarget(evt.target);
			if (!stepTarget) {
				return;
			}

			suppressEvent(evt);
			if (evt instanceof PointerEvent) {
				lastPointerHandledAt = performance.now();
			}
			startFontSizeHold(stepTarget.button, stepTarget.direction);
		};

		const handleMouseDown = (evt: MouseEvent) => {
			const stepTarget = getScopedTarget(evt.target);
			if (stepTarget && performance.now() - lastPointerHandledAt < 100) {
				suppressEvent(evt);
				return;
			}

			handlePressStart(evt);
		};

		const handleClick = (evt: MouseEvent) => {
			if (getScopedTarget(evt.target)) {
				suppressEvent(evt);
			}
		};

		const handleKeyDown = (evt: KeyboardEvent) => {
			if (evt.key !== 'Enter' && evt.key !== ' ') {
				return;
			}

			const stepTarget = getScopedTarget(evt.target);
			if (!stepTarget) {
				return;
			}

			suppressEvent(evt);
			applyFontSizeStepToSelection(stepTarget.direction, getFontSizeControl(stepTarget.button));
		};

		activeDocument.addEventListener('pointerdown', handlePressStart, true);
		activeDocument.addEventListener('mousedown', handleMouseDown, true);
		activeDocument.addEventListener('click', handleClick, true);
		activeDocument.addEventListener('keydown', handleKeyDown, true);
		activeDocument.addEventListener('pointerup', stopFontSizeHold, true);
		activeDocument.addEventListener('pointercancel', stopFontSizeHold, true);
		activeDocument.addEventListener('mouseup', stopFontSizeHold, true);
		window.addEventListener('blur', stopFontSizeHold);

		return () => {
			activeDocument.removeEventListener('pointerdown', handlePressStart, true);
			activeDocument.removeEventListener('mousedown', handleMouseDown, true);
			activeDocument.removeEventListener('click', handleClick, true);
			activeDocument.removeEventListener('keydown', handleKeyDown, true);
			activeDocument.removeEventListener('pointerup', stopFontSizeHold, true);
			activeDocument.removeEventListener('pointercancel', stopFontSizeHold, true);
			activeDocument.removeEventListener('mouseup', stopFontSizeHold, true);
			window.removeEventListener('blur', stopFontSizeHold);
			stopFontSizeHold();
		};
	}, [applyFontSizeStepToSelection, startFontSizeHold, stopFontSizeHold]);

	useEffect(() => {
		let activeTarget: HTMLElement | null = null;
		let tooltipEl: HTMLDivElement | null = null;
		let tooltipTimer: number | null = null;

		const clearTooltipTimer = () => {
			if (tooltipTimer !== null) {
				window.clearTimeout(tooltipTimer);
				tooltipTimer = null;
			}
		};

		const removeTooltip = () => {
			tooltipEl?.remove();
			tooltipEl = null;
		};

		const hideTooltip = () => {
			clearTooltipTimer();
			removeTooltip();
			restoreNativeTitle(activeTarget);
			activeTarget = null;
		};

		const positionTooltip = (target: HTMLElement, tooltip: HTMLDivElement) => {
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
		};

		const showTooltip = (target: HTMLElement) => {
			const label = getToolbarTooltipText(target);
			if (!label || target.isConnected === false) {
				return;
			}

			removeTooltip();
			const tooltip = activeDocument.createElement('div');
			tooltip.className = 'docxidian-toolbar-tooltip';
			tooltip.textContent = label;
			activeDocument.body.appendChild(tooltip);
			tooltipEl = tooltip;
			positionTooltip(target, tooltip);
		};

		const scheduleTooltip = (target: HTMLElement) => {
			if (target === activeTarget) {
				return;
			}

			hideTooltip();
			activeTarget = target;
			suspendNativeTitle(target);
			tooltipTimer = window.setTimeout(() => {
				tooltipTimer = null;
				if (activeTarget === target) {
					showTooltip(target);
				}
			}, TOOLBAR_TOOLTIP_DELAY_MS);
		};

		const getEditorRoot = () => activeDocument.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);

		const handlePointerOver = (evt: PointerEvent) => {
			const target = getToolbarTooltipTarget(evt.target, getEditorRoot());
			if (target) {
				scheduleTooltip(target);
			}
		};

		const handlePointerOut = (evt: PointerEvent) => {
			if (!activeTarget || evt.relatedTarget instanceof Node && activeTarget.contains(evt.relatedTarget)) {
				return;
			}

			hideTooltip();
		};

		const handleScrollOrResize = () => {
			hideTooltip();
		};

		activeDocument.addEventListener('pointerover', handlePointerOver, true);
		activeDocument.addEventListener('pointerout', handlePointerOut, true);
		activeDocument.addEventListener('scroll', handleScrollOrResize, true);
		window.addEventListener('resize', handleScrollOrResize);
		window.addEventListener('blur', hideTooltip);

		return () => {
			activeDocument.removeEventListener('pointerover', handlePointerOver, true);
			activeDocument.removeEventListener('pointerout', handlePointerOut, true);
			activeDocument.removeEventListener('scroll', handleScrollOrResize, true);
			window.removeEventListener('resize', handleScrollOrResize);
			window.removeEventListener('blur', hideTooltip);
			hideTooltip();
		};
	}, []);

	const importFontFile = useCallback(async (fontFile: File) => {
		if (!isSupportedFontFile(fontFile)) {
			new Notice('Choose a TTF, OTF, WOFF, or WOFF2 font file.');
			return;
		}

		const fontName = getUniqueImportedFontName(fontFile, fontFamilies);
		const fontFamily = cssFontFamilyName(fontName);

		try {
			const fontBuffer = await fontFile.arrayBuffer();
			const loaded = await loadFontFromBuffer(fontName, fontBuffer);
			if (!loaded) {
				new Notice(`Could not import ${fontFile.name}.`);
				return;
			}

			setImportedFonts((fonts) => [
				...fonts,
				{ name: fontName, fontFamily, category: 'other' },
			]);

			const wasApplied = applyFontFamilyToSelection(fontFamily);
			new Notice(wasApplied ? `Imported and applied ${fontName}.` : `Imported ${fontName}.`);
		} catch (fontError) {
			const message = fontError instanceof Error ? fontError.message : 'Unknown font import error';
			errorLog('editor', `Could not import font ${fontFile.name}`, fontError);
			new Notice(`Could not import font: ${message}`);
		}
	}, [applyFontFamilyToSelection, fontFamilies]);

	const handleFontInputChange = useCallback((evt: ChangeEvent<HTMLInputElement>) => {
		const fontFile = evt.currentTarget.files?.[0];
		evt.currentTarget.value = '';
		if (!fontFile) {
			return;
		}

		void importFontFile(fontFile);
	}, [importFontFile]);

	const insertImageFile = useCallback(async (imageFile: File) => {
		if (editorMode === 'viewing') {
			new Notice('Switch to editing mode to insert an image.');
			return;
		}

		if (!file || !buffer) {
			new Notice('Open a loaded docx file to insert an image.');
			return;
		}

		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view) {
			new Notice('The DOCX editor is not ready yet.');
			return;
		}

		const imageNodeType = view.state.schema.nodes.image;
		if (!imageNodeType) {
			new Notice('This DOCX editor cannot insert images here.');
			return;
		}

		try {
			const src = await readFileAsDataUrl(imageFile);
			const { width, height } = await loadImageDimensions(src);
			const imageNode = imageNodeType.create({
				src,
				alt: imageFile.name,
				width,
				height,
				rId: `rId_img_${Date.now()}`,
				wrapType: 'inline',
				displayMode: 'inline',
			});
			const { from } = view.state.selection;
			view.dispatch(view.state.tr.insert(from, imageNode).scrollIntoView());
			view.focus();
		} catch (insertError) {
			const message = insertError instanceof Error ? insertError.message : 'Unknown image insert error';
			errorLog('editor', `Could not insert image into ${file.name}`, insertError);
			new Notice(`Could not insert image: ${message}`);
		}
	}, [buffer, editorMode, file]);

	const openImagePicker = useCallback(() => {
		if (editorMode === 'viewing') {
			new Notice('Switch to editing mode to insert an image.');
			return;
		}

		if (!file || !buffer) {
			new Notice('Open a loaded docx file to insert an image.');
			return;
		}

		if (!imageInputRef.current) {
			new Notice('The image picker is not ready yet.');
			return;
		}

		imageInputRef.current.click();
	}, [buffer, editorMode, file]);

	const handleImageInputChange = useCallback((evt: ChangeEvent<HTMLInputElement>) => {
		const imageFile = evt.currentTarget.files?.[0];
		evt.currentTarget.value = '';
		if (!imageFile) {
			return;
		}

		void insertImageFile(imageFile);
	}, [insertImageFile]);

	useImperativeHandle(ref, () => ({
		save: () => saveDocument(),
		exportBuffer: (options?: ExportDocumentBufferOptions) => exportDocumentBuffer(options),
		exportRenderedPdf: () => exportRenderedPdfBuffer(),
		pasteFromClipboard: async (options: PasteClipboardOptions) => {
			const view = editorRef.current?.getEditorRef()?.getView();
			return view ? await pasteClipboardIntoEditor(view, options) : false;
		},
		rewriteClipboardTextWithListMarkers: async () => {
			const view = editorRef.current?.getEditorRef()?.getView();
			return view ? await rewritePlainTextClipboardWithListMarkers(view) : false;
		},
		openFind: () => openFindReplaceDialog('find'),
		openFindReplace: () => openFindReplaceDialog('replace'),
		openImagePicker,
		openCustomTableDialog,
		openFontPicker,
		setMode,
		setZoom: (zoom: number) => editorRef.current?.setZoom(zoom),
	}), [exportDocumentBuffer, exportRenderedPdfBuffer, openCustomTableDialog, openFindReplaceDialog, openFontPicker, openImagePicker, saveDocument, setMode]);

	if (isLoading) {
		return null;
	}

	if (error) {
		return <div>{error}</div>;
	}

	if (!file || !buffer) {
		return null;
	}

	return (
		<>
			<input
				ref={fontInputRef}
				type="file"
				accept={FONT_FILE_ACCEPT}
				style={{ display: 'none' }}
				onChange={handleFontInputChange}
			/>
			<input
				ref={imageInputRef}
				type="file"
				accept="image/*"
				style={{ display: 'none' }}
				onChange={handleImageInputChange}
			/>
			<DocxEditor
				key={documentKey}
				ref={editorRef}
				documentBuffer={buffer}
				mode={editorMode}
				onModeChange={setMode}
				author={authorName}
				i18n={i18n}
				initialZoom={defaultZoom}
				className={editorClassNameRef.current}
				showRuler={showRuler}
				disableFindReplaceShortcuts
				externalPlugins={externalPlugins}
				fontFamilies={fontFamilies}
				documentName={documentName}
				documentNameEditable
				pluginSidebarItems={pluginSidebarItems.length > 0 ? pluginSidebarItems : undefined}
				onRenderedDomContextReady={handleRenderedDomContextReady}
				onEditorViewReady={scheduleListMarkerSelectionHighlightSync}
				onSelectionChange={scheduleListMarkerSelectionHighlightSync}
				onDocumentNameChange={(name) => {
					setDocumentName(name);
					scheduleRename(name);
				}}
				renderLogo={() => (
					<SaveButton onClick={() => void saveDocument()} />
				)}
				renderTitleBarRight={() => (
					<SaveStatusIndicator status={saveStatus} />
				)}
				onChange={() => {
					if (dirtyTrackingEnabledRef.current) {
						dirtyVersionRef.current += 1;
						setSaveStatus('unsaved');
						onDirtyChange(true);
						scheduleAutosave();
					}
					scheduleVerticalRulerMarkerSync(editorRef.current?.getDocument());
					scheduleListMarkerSelectionHighlightSync();
				}}
				onSave={(output) => {
					if (pendingSaveModeRef.current === 'export') {
						return;
					}

					const savePromise = persistDocument(output, pendingSaveOptionsRef.current);
					pendingSavePromiseRef.current = savePromise;
					void savePromise;
				}}
				onError={(docxError) => {
					errorLog('render', `Could not render ${file.name}`, docxError);
					new Notice(`Could not render ${file.name}: ${docxError.message}`);
				}}
			/>
			<CustomTableDialog
				isOpen={customTableDialogOpen}
				onClose={() => setCustomTableDialogOpen(false)}
				onInsert={insertCustomTable}
			/>
			<FindReplaceDialog
				isOpen={findDialogMode !== null}
				labels={findReplaceLabels}
				mode={findDialogMode ?? 'find'}
				searchText={findSearchText}
				replaceText={findReplaceText}
				matchCase={findMatchCase}
				wholeWord={findWholeWord}
				matchCount={findMatches.length}
				currentIndex={currentFindIndex}
				onSearchTextChange={(value) => {
					setFindSearchText(value);
					refreshFindMatches(value, findMatchCase, findWholeWord, 0);
				}}
				onReplaceTextChange={setFindReplaceText}
				onMatchCaseChange={(value) => {
					setFindMatchCase(value);
					refreshFindMatches(findSearchText, value, findWholeWord, currentFindIndex);
				}}
				onWholeWordChange={(value) => {
					setFindWholeWord(value);
					refreshFindMatches(findSearchText, findMatchCase, value, currentFindIndex);
				}}
				onModeChange={setFindDialogMode}
				onNext={() => moveFindMatch(1)}
				onPrevious={() => moveFindMatch(-1)}
				onReplace={replaceCurrentMatch}
				onReplaceAll={replaceAllMatches}
				onClose={() => {
					setFindDialogMode(null);
					setFindMatches([]);
					publishFindHighlights([], 0);
				}}
			/>
		</>
	);
});
