import { App, FileView, Modal, Notice, Platform, TFile, WorkspaceLeaf, normalizePath } from 'obsidian';
import type { Translations } from '@eigenpal/docx-editor-i18n';
import { createEditorTranslator } from './editorTranslations';
import { loadDocxEditorChunk } from './docxEditorLoader';
import { findHiddenDocxText, type HiddenTextFinding } from './docxHiddenTextScanner';
import { extractDocxText } from './docxTextExtractor';
import { isHTMLElement } from './domGuards';
import { debugLog, errorLog, infoLog, warnLog } from './logger';
import { DOCXIDIAN_LANGUAGE_OPTIONS, DEFAULT_LANGUAGE, normalizeDocxidianLanguage, type DocxidianLanguage } from './locales';
import { DEFAULT_SETTINGS, normalizeDefaultZoom } from './settings';
import type { DocxReactMount } from './DocxReactMount';
import type { DocxReactViewHandle, DocxReactViewProps } from './DocxReactView';

export const VIEW_TYPE_DOCX = 'docxidian-docx-view';

type UnsavedDocxChoice = 'save' | 'discard';
type DocxPathChoice = string | null;
type DocxExportFormatId = 'pdf' | 'docx' | 'html' | 'txt' | 'md' | 'rtf';
type DocxExportChoice = { name: string; format: DocxExportFormatId } | null;
type DocxConflictChoice = 'overwrite' | 'cancel';
type ExistingFileChoice = 'replace' | 'copy';
type EditorOptionSearchActionId =
	| 'save'
	| 'save-as'
	| 'duplicate'
	| 'paste'
	| 'paste-without-formatting'
	| 'export-pdf'
	| 'export-docx'
	| 'export-html'
	| 'export-txt'
	| 'export-md'
	| 'export-rtf'
	| 'find'
	| 'find-replace'
	| 'insert-image'
	| 'custom-table'
	| 'import-font'
	| 'find-hidden-text'
	| 'page-setup'
	| 'page-break'
	| 'table-of-contents'
	| 'left-to-right'
	| 'right-to-left'
	| 'mode-editing'
	| 'mode-suggesting'
	| 'mode-viewing'
	| 'zoom-75'
	| 'zoom-100'
	| 'zoom-125';

interface DocxFileSignature {
	path: string;
	mtime: number;
	size: number;
}

export interface DocxEditorSettingsSnapshot {
	authorName: string;
	editorLanguage: DocxidianLanguage;
	showRuler: boolean;
	autosave: boolean;
	createBackupsBeforeSave: boolean;
	defaultZoom: number;
	enableDocxSearchIndex: boolean;
	autoIndexDocxSearch: boolean;
	debugLogging: boolean;
	disableDocxFiles: boolean;
}

export interface DocxEditorSettingsController {
	getSettings: () => DocxEditorSettingsSnapshot;
	setAuthorName: (value: string) => Promise<void>;
	setEditorLanguage: (value: string) => Promise<void>;
	setShowRuler: (value: boolean) => Promise<void>;
	setAutosave: (value: boolean) => Promise<void>;
	setCreateBackupsBeforeSave: (value: boolean) => Promise<void>;
	setDefaultZoom: (value: number) => Promise<void>;
	setEnableDocxSearchIndex: (value: boolean) => Promise<void>;
	setAutoIndexDocxSearch: (value: boolean) => Promise<void>;
	setDebugLogging: (value: boolean) => Promise<void>;
	setDisableDocxFiles: (value: boolean) => Promise<void>;
	rebuildDocxSearchIndex: () => Promise<void>;
	copyDocxLog: (filePath?: string) => Promise<void>;
}

interface DocxExportFormat {
	id: DocxExportFormatId;
	label: string;
	extension: string;
}

interface EditorOptionSearchBaseItem {
	id: string;
	label: string;
	keywords: readonly string[];
}

interface EditorOptionSearchActionItem extends EditorOptionSearchBaseItem {
	kind: 'action';
	actionId: EditorOptionSearchActionId;
}

interface EditorOptionSearchControlQueryItem extends EditorOptionSearchBaseItem {
	kind: 'control-query';
	labels: readonly string[];
}

interface EditorOptionSearchControlItem extends EditorOptionSearchBaseItem {
	kind: 'control';
	element: HTMLElement;
}

type EditorOptionSearchItem =
	| EditorOptionSearchActionItem
	| EditorOptionSearchControlQueryItem
	| EditorOptionSearchControlItem;

const DOCX_EXPORT_FORMATS: readonly DocxExportFormat[] = [
	{ id: 'pdf', label: 'PDF document (.pdf)', extension: 'pdf' },
	{ id: 'docx', label: 'Word document (.docx)', extension: 'docx' },
	{ id: 'html', label: 'Web page (.html)', extension: 'html' },
	{ id: 'txt', label: 'Plain text (.txt)', extension: 'txt' },
	{ id: 'md', label: 'Markdown (.md)', extension: 'md' },
	{ id: 'rtf', label: 'Rich Text Format (.rtf)', extension: 'rtf' },
];
const KNOWN_EXPORT_EXTENSION_PATTERN = /\.(?:docx|pdf|html?|txt|md|markdown|rtf)$/i;
const DEFAULT_EXPORT_FORMAT: DocxExportFormatId = 'pdf';
const EDITOR_OPTION_SEARCH_BASE_ITEMS: readonly (EditorOptionSearchActionItem | EditorOptionSearchControlQueryItem)[] = [
	{ kind: 'action', id: 'action:save', actionId: 'save', label: 'Save', keywords: ['write', 'autosave'] },
	{ kind: 'action', id: 'action:save-as', actionId: 'save-as', label: 'Save as...', keywords: ['copy', 'new docx'] },
	{ kind: 'action', id: 'action:duplicate', actionId: 'duplicate', label: 'Duplicate current DOCX', keywords: ['copy', 'clone'] },
	{ kind: 'action', id: 'action:paste', actionId: 'paste', label: 'Paste', keywords: ['edit', 'clipboard', 'command v', 'cmd v', 'ctrl v'] },
	{ kind: 'action', id: 'action:paste-without-formatting', actionId: 'paste-without-formatting', label: 'Paste without formatting', keywords: ['edit', 'clipboard', 'plain text', 'paste as text'] },
	{ kind: 'action', id: 'action:export-pdf', actionId: 'export-pdf', label: 'Export as PDF', keywords: ['pdf', 'export as'] },
	{ kind: 'action', id: 'action:export-docx', actionId: 'export-docx', label: 'Export as DOCX', keywords: ['word', 'export as'] },
	{ kind: 'action', id: 'action:export-html', actionId: 'export-html', label: 'Export as HTML', keywords: ['web page', 'export as'] },
	{ kind: 'action', id: 'action:export-txt', actionId: 'export-txt', label: 'Export as plain text', keywords: ['txt', 'text', 'export as'] },
	{ kind: 'action', id: 'action:export-md', actionId: 'export-md', label: 'Export as Markdown', keywords: ['md', 'export as'] },
	{ kind: 'action', id: 'action:export-rtf', actionId: 'export-rtf', label: 'Export as RTF', keywords: ['rich text', 'export as'] },
	{ kind: 'action', id: 'action:find', actionId: 'find', label: 'Find in document', keywords: ['search', 'text'] },
	{ kind: 'action', id: 'action:find-replace', actionId: 'find-replace', label: 'Find and replace', keywords: ['replace', 'search'] },
	{ kind: 'action', id: 'action:insert-image', actionId: 'insert-image', label: 'Insert image', keywords: ['picture', 'photo', 'media'] },
	{ kind: 'action', id: 'action:custom-table', actionId: 'custom-table', label: 'Custom table', keywords: ['insert table', 'rows', 'columns'] },
	{ kind: 'action', id: 'action:import-font', actionId: 'import-font', label: 'Import font', keywords: ['typeface', 'ttf', 'otf', 'woff'] },
	{ kind: 'action', id: 'action:find-hidden-text', actionId: 'find-hidden-text', label: 'Find hidden text', keywords: ['security', 'prompt injection', 'invisible'] },
	{ kind: 'action', id: 'action:page-setup', actionId: 'page-setup', label: 'Page setup', keywords: ['paper', 'margins', 'orientation'] },
	{ kind: 'action', id: 'action:page-break', actionId: 'page-break', label: 'Page break', keywords: ['insert page'] },
	{ kind: 'action', id: 'action:table-of-contents', actionId: 'table-of-contents', label: 'Table of contents', keywords: ['toc', 'outline'] },
	{ kind: 'action', id: 'action:left-to-right', actionId: 'left-to-right', label: 'Left to right', keywords: ['ltr', 'direction'] },
	{ kind: 'action', id: 'action:right-to-left', actionId: 'right-to-left', label: 'Right to left', keywords: ['rtl', 'direction'] },
	{ kind: 'action', id: 'action:mode-editing', actionId: 'mode-editing', label: 'Editing mode', keywords: ['edit'] },
	{ kind: 'action', id: 'action:mode-suggesting', actionId: 'mode-suggesting', label: 'Suggesting mode', keywords: ['suggest', 'review'] },
	{ kind: 'action', id: 'action:mode-viewing', actionId: 'mode-viewing', label: 'Viewing mode', keywords: ['view', 'read only'] },
	{ kind: 'action', id: 'action:zoom-75', actionId: 'zoom-75', label: 'Zoom 75%', keywords: ['small'] },
	{ kind: 'action', id: 'action:zoom-100', actionId: 'zoom-100', label: 'Zoom 100%', keywords: ['actual size', 'normal'] },
	{ kind: 'action', id: 'action:zoom-125', actionId: 'zoom-125', label: 'Zoom 125%', keywords: ['large'] },
	{ kind: 'control-query', id: 'control:bold', label: 'Bold', labels: ['Bold', 'Bold (Ctrl+B)', 'format bold', 'format_bold'], keywords: ['ctrl b', 'text formatting'] },
	{ kind: 'control-query', id: 'control:italic', label: 'Italic', labels: ['Italic', 'Italic (Ctrl+I)', 'format italic', 'format_italic'], keywords: ['ctrl i', 'text formatting'] },
	{ kind: 'control-query', id: 'control:underline', label: 'Underline', labels: ['Underline', 'Underline (Ctrl+U)', 'format underlined', 'format_underlined'], keywords: ['ctrl u', 'text formatting'] },
	{ kind: 'control-query', id: 'control:strikethrough', label: 'Strikethrough', labels: ['Strikethrough', 'strikethrough s', 'strikethrough_s'], keywords: ['strike', 'text formatting'] },
	{ kind: 'control-query', id: 'control:superscript', label: 'Superscript', labels: ['Superscript'], keywords: ['script', 'raise text'] },
	{ kind: 'control-query', id: 'control:subscript', label: 'Subscript', labels: ['Subscript'], keywords: ['script', 'lower text'] },
	{ kind: 'control-query', id: 'control:clear-formatting', label: 'Clear formatting', labels: ['Clear formatting', 'format clear', 'format_clear'], keywords: ['remove formatting'] },
	{ kind: 'control-query', id: 'control:insert-link', label: 'Insert link', labels: ['Insert link', 'Insert link (Ctrl+K)', 'link'], keywords: ['hyperlink', 'url', 'ctrl k'] },
	{ kind: 'control-query', id: 'control:undo', label: 'Undo', labels: ['Undo', 'Undo (Ctrl+Z)'], keywords: ['history', 'ctrl z'] },
	{ kind: 'control-query', id: 'control:redo', label: 'Redo', labels: ['Redo', 'Redo (Ctrl+Y)'], keywords: ['history', 'ctrl y'] },
	{ kind: 'control-query', id: 'control:bullet-list', label: 'Bullet list', labels: ['Bullet List', 'format list bulleted', 'format_list_bulleted'], keywords: ['bullets', 'list'] },
	{ kind: 'control-query', id: 'control:numbered-list', label: 'Numbered list', labels: ['Numbered List', 'format list numbered', 'format_list_numbered'], keywords: ['numbers', 'ordered list'] },
	{ kind: 'control-query', id: 'control:decrease-indent', label: 'Decrease indent', labels: ['Decrease Indent', 'format indent decrease', 'format_indent_decrease'], keywords: ['outdent', 'list'] },
	{ kind: 'control-query', id: 'control:increase-indent', label: 'Increase indent', labels: ['Increase Indent', 'format indent increase', 'format_indent_increase'], keywords: ['indent', 'list'] },
	{ kind: 'control-query', id: 'control:align-left', label: 'Align left', labels: ['Align Left', 'format align left', 'format_align_left'], keywords: ['alignment'] },
	{ kind: 'control-query', id: 'control:align-center', label: 'Center align', labels: ['Center', 'Align Center', 'format align center', 'format_align_center'], keywords: ['alignment'] },
	{ kind: 'control-query', id: 'control:align-right', label: 'Align right', labels: ['Align Right', 'format align right', 'format_align_right'], keywords: ['alignment'] },
	{ kind: 'control-query', id: 'control:justify', label: 'Justify', labels: ['Justify', 'format align justify', 'format_align_justify'], keywords: ['alignment'] },
	{ kind: 'control-query', id: 'control:font-family', label: 'Font family', labels: ['Font family', 'Font', 'Select font family'], keywords: ['typeface'] },
	{ kind: 'control-query', id: 'control:font-size', label: 'Font size', labels: ['Font size', 'Select font size'], keywords: ['text size', 'point size'] },
	{ kind: 'control-query', id: 'control:font-size-decrease', label: 'Decrease font size', labels: ['Decrease font size', 'font size decrease'], keywords: ['smaller text'] },
	{ kind: 'control-query', id: 'control:font-size-increase', label: 'Increase font size', labels: ['Increase font size', 'font size increase'], keywords: ['larger text'] },
	{ kind: 'control-query', id: 'control:font-color', label: 'Font color', labels: ['Font Color', 'Text color'], keywords: ['text color'] },
	{ kind: 'control-query', id: 'control:highlight-color', label: 'Text highlight color', labels: ['Text Highlight Color', 'Highlight Color'], keywords: ['highlight', 'background color'] },
	{ kind: 'control-query', id: 'control:line-spacing', label: 'Line spacing', labels: ['Line spacing', 'Line spacing: Single', 'Line spacing: Double'], keywords: ['paragraph spacing'] },
];

function getExportFormat(formatId: DocxExportFormatId): DocxExportFormat {
	return DOCX_EXPORT_FORMATS.find(format => format.id === formatId) ?? DOCX_EXPORT_FORMATS[0]!;
}

function getExportBaseName(name: string): string {
	return name
		.trim()
		.replace(/[\\/]/g, '-')
		.replace(KNOWN_EXPORT_EXTENSION_PATTERN, '')
		.trim();
}

function withExportExtension(name: string, formatId: DocxExportFormatId): string {
	const baseName = getExportBaseName(name);
	if (!baseName) {
		return '';
	}

	return `${baseName}.${getExportFormat(formatId).extension}`;
}

function escapeHtml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}

function createPlainTextHtml(text: string, title: string): string {
	const paragraphs = text
		.split(/\n{2,}/)
		.map(paragraph => paragraph.trim())
		.filter(Boolean);
	const body = paragraphs.length > 0
		? paragraphs.map(paragraph => `<p>${escapeHtml(paragraph).replace(/\n/g, '<br>')}</p>`).join('\n')
		: '<p></p>';

	return [
		'<!doctype html>',
		'<html>',
		'<head>',
		'<meta charset="utf-8">',
		`<title>${escapeHtml(title)}</title>`,
		'<style>body{font-family:Arial,Helvetica,sans-serif;line-height:1.5;margin:48px;max-width:760px;}p{margin:0 0 1em;}</style>',
		'</head>',
		'<body>',
		body,
		'</body>',
		'</html>',
		'',
	].join('\n');
}

function escapeRtf(value: string): string {
	return Array.from(value).map((char) => {
		if (char === '\\') {
			return '\\\\';
		}
		if (char === '{') {
			return '\\{';
		}
		if (char === '}') {
			return '\\}';
		}
		if (char === '\n') {
			return '\\par\n';
		}

		const codePoint = char.codePointAt(0) ?? 0;
		if (codePoint > 127) {
			const signedCodePoint = codePoint > 32767 ? codePoint - 65536 : codePoint;
			return `\\u${signedCodePoint}?`;
		}

		return char;
	}).join('');
}

function createPlainTextRtf(text: string): string {
	return `{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Helvetica;}}\\f0\\fs24\n${escapeRtf(text)}\n}\n`;
}

function getBinaryExportContent(content: ArrayBuffer | ArrayBufferView | string): ArrayBuffer | null {
	if (typeof content === 'string') {
		return null;
	}

	if (content instanceof ArrayBuffer || Object.prototype.toString.call(content) === '[object ArrayBuffer]') {
		return content as ArrayBuffer;
	}

	if (ArrayBuffer.isView(content)) {
		const copy = new Uint8Array(content.byteLength);
		copy.set(new Uint8Array(content.buffer, content.byteOffset, content.byteLength));
		return copy.buffer;
	}

	return null;
}

class UnsavedDocxModal extends Modal {
	private resolveChoice: (choice: UnsavedDocxChoice) => void;
	private resolved = false;

	constructor(
		app: App,
		private fileName: string,
		resolveChoice: (choice: UnsavedDocxChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'Save changes?' });
		contentEl.createEl('p', { text: `${this.fileName} has unsaved changes.` });

		const buttonRow = contentEl.createDiv({ cls: 'docxidian-unsaved-actions' });
		const discardButton = buttonRow.createEl('button', { text: 'Discard' });
		const saveButton = buttonRow.createEl('button', { text: 'Save' });
		saveButton.addClass('mod-cta');

		discardButton.addEventListener('click', () => this.choose('discard'));
		saveButton.addEventListener('click', () => this.choose('save'));
	}

	onClose() {
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose('discard');
		}
	}

	private choose(choice: UnsavedDocxChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

class DocxPathModal extends Modal {
	private resolveChoice: (choice: DocxPathChoice) => void;
	private resolved = false;

	constructor(
		app: App,
		private title: string,
		private description: string,
		private initialPath: string,
		private actionLabel: string,
		resolveChoice: (choice: DocxPathChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: this.title });
		contentEl.createEl('p', { text: this.description });

		const form = contentEl.createEl('form', { cls: 'docxidian-path-form' });
		const input = form.createEl('input', {
			cls: 'docxidian-path-input',
			type: 'text',
		});
		input.value = this.initialPath;
		input.setAttribute('spellcheck', 'false');

		const buttonRow = form.createDiv({ cls: 'docxidian-unsaved-actions' });
		const cancelButton = buttonRow.createEl('button', { text: 'Cancel', type: 'button' });
		const saveButton = buttonRow.createEl('button', { text: this.actionLabel, type: 'submit' });
		saveButton.addClass('mod-cta');

		cancelButton.addEventListener('click', () => this.choose(null));
		form.addEventListener('submit', (evt) => {
			evt.preventDefault();
			this.choose(input.value);
		});

		window.setTimeout(() => {
			input.focus();
			input.select();
		});
	}

	onClose() {
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose(null);
		}
	}

	private choose(choice: DocxPathChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

class DocxExportModal extends Modal {
	private resolveChoice: (choice: DocxExportChoice) => void;
	private resolved = false;

	constructor(
		app: App,
		private initialName: string,
		private initialFormat: DocxExportFormatId,
		resolveChoice: (choice: DocxExportChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'Export as' });
		contentEl.createEl('p', { text: 'Export a copy next to the original file. If the file already exists, you can replace it or keep both.' });

		const form = contentEl.createEl('form', { cls: 'docxidian-path-form' });
		const formatLabel = form.createEl('label', { cls: 'docxidian-export-field' });
		formatLabel.createSpan({ text: 'Format' });
		const formatSelect = formatLabel.createEl('select', { cls: 'docxidian-path-select' });
		for (const format of DOCX_EXPORT_FORMATS) {
			const option = formatSelect.createEl('option', { text: format.label, value: format.id });
			option.selected = format.id === this.initialFormat;
		}

		const nameLabel = form.createEl('label', { cls: 'docxidian-export-field' });
		nameLabel.createSpan({ text: 'File name' });
		const input = nameLabel.createEl('input', {
			cls: 'docxidian-path-input',
			type: 'text',
		});
		input.value = withExportExtension(this.initialName, this.initialFormat);
		input.setAttribute('spellcheck', 'false');

		formatSelect.addEventListener('change', () => {
			input.value = withExportExtension(input.value, formatSelect.value as DocxExportFormatId);
		});

		const buttonRow = form.createDiv({ cls: 'docxidian-unsaved-actions' });
		const cancelButton = buttonRow.createEl('button', { text: 'Cancel', type: 'button' });
		const exportButton = buttonRow.createEl('button', { text: 'Export', type: 'submit' });
		exportButton.addClass('mod-cta');

		cancelButton.addEventListener('click', () => this.choose(null));
		form.addEventListener('submit', (evt) => {
			evt.preventDefault();
			this.choose({
				name: input.value,
				format: formatSelect.value as DocxExportFormatId,
			});
		});

		window.setTimeout(() => {
			input.focus();
			input.select();
		});
	}

	onClose() {
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose(null);
		}
	}

	private choose(choice: DocxExportChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

class ExistingFileModal extends Modal {
	private resolveChoice: (choice: ExistingFileChoice) => void;
	private resolved = false;

	constructor(
		app: App,
		private filePath: string,
		resolveChoice: (choice: ExistingFileChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'This file already exists.' });
		contentEl.createEl('p', { text: 'Replace it?' });
		contentEl.createEl('p', { text: this.filePath });

		const buttonRow = contentEl.createDiv({ cls: 'docxidian-unsaved-actions' });
		const noButton = buttonRow.createEl('button', { text: 'No' });
		const yesButton = buttonRow.createEl('button', { text: 'Yes' });
		yesButton.addClass('mod-warning');

		noButton.addEventListener('click', () => this.choose('copy'));
		yesButton.addEventListener('click', () => this.choose('replace'));
	}

	onClose() {
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose('copy');
		}
	}

	private choose(choice: ExistingFileChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

class HiddenTextScanModal extends Modal {
	constructor(
		app: App,
		private fileName: string,
		private findings: HiddenTextFinding[],
		private partsScanned: number,
	) {
		super(app);
	}

	onOpen() {
		const { contentEl } = this;
		this.modalEl.addClass('docxidian-hidden-text-shell');
		contentEl.empty();
		contentEl.addClass('docxidian-hidden-text-modal');
		contentEl.createEl('h2', { text: 'Find hidden text' });

		if (this.findings.length === 0) {
			contentEl.createEl('p', {
				cls: 'docxidian-hidden-text-summary docxidian-hidden-text-empty',
				text: `No hidden, white, or tiny text was found in ${this.fileName}. Scanned ${this.partsScanned} document part(s).`,
			});
			return;
		}

		contentEl.createEl('p', {
			cls: 'docxidian-hidden-text-summary',
			text: `Found ${this.findings.length} suspicious hidden text item(s) in ${this.fileName}. Review before pasting this document into an AI tool.`,
		});

		const list = contentEl.createDiv({ cls: 'docxidian-hidden-text-results' });
		for (const finding of this.findings) {
			const resultEl = list.createDiv({ cls: 'docxidian-hidden-text-result' });
			const header = resultEl.createDiv({ cls: 'docxidian-hidden-text-header' });
			header.createSpan({
				cls: 'docxidian-hidden-text-location',
				text: `${finding.partLabel}, paragraph ${finding.paragraphNumber}`,
			});
			header.createSpan({
				cls: 'docxidian-hidden-text-path',
				text: finding.partPath,
			});

			const reasons = resultEl.createDiv({ cls: 'docxidian-hidden-text-reasons' });
			for (const reason of finding.reasons) {
				reasons.createSpan({ cls: 'docxidian-hidden-text-reason', text: reason });
			}
			for (const signal of finding.promptInjectionSignals) {
				reasons.createSpan({ cls: 'docxidian-hidden-text-reason mod-warning', text: `Prompt-like text: ${signal}` });
			}

			resultEl.createEl('pre', {
				cls: 'docxidian-hidden-text-snippet',
				text: finding.text,
			});
		}
	}

	onClose() {
		this.modalEl.removeClass('docxidian-hidden-text-shell');
		this.contentEl.empty();
	}
}

class ExternalDocxChangeModal extends Modal {
	private resolveChoice: (choice: DocxConflictChoice) => void;
	private resolved = false;

	constructor(
		app: App,
		private fileName: string,
		resolveChoice: (choice: DocxConflictChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'File changed on disk' });
		contentEl.createEl('p', {
			text: `${this.fileName} was modified outside Native PowerPoint Doc Editor after it was opened.`,
		});
		contentEl.createEl('p', {
			text: 'Saving now will overwrite those outside changes. Cancel and use Save as... or Duplicate current DOCX if you want to keep both versions.',
		});

		const buttonRow = contentEl.createDiv({ cls: 'docxidian-unsaved-actions' });
		const cancelButton = buttonRow.createEl('button', { text: 'Cancel save' });
		const overwriteButton = buttonRow.createEl('button', { text: 'Overwrite anyway' });
		overwriteButton.addClass('mod-warning');

		cancelButton.addEventListener('click', () => this.choose('cancel'));
		overwriteButton.addEventListener('click', () => this.choose('overwrite'));
	}

	onClose() {
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose('cancel');
		}
	}

	private choose(choice: DocxConflictChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

function normalizeMenuText(text: string): string {
	return text.replace(/\s+/g, ' ').trim().toLowerCase();
}

function toTitleCase(value: string): string {
	return value.replace(/\b[a-z]/g, char => char.toUpperCase());
}

function cleanEditorOptionLabel(rawLabel: string): string {
	let label = rawLabel
		.replace(/[_-]+/g, ' ')
		.replace(/\s+/g, ' ')
		.trim();

	label = label
		.replace(/\s*\((?:ctrl|cmd|command|shift|alt|option|⌘)[^)]+\)\s*$/i, '')
		.replace(/\s*(?:ctrl|cmd|command|⌘)\s*\+.*$/i, '')
		.replace(/^format\s+/i, '')
		.trim();

	return /^[a-z0-9 ]+$/.test(label) ? toTitleCase(label) : label;
}

function getEditorOptionRawLabel(element: HTMLElement): string {
	return element.getAttribute('data-docxidian-search-label')
		?? element.getAttribute('aria-label')
		?? element.dataset.docxidianNativeTitle
		?? element.dataset.docxidianTooltipTitle
		?? element.getAttribute('title')
		?? element.textContent
		?? element.dataset.testid
		?? '';
}

function getEditorOptionControlLabel(element: HTMLElement): string {
	return cleanEditorOptionLabel(getEditorOptionRawLabel(element));
}

function getEditorOptionControlKeywords(element: HTMLElement, label: string): string[] {
	const keywords = new Set<string>();
	const testId = element.dataset.testid;
	const rawText = element.textContent?.trim();
	const groupLabel = element.closest<HTMLElement>('[role="group"][aria-label]')?.getAttribute('aria-label');
	const nativeTitle = element.dataset.docxidianNativeTitle ?? element.dataset.docxidianTooltipTitle ?? element.getAttribute('title');

	for (const value of [label, testId, rawText, groupLabel, nativeTitle]) {
		const normalized = cleanEditorOptionLabel(value ?? '');
		if (normalized) {
			keywords.add(normalized);
		}
	}

	return Array.from(keywords);
}

function isTopLevelEditorMenuButton(element: HTMLElement): boolean {
	const parent = element.parentElement;
	return Boolean(
		parent
		&& parent.parentElement?.getAttribute('role') === 'menubar'
		&& parent.querySelector(':scope > button') === element
	);
}

function shouldSkipEditorOptionControl(element: HTMLElement): boolean {
	return Boolean(
		element.closest('.docxidian-option-search-menu')
		|| element.closest('.docxidian-edit-menu')
		|| element.closest('[data-docxidian-edit-menu-item]')
		|| element.closest('[data-docxidian-search-menu-item]')
		|| element.closest('[data-docxidian-settings-menu-item]')
		|| element.closest('[data-docxidian-no-toolbar-tooltip]')
		|| element.closest('[data-testid="title-bar"] input')
		|| isTopLevelEditorMenuButton(element)
	);
}

function isSearchableEditorOptionLabel(label: string): boolean {
	const normalizedLabel = normalizeMenuText(label);
	return normalizedLabel.length > 1 && !['file', 'edit', 'format', 'insert', 'search', 'settings', 'help'].includes(normalizedLabel);
}

function isVisibleEditorOptionControl(element: HTMLElement): boolean {
	const style = window.getComputedStyle(element);
	return style.display !== 'none'
		&& style.visibility !== 'hidden'
		&& style.pointerEvents !== 'none'
		&& element.getClientRects().length > 0;
}

function textStartsWithMenuLabel(text: string, label: string): boolean {
	const normalizedLabel = normalizeMenuText(label);
	if (!normalizedLabel || !text.startsWith(normalizedLabel)) {
		return false;
	}

	const suffix = text.slice(normalizedLabel.length);
	return suffix === '' || /^\s|^ctrl|^cmd|^⌘/.test(suffix);
}

function isPrimaryFindShortcut(evt: KeyboardEvent): boolean {
	const key = evt.key.toLowerCase();
	const isMacFind = evt.metaKey && !evt.ctrlKey;
	const isNonMacFind = evt.ctrlKey && !evt.metaKey && !Platform.isMacOS;
	const hasPrimaryModifier = isMacFind || isNonMacFind;
	return key === 'f' && hasPrimaryModifier && !evt.altKey && !evt.shiftKey;
}

function getEditorMenuLabels(locale: Translations | undefined) {
	const translate = createEditorTranslator(locale);

	return {
		file: normalizeMenuText(translate('toolbar.file', undefined, 'File')),
		edit: 'edit',
		format: normalizeMenuText(translate('toolbar.format', undefined, 'Format')),
		insert: normalizeMenuText(translate('toolbar.insert', undefined, 'Insert')),
		help: normalizeMenuText(translate('toolbar.help', undefined, 'Help')),
		save: [
			translate('toolbar.save', undefined, 'Save'),
			translate('common.save', undefined, 'Save'),
		],
		pageSetup: [translate('toolbar.pageSetup', undefined, 'Page Setup'), 'Page setup'],
		pageBreak: [translate('toolbar.pageBreak', undefined, 'Page Break'), 'Page break'],
		tableOfContents: [translate('toolbar.tableOfContents', undefined, 'Table of Contents'), 'Table of contents'],
		leftToRight: [translate('toolbar.leftToRight', undefined, 'Left to Right'), 'Left to right'],
		rightToLeft: [translate('toolbar.rightToLeft', undefined, 'Right to Left'), 'Right to left'],
	};
}

function shouldHandleEditorSaveClick(target: EventTarget | null, saveLabels: string[]) {
	if (!(target instanceof Element)) {
		return false;
	}

	let candidate: Element | null = target;
	while (candidate && candidate !== activeDocument.body) {
		if (isHTMLElement(candidate)) {
			const text = normalizeMenuText(candidate.textContent ?? '');
			if (saveLabels.some((label) => textStartsWithMenuLabel(text, label))) {
				return true;
			}
		}

		candidate = candidate.parentElement;
	}

	return false;
}

export class DocxView extends FileView {
	private hostEl: HTMLDivElement | null = null;
	private reactMount: DocxReactMount | null = null;
	private reactMountLoading = false;
	private buffer: ArrayBuffer | null = null;
	private error: string | null = null;
	private isLoading = false;
	private isDirty = false;
	private documentSession = 0;
	private lastKnownFileSignature: DocxFileSignature | null = null;
	private backupCreatedForOpenFile = false;
	private reserveReviewSidebar = false;
	private hostResizeObserver: ResizeObserver | null = null;
	private titleObserver: MutationObserver | null = null;
	private helpMenuObserver: MutationObserver | null = null;
	private searchMenuObserver: MutationObserver | null = null;
	private editMenuObserver: MutationObserver | null = null;
	private fileMenuObserver: MutationObserver | null = null;
	private insertMenuObserver: MutationObserver | null = null;
	private nativeMenuStyleObserver: MutationObserver | null = null;
	private settingsMenuObserver: MutationObserver | null = null;
	private optionSearchPopoverEl: HTMLElement | null = null;
	private optionSearchCleanup: (() => void) | null = null;
	private editorEditPopoverEl: HTMLElement | null = null;
	private editorEditCleanup: (() => void) | null = null;
	private editorSettingsPopoverEl: HTMLElement | null = null;
	private editorSettingsCleanup: (() => void) | null = null;

	constructor(
		leaf: WorkspaceLeaf,
		private getAuthorName: () => string,
		private getEditorLocale: () => Translations | undefined,
		private getShowRuler: () => boolean,
		private getAutosave: () => boolean,
		private getCreateBackupsBeforeSave: () => boolean,
		private getDefaultZoom: () => number,
		private settingsController: DocxEditorSettingsController,
	) {
		super(leaf);
	}

	getViewType() {
		return VIEW_TYPE_DOCX;
	}

	getDisplayText() {
		return this.file?.basename ?? 'DOCX';
	}

	getIcon() {
		return 'file-text';
	}

	canAcceptExtension(extension: string) {
		return extension.toLowerCase() === 'docx';
	}

	async onOpen() {
		debugLog('view', 'Opening DOCX view');
		this.contentEl.empty();
		this.hostEl = this.contentEl.createDiv({ cls: 'docxidian-host' });
		this.prepareViewHost();
		this.registerHostMetrics();
		this.removeNativeButtonTitles();
		this.removeEditorHelpMenu();
			this.addEditorEditMenuButton();
			this.addEditorSearchMenuButton();
			this.addEditorSettingsMenuButton();
			this.addEditorFileExportAsMenuItem();
			this.addEditorInsertMenuItems();
			this.normalizeNativeEditorMenuActionItems();
		this.trackEditorHoverState();
		this.registerEditorSaveInterceptor();
		this.registerEditorListAwareCopyInterceptor();
		this.registerSaveShortcut();
		this.registerFindShortcut();
		this.registerEditorDropdownScrollGuard();
		this.render();
	}

	async onClose() {
		debugLog('view', 'Closing DOCX view', { file: this.file?.path });
		await this.promptToSaveIfDirty();
		this.reactMount?.unmount();
		this.reactMount = null;
		this.reactMountLoading = false;
		this.hostResizeObserver?.disconnect();
		this.hostResizeObserver = null;
		this.titleObserver?.disconnect();
		this.titleObserver = null;
		this.helpMenuObserver?.disconnect();
		this.helpMenuObserver = null;
		this.searchMenuObserver?.disconnect();
		this.searchMenuObserver = null;
		this.editMenuObserver?.disconnect();
		this.editMenuObserver = null;
		this.fileMenuObserver?.disconnect();
		this.fileMenuObserver = null;
		this.insertMenuObserver?.disconnect();
		this.insertMenuObserver = null;
		this.nativeMenuStyleObserver?.disconnect();
		this.nativeMenuStyleObserver = null;
		this.settingsMenuObserver?.disconnect();
		this.settingsMenuObserver = null;
		this.closeEditorOptionSearchMenu();
		this.closeEditorEditMenu();
		this.closeEditorSettingsMenu();
		activeDocument.body.classList.remove('docxidian-editor-hovering');
		this.hostEl = null;
		this.buffer = null;
		this.error = null;
		this.isDirty = false;
		this.documentSession += 1;
		this.lastKnownFileSignature = null;
		this.backupCreatedForOpenFile = false;
		this.reserveReviewSidebar = false;
	}

	async onLoadFile(file: TFile) {
		infoLog('file', `Loading ${file.path}`, {
			mtime: file.stat.mtime,
			size: file.stat.size,
		});
		await this.promptToSaveIfDirty();
		this.isLoading = true;
		this.error = null;
		this.buffer = null;
		this.isDirty = false;
		this.lastKnownFileSignature = null;
		this.backupCreatedForOpenFile = false;
		this.reserveReviewSidebar = false;
		this.render();

		try {
			this.buffer = await this.app.vault.readBinary(file);
			this.lastKnownFileSignature = await this.readFileSignature(file);
			infoLog('file', `Loaded ${file.path}`, {
				bytes: this.buffer.byteLength,
				signature: this.lastKnownFileSignature,
			});
		} catch (readError) {
			const message = readError instanceof Error ? readError.message : 'Unknown read error';
			this.error = `Could not load ${file.name}: ${message}`;
			errorLog('file', this.error, readError);
			new Notice(this.error);
		} finally {
			this.isLoading = false;
			this.render();
		}

		void this.updateReviewSidebarReservation();
	}

	async onUnloadFile(_file: TFile) {
		debugLog('file', `Unloading ${_file.path}`);
		await this.promptToSaveIfDirty();
		this.buffer = null;
		this.error = null;
		this.isDirty = false;
		this.lastKnownFileSignature = null;
		this.backupCreatedForOpenFile = false;
		this.reserveReviewSidebar = false;
		this.render();
	}

	async onRename(file: TFile) {
		await super.onRename(file);
		this.lastKnownFileSignature = await this.readFileSignature(file);
		infoLog('file', `File renamed or moved to ${file.path}`, {
			signature: this.lastKnownFileSignature,
		});
		this.render();
	}

	async saveCurrentDocument() {
		debugLog('save', 'Save requested', { file: this.file?.path, isLoading: this.isLoading });
		if (!this.file) {
			new Notice('No docx file is open.');
			return false;
		}

		if (this.isLoading) {
			new Notice(`Still loading ${this.file.name}.`);
			return false;
		}

		const saved = await this.getReactHandle()?.save() ?? false;
		if (saved) {
			this.isDirty = false;
			infoLog('save', `Save completed for ${this.file.path}`);
		} else {
			warnLog('save', `Save did not complete for ${this.file.path}`);
		}

		return saved;
	}

	async saveCurrentDocumentAs() {
		const file = this.file;
		debugLog('copy', 'Save as requested', { file: file?.path });
		if (!file) {
			new Notice('Open a docx file to save a copy.');
			return false;
		}

		if (this.isLoading) {
			new Notice(`Still loading ${file.name}.`);
			return false;
		}

		const initialPath = this.getAvailableCopyPath(file);
		const chosenPath = await new Promise<DocxPathChoice>((resolve) => {
			new DocxPathModal(
				this.app,
				'Save as',
				'Create a new DOCX in this vault. Existing files will not be overwritten.',
				initialPath,
				'Save as',
				resolve,
			).open();
		});

		if (!chosenPath) {
			return false;
		}

		return this.createCurrentDocumentCopy(chosenPath, 'Saved as');
	}

	async duplicateCurrentDocument() {
		const file = this.file;
		debugLog('copy', 'Duplicate requested', { file: file?.path });
		if (!file) {
			new Notice('Open a docx file to duplicate it.');
			return false;
		}

		if (this.isLoading) {
			new Notice(`Still loading ${file.name}.`);
			return false;
		}

		return this.createCurrentDocumentCopy(this.getAvailableCopyPath(file), 'Duplicated to');
	}

	async exportCurrentDocumentAs(initialFormat: DocxExportFormatId = DEFAULT_EXPORT_FORMAT) {
		const file = this.file;
		debugLog('copy', 'Export as requested', { file: file?.path, initialFormat });
		if (!file) {
			new Notice('Open a docx file to export it.');
			return false;
		}

		if (this.isLoading) {
			new Notice(`Still loading ${file.name}.`);
			return false;
		}

		const initialPath = this.getAvailableExportPath(file, initialFormat);
		const initialName = initialPath.split('/').pop() ?? initialPath;
		const choice = await new Promise<DocxExportChoice>((resolve) => {
			new DocxExportModal(
				this.app,
				initialName,
				initialFormat,
				resolve,
			).open();
		});

		if (!choice) {
			return false;
		}

		const exportPath = this.getSiblingExportPath(file, choice.name, choice.format);
		if (!exportPath) {
			new Notice('Enter a file name.');
			return false;
		}

		return this.createCurrentDocumentExport(exportPath, choice.format);
	}

	async findHiddenText() {
		const file = this.file;
		debugLog('security', 'Find Hidden Text requested', { file: file?.path });
		if (!file) {
			new Notice('Open a docx file to scan it for hidden text.');
			return false;
		}

		if (this.isLoading) {
			new Notice(`Still loading ${file.name}.`);
			return false;
		}

		try {
			const liveBuffer = await this.getReactHandle()?.exportBuffer({ preserveAutosave: true });
			const scanBuffer = liveBuffer ?? this.buffer;
			if (!scanBuffer) {
				new Notice('No loaded DOCX data is available to scan.');
				return false;
			}

			const result = await findHiddenDocxText(scanBuffer);
			infoLog('security', `Hidden text scan finished for ${file.path}`, {
				findings: result.findings.length,
				partsScanned: result.partsScanned,
			});
			new HiddenTextScanModal(this.app, file.name, result.findings, result.partsScanned).open();
			if (result.findings.length > 0) {
				new Notice(`Found ${result.findings.length} suspicious hidden text item(s).`);
			}
			return true;
		} catch (error) {
			errorLog('security', `Could not scan ${file.path} for hidden text`, error);
			const message = error instanceof Error ? error.message : 'Unknown error';
			new Notice(`Could not scan hidden text: ${message}`);
			return false;
		}
	}

	openFindDialog() {
		if (!this.file || this.isLoading) {
			new Notice('Open a loaded docx file to search it.');
			return;
		}

		const handle = this.getReactHandle();
		if (handle) {
			handle.openFind();
			return;
		}

		window.setTimeout(() => this.getReactHandle()?.openFind(), 50);
	}

	openFindReplaceDialog() {
		if (!this.file || this.isLoading) {
			new Notice('Open a loaded docx file to search it.');
			return;
		}

		this.getReactHandle()?.openFindReplace();
	}

	openImagePicker() {
		const editor = this.getReactHandle();
		if (!editor) {
			new Notice('The DOCX editor is still loading.');
			return;
		}

		editor.openImagePicker();
	}

	openCustomTableDialog() {
		const editor = this.getReactHandle();
		if (!editor) {
			new Notice('The DOCX editor is still loading.');
			return;
		}

		editor.openCustomTableDialog();
	}

	openFontPicker() {
		const editor = this.getReactHandle();
		if (!editor) {
			new Notice('The DOCX editor is still loading.');
			return;
		}

		editor.openFontPicker();
	}

	setEditorMode(mode: 'editing' | 'suggesting' | 'viewing') {
		const editor = this.getReactHandle();
		if (!editor) {
			new Notice('The DOCX editor is still loading.');
			return;
		}

		editor.setMode(mode);
	}

	setEditorZoom(zoom: number) {
		const editor = this.getReactHandle();
		if (!editor) {
			new Notice('The DOCX editor is still loading.');
			return;
		}

		editor.setZoom(zoom);
	}

	pasteFromClipboard(preserveFormatting: boolean) {
		const editor = this.getReactHandle();
		if (!editor) {
			new Notice('The DOCX editor is still loading.');
			return;
		}

		void editor.pasteFromClipboard({ preserveFormatting }).then((pasted) => {
			if (!pasted) {
				new Notice('Nothing was pasted. Check that the clipboard contains text.');
			}
		});
	}

	private getLiveEditorOptionSearchItems(): EditorOptionSearchControlItem[] {
		const root = this.hostEl?.querySelector<HTMLElement>('.ep-root');
		if (!root) {
			return [];
		}

		const items: EditorOptionSearchControlItem[] = [];
		const seen = new Set<string>();
		const controls = root.querySelectorAll<HTMLElement>('button, [role="button"], [role="menuitem"], [role="option"]');
		controls.forEach((element, index) => {
			if (shouldSkipEditorOptionControl(element) || !isVisibleEditorOptionControl(element)) {
				return;
			}

			const label = getEditorOptionControlLabel(element);
			if (!isSearchableEditorOptionLabel(label)) {
				return;
			}

			const key = normalizeMenuText(label);
			if (seen.has(key)) {
				return;
			}

			seen.add(key);
			items.push({
				kind: 'control',
				id: `control-live:${index}:${key}`,
				label,
				keywords: getEditorOptionControlKeywords(element, label),
				element,
			});
		});

		return items;
	}

	private getEditorOptionSearchItems(): EditorOptionSearchItem[] {
		const items: EditorOptionSearchItem[] = [...EDITOR_OPTION_SEARCH_BASE_ITEMS];
		const seen = new Set(items.map(item => normalizeMenuText(item.label)));

		for (const item of this.getLiveEditorOptionSearchItems()) {
			const key = normalizeMenuText(item.label);
			if (!seen.has(key)) {
				seen.add(key);
				items.push(item);
			}
		}

		return items;
	}

	private activateEditorControl(element: HTMLElement) {
		if (element.getAttribute('aria-disabled') === 'true' || ('disabled' in element && element.disabled === true)) {
			new Notice(`${getEditorOptionControlLabel(element)} is not available right now.`);
			return false;
		}

		element.focus({ preventScroll: true });
		element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
		element.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
		element.click();
		return true;
	}

	private findEditorControlByLabels(labels: readonly string[]) {
		const root = this.hostEl?.querySelector<HTMLElement>('.ep-root');
		if (!root) {
			return null;
		}

			const normalizedLabels = labels.map(label => normalizeMenuText(cleanEditorOptionLabel(label))).filter(Boolean);
			const controls = root.querySelectorAll<HTMLElement>('button, [role="button"], [role="menuitem"], [role="option"]');

			for (const element of Array.from(controls)) {
			if (shouldSkipEditorOptionControl(element) || !isVisibleEditorOptionControl(element)) {
				continue;
			}

			const candidateLabels = [
				getEditorOptionControlLabel(element),
				element.dataset.testid ?? '',
				element.textContent ?? '',
				element.dataset.docxidianNativeTitle ?? '',
				element.dataset.docxidianTooltipTitle ?? '',
			].map(label => normalizeMenuText(cleanEditorOptionLabel(label))).filter(Boolean);

			if (candidateLabels.some(candidate => normalizedLabels.some(label => (
				candidate === label
				|| candidate.startsWith(label)
				|| label.startsWith(candidate)
			)))) {
				return element;
			}
		}

		return null;
	}

	private clickEditorControlByLabels(labels: readonly string[]) {
		const control = this.findEditorControlByLabels(labels);
		if (!control) {
			new Notice(`${labels[0] ?? 'That option'} is not available right now.`);
			return false;
		}

		return this.activateEditorControl(control);
	}

	private findEditorMenuButton(menuLabel: string) {
		const root = this.hostEl?.querySelector<HTMLElement>('.ep-root');
		if (!root) {
			return null;
		}

		const normalizedMenuLabel = normalizeMenuText(menuLabel);
		for (const menuItem of Array.from(root.querySelectorAll<HTMLElement>('[role="menubar"] > div'))) {
			const button = menuItem.querySelector<HTMLButtonElement>(':scope > button');
			const label = normalizeMenuText(button?.textContent ?? '');
			if (button && label === normalizedMenuLabel) {
				return button;
			}
		}

		return null;
	}

	private clickEditorMenuOption(menuLabel: string, optionLabels: readonly string[]) {
		const menuButton = this.findEditorMenuButton(menuLabel);
		if (!menuButton) {
			new Notice(`${menuLabel} menu is not available right now.`);
			return;
		}

		this.activateEditorControl(menuButton);
		window.setTimeout(() => {
			if (!this.clickEditorControlByLabels(optionLabels)) {
				activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
			}
		});
	}

	private runEditorOptionSearchItem(item: EditorOptionSearchItem) {
		if (item.kind === 'control') {
			this.activateEditorControl(item.element);
			return;
		}

		if (item.kind === 'control-query') {
			this.clickEditorControlByLabels(item.labels);
			return;
		}

		switch (item.actionId) {
			case 'save':
				void this.saveCurrentDocument();
				break;
			case 'save-as':
				void this.saveCurrentDocumentAs();
				break;
			case 'duplicate':
				void this.duplicateCurrentDocument();
				break;
			case 'paste':
				this.pasteFromClipboard(true);
				break;
			case 'paste-without-formatting':
				this.pasteFromClipboard(false);
				break;
			case 'export-pdf':
				void this.exportCurrentDocumentAs('pdf');
				break;
			case 'export-docx':
				void this.exportCurrentDocumentAs('docx');
				break;
			case 'export-html':
				void this.exportCurrentDocumentAs('html');
				break;
			case 'export-txt':
				void this.exportCurrentDocumentAs('txt');
				break;
			case 'export-md':
				void this.exportCurrentDocumentAs('md');
				break;
			case 'export-rtf':
				void this.exportCurrentDocumentAs('rtf');
				break;
			case 'find':
				this.openFindDialog();
				break;
			case 'find-replace':
				this.openFindReplaceDialog();
				break;
			case 'insert-image':
				this.openImagePicker();
				break;
			case 'custom-table':
				this.openCustomTableDialog();
				break;
			case 'import-font':
				this.openFontPicker();
				break;
			case 'find-hidden-text':
				void this.findHiddenText();
				break;
			case 'page-setup': {
				const labels = getEditorMenuLabels(this.getEditorLocale());
				this.clickEditorMenuOption(labels.file, labels.pageSetup);
				break;
			}
			case 'page-break': {
				const labels = getEditorMenuLabels(this.getEditorLocale());
				this.clickEditorMenuOption(labels.insert, labels.pageBreak);
				break;
			}
			case 'table-of-contents': {
				const labels = getEditorMenuLabels(this.getEditorLocale());
				this.clickEditorMenuOption(labels.insert, labels.tableOfContents);
				break;
			}
			case 'left-to-right': {
				const labels = getEditorMenuLabels(this.getEditorLocale());
				this.clickEditorMenuOption(labels.format, labels.leftToRight);
				break;
			}
			case 'right-to-left': {
				const labels = getEditorMenuLabels(this.getEditorLocale());
				this.clickEditorMenuOption(labels.format, labels.rightToLeft);
				break;
			}
			case 'mode-editing':
				this.setEditorMode('editing');
				break;
			case 'mode-suggesting':
				this.setEditorMode('suggesting');
				break;
			case 'mode-viewing':
				this.setEditorMode('viewing');
				break;
			case 'zoom-75':
				this.setEditorZoom(0.75);
				break;
			case 'zoom-100':
				this.setEditorZoom(1);
				break;
			case 'zoom-125':
				this.setEditorZoom(1.25);
				break;
		}
	}

	refreshSettings() {
		this.render();
	}

	private async updateReviewSidebarReservation() {
		const buffer = this.buffer;
		const file = this.file;
		if (!buffer || !file) {
			return;
		}

		try {
			debugLog('review', `Inspecting review markup for ${file.path}`);
			const { hasReviewMarkup } = await loadDocxEditorChunk();
			const hasMarkup = await hasReviewMarkup(buffer);
			if (buffer !== this.buffer || file !== this.file) {
				debugLog('review', `Discarded stale review markup result for ${file.path}`);
				return;
			}

			this.reserveReviewSidebar = hasMarkup;
			infoLog('review', `Review markup inspection finished for ${file.path}`, { hasMarkup });
			this.render();
		} catch (error) {
			this.reserveReviewSidebar = false;
			errorLog('review', 'Could not inspect DOCX review markup.', error);
		}
	}

	private async saveFile(buffer: ArrayBuffer) {
		const file = this.file;
		if (!file) {
			throw new Error('No docx file is open.');
		}

		infoLog('save', `Writing ${file.path}`, { bytes: buffer.byteLength });
		if (await this.hasFileChangedOnDisk(file)) {
			warnLog('save', `Detected external change before saving ${file.path}`);
			const choice = await this.promptForExternalChange(file.name);
			if (choice !== 'overwrite') {
				warnLog('save', `Save canceled after external change warning for ${file.path}`);
				throw new Error('Save canceled because the file changed on disk.');
			}
			warnLog('save', `External change warning overwritten for ${file.path}`);
		}

		if (this.getCreateBackupsBeforeSave()) {
			await this.createBackupBeforeOverwrite(file);
		}

		await this.app.vault.modifyBinary(file, buffer);
		this.buffer = buffer;
		this.lastKnownFileSignature = await this.readFileSignature(file);
		this.isDirty = false;
		infoLog('save', `Wrote ${file.path}`, {
			signature: this.lastKnownFileSignature,
		});
	}

	private async createCurrentDocumentCopy(path: string, successPrefix: string, options: { openFile?: boolean } = {}) {
		const normalizedPath = this.normalizeDocxPath(path, this.file?.parent?.path);
		debugLog('copy', 'Creating DOCX copy', {
			requestedPath: path,
			normalizedPath,
			openFile: options.openFile !== false,
		});
		if (!normalizedPath) {
			new Notice('Enter a DOCX file path.');
			return false;
		}

		const outputPath = await this.resolveOutputPathConflict(normalizedPath);
		if (!outputPath) {
			return false;
		}

		const editor = this.getReactHandle();
		if (!editor) {
			new Notice('The docx editor is not ready yet.');
			return false;
		}

		const buffer = await editor.exportBuffer();
		if (!buffer) {
			return false;
		}

		let newFile: TFile;
		try {
			await this.ensureParentFolders(outputPath.path);
			if (outputPath.existingFile) {
				await this.app.vault.modifyBinary(outputPath.existingFile, buffer);
				newFile = outputPath.existingFile;
			} else {
				newFile = await this.app.vault.createBinary(outputPath.path, buffer);
			}
			infoLog('copy', `${outputPath.replace ? 'Replaced' : 'Created'} ${newFile.path}`, { bytes: buffer.byteLength });
		} catch (copyError) {
			const message = copyError instanceof Error ? copyError.message : 'Unknown copy error';
			errorLog('copy', `Could not create ${outputPath.path}`, copyError);
			new Notice(`Could not create ${outputPath.path}: ${message}`);
			return false;
		}

		if (options.openFile !== false) {
			const wasDirty = this.isDirty;
			try {
				this.isDirty = false;
				await this.leaf.openFile(newFile);
			} catch (openError) {
				this.isDirty = wasDirty;
				const message = openError instanceof Error ? openError.message : 'Unknown open error';
				errorLog('copy', `Created ${newFile.path}, but could not open it`, openError);
				new Notice(`Created ${newFile.path}, but could not open it: ${message}`);
				return true;
			}
		}

		new Notice(`${outputPath.replace ? 'Replaced' : successPrefix} ${newFile.path}`);
		return true;
	}

	private async createCurrentDocumentExport(path: string, formatId: DocxExportFormatId) {
		const normalizedPath = this.normalizeExportPath(path, formatId, this.file?.parent?.path);
		debugLog('copy', 'Creating document export', {
			requestedPath: path,
			normalizedPath,
			format: formatId,
		});
		if (!normalizedPath) {
			new Notice('Enter a file path.');
			return false;
		}

		const outputPath = await this.resolveOutputPathConflict(normalizedPath);
		if (!outputPath) {
			return false;
		}

		const editor = this.getReactHandle();
		if (!editor) {
			new Notice('The docx editor is not ready yet.');
			return false;
		}

		try {
			await this.ensureParentFolders(outputPath.path);
			new Notice(`Exporting to ${outputPath.path}...`);
			let exportContent: ArrayBuffer | ArrayBufferView | string | null = null;
			if (formatId === 'pdf') {
				exportContent = await editor.exportRenderedPdf();
				if (!exportContent) {
					warnLog('copy', 'Rendered PDF export did not finish; no PDF file was written', {
						path: outputPath.path,
					});
					new Notice(`Could not export ${outputPath.path}: formatted PDF rendering failed. No file was written.`);
					return false;
				}
			}
			if (!exportContent) {
				const buffer = await editor.exportBuffer();
				if (!buffer) {
					new Notice(`Could not export ${outputPath.path}: the editor did not return a document.`);
					return false;
				}
				exportContent = await this.createExportContent(buffer, formatId, this.file?.basename ?? 'Document');
			}
			const binaryContent = getBinaryExportContent(exportContent);
			const textContent = typeof exportContent === 'string' ? exportContent : null;
			if (!binaryContent && textContent === null) {
				throw new Error('The editor returned an unsupported export payload.');
			}
			let newFile: TFile;
			if (outputPath.existingFile) {
				if (binaryContent) {
					await this.app.vault.modifyBinary(outputPath.existingFile, binaryContent);
				} else if (textContent !== null) {
					await this.app.vault.modify(outputPath.existingFile, textContent);
				}
				newFile = outputPath.existingFile;
			} else {
				if (binaryContent) {
					newFile = await this.app.vault.createBinary(outputPath.path, binaryContent);
				} else if (textContent !== null) {
					newFile = await this.app.vault.create(outputPath.path, textContent);
				} else {
					throw new Error('The editor returned an unsupported export payload.');
				}
			}
			infoLog('copy', `${outputPath.replace ? 'Replaced' : 'Exported'} ${newFile.path}`, { format: formatId });
			new Notice(`${outputPath.replace ? 'Replaced' : 'Exported to'} ${newFile.path}`);
			return true;
		} catch (exportError) {
			const message = exportError instanceof Error ? exportError.message : 'Unknown export error';
			errorLog('copy', `Could not export ${outputPath.path}`, exportError);
			new Notice(`Could not export ${outputPath.path}: ${message}`);
			return false;
		}
	}

	private async createExportContent(buffer: ArrayBuffer, formatId: DocxExportFormatId, title: string): Promise<ArrayBuffer | string> {
		if (formatId === 'docx') {
			return buffer;
		}

		const text = await extractDocxText(buffer);
		const exportText = text || title;

		switch (formatId) {
			case 'html':
				return createPlainTextHtml(exportText, title);
			case 'txt':
				return `${exportText}\n`;
			case 'md':
				return `${exportText}\n`;
			case 'rtf':
				return createPlainTextRtf(exportText);
			default:
				throw new Error(`Unsupported export format: ${formatId}`);
		}
	}

	private async resolveOutputPathConflict(path: string): Promise<{ path: string; existingFile: TFile | null; replace: boolean } | null> {
		const existingFile = this.app.vault.getAbstractFileByPath(path);
		if (!existingFile) {
			return { path, existingFile: null, replace: false };
		}

		if (!(existingFile instanceof TFile)) {
			new Notice(`${path} already exists and is not a file.`);
			return null;
		}

		const choice = await new Promise<ExistingFileChoice>((resolve) => {
			new ExistingFileModal(this.app, path, resolve).open();
		});

		if (choice === 'replace') {
			return { path, existingFile, replace: true };
		}

		const copyPath = this.getAvailableNumberedPath(path);
		infoLog('copy', `Keeping existing file and writing numbered copy`, {
			originalPath: path,
			copyPath,
		});
		return { path: copyPath, existingFile: null, replace: false };
	}

	private getAvailableNumberedPath(path: string) {
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

	private async renameFile(name: string) {
		const file = this.file;
		if (!file) {
			throw new Error('No docx file is open.');
		}

		const normalizedName = this.normalizeDocxFileName(name);
		if (!normalizedName) {
			throw new Error('Document name cannot be empty.');
		}

		if (normalizedName === file.name) {
			return;
		}

		const folderPath = file.parent?.path;
		const newPath = folderPath && folderPath !== '/' ? `${folderPath}/${normalizedName}` : normalizedName;
		infoLog('file', `Renaming ${file.path} to ${newPath}`);
		await this.app.fileManager.renameFile(file, newPath);
		this.lastKnownFileSignature = await this.readFileSignature(file);
		new Notice(`Renamed to ${normalizedName}`);
	}

	private async readFileSignature(file: TFile): Promise<DocxFileSignature> {
		try {
			const stat = await this.app.vault.adapter.stat(file.path);
			if (stat?.type === 'file') {
				return {
					path: file.path,
					mtime: stat.mtime,
					size: stat.size,
				};
			}
		} catch (error) {
			debugLog('file', `Falling back to cached file stat for ${file.path}`, error);
			// Fall back to Obsidian's cached stat below.
		}

		return {
			path: file.path,
			mtime: file.stat.mtime,
			size: file.stat.size,
		};
	}

	private signaturesMatch(a: DocxFileSignature, b: DocxFileSignature) {
		return a.path === b.path && a.mtime === b.mtime && a.size === b.size;
	}

	private async hasFileChangedOnDisk(file: TFile) {
		if (!this.lastKnownFileSignature) {
			this.lastKnownFileSignature = await this.readFileSignature(file);
			debugLog('save', `Initialized save conflict signature for ${file.path}`, this.lastKnownFileSignature);
			return false;
		}

		const currentSignature = await this.readFileSignature(file);
		const changed = !this.signaturesMatch(this.lastKnownFileSignature, currentSignature);
		debugLog('save', `Compared disk signature for ${file.path}`, {
			changed,
			lastKnown: this.lastKnownFileSignature,
			current: currentSignature,
		});
		return changed;
	}

	private async promptForExternalChange(fileName: string) {
		return new Promise<DocxConflictChoice>((resolve) => {
			new ExternalDocxChangeModal(this.app, fileName, resolve).open();
		});
	}

	private async createBackupBeforeOverwrite(file: TFile) {
		if (this.backupCreatedForOpenFile) {
			debugLog('backup', `Backup already created for this open session: ${file.path}`);
			return;
		}

		const sourceBuffer = await this.app.vault.readBinary(file);
		const backupPath = this.getAvailableBackupPath(file);
		await this.ensureParentFolders(backupPath);
		await this.app.vault.createBinary(backupPath, sourceBuffer);
		this.backupCreatedForOpenFile = true;
		infoLog('backup', `Created backup for ${file.path}`, {
			backupPath,
			bytes: sourceBuffer.byteLength,
		});
	}

	private getAvailableBackupPath(file: TFile) {
		const folderPath = file.parent?.path;
		const backupFolder = normalizePath(`${folderPath && folderPath !== '/' ? `${folderPath}/` : ''}.docxidian-backups`);
		const baseName = file.basename || file.name.replace(/\.docx$/i, '');
		const timestamp = new Date().toISOString().replace(/\.\d{3}Z$/, '').replace('T', ' ').replace(/:/g, '-');

		for (let index = 0; index < 1000; index += 1) {
			const suffix = index === 0 ? '' : ` ${index + 1}`;
			const candidatePath = normalizePath(`${backupFolder}/${baseName} backup ${timestamp}${suffix}.docx`);
			if (!this.app.vault.getAbstractFileByPath(candidatePath)) {
				return candidatePath;
			}
		}

		return normalizePath(`${backupFolder}/${baseName} backup ${Date.now()}.docx`);
	}

	private normalizeDocxFileName(name: string) {
		const trimmedName = name.trim().replace(/[\\/]/g, '-');
		if (!trimmedName || trimmedName === '.docx') {
			return null;
		}

		return trimmedName.toLowerCase().endsWith('.docx') ? trimmedName : `${trimmedName}.docx`;
	}

	private normalizeExportFileName(name: string, formatId: DocxExportFormatId) {
		const normalizedName = withExportExtension(name, formatId);
		return normalizedName || null;
	}

	private normalizeDocxPath(path: string, fallbackFolderPath?: string) {
		const trimmedPath = path.trim().replace(/\\/g, '/');
		if (!trimmedPath || trimmedPath === '.docx' || trimmedPath.endsWith('/')) {
			return null;
		}

		const pathWithExtension = trimmedPath.toLowerCase().endsWith('.docx') ? trimmedPath : `${trimmedPath}.docx`;
		const hasFolder = pathWithExtension.includes('/');
		const parentPath = fallbackFolderPath && fallbackFolderPath !== '/' ? fallbackFolderPath : '';
		const fullPath = parentPath && !hasFolder ? `${parentPath}/${pathWithExtension}` : pathWithExtension;
		const normalizedDocxPath = normalizePath(fullPath);
		const fileName = normalizedDocxPath.split('/').pop();

		return fileName && fileName !== '.docx' ? normalizedDocxPath : null;
	}

	private normalizeExportPath(path: string, formatId: DocxExportFormatId, fallbackFolderPath?: string) {
		const trimmedPath = path.trim().replace(/\\/g, '/');
		if (!trimmedPath || trimmedPath.endsWith('/')) {
			return null;
		}

		const pathParts = trimmedPath.split('/');
		const rawFileName = pathParts.pop() ?? '';
		const normalizedFileName = this.normalizeExportFileName(rawFileName, formatId);
		if (!normalizedFileName) {
			return null;
		}

		const pathWithExtension = [...pathParts, normalizedFileName].filter(Boolean).join('/');
		const hasFolder = pathWithExtension.includes('/');
		const parentPath = fallbackFolderPath && fallbackFolderPath !== '/' ? fallbackFolderPath : '';
		return normalizePath(parentPath && !hasFolder ? `${parentPath}/${pathWithExtension}` : pathWithExtension);
	}

	private getAvailableCopyPath(file: TFile) {
		const folderPath = file.parent?.path;
		const folderPrefix = folderPath && folderPath !== '/' ? `${folderPath}/` : '';
		const baseName = file.basename || file.name.replace(/\.docx$/i, '');

		for (let index = 1; index < 1000; index += 1) {
			const suffix = index === 1 ? 'copy' : `copy ${index}`;
			const candidatePath = normalizePath(`${folderPrefix}${baseName} ${suffix}.docx`);
			if (!this.app.vault.getAbstractFileByPath(candidatePath)) {
				return candidatePath;
			}
		}

		return normalizePath(`${folderPrefix}${baseName} copy ${Date.now()}.docx`);
	}

	private getAvailableExportPath(file: TFile, formatId: DocxExportFormatId) {
		const folderPath = file.parent?.path;
		const folderPrefix = folderPath && folderPath !== '/' ? `${folderPath}/` : '';
		const baseName = file.basename || file.name.replace(/\.docx$/i, '');
		const extension = getExportFormat(formatId).extension;
		const preferredPath = normalizePath(`${folderPrefix}${baseName}.${extension}`);

		if (preferredPath.toLowerCase() === normalizePath(file.path).toLowerCase()) {
			return this.getAvailableNumberedPath(preferredPath);
		}

		return preferredPath;
	}

	private getSiblingDocxPath(file: TFile, name: string) {
		const normalizedName = this.normalizeDocxFileName(name);
		if (!normalizedName) {
			return null;
		}

		const folderPath = file.parent?.path;
		const folderPrefix = folderPath && folderPath !== '/' ? `${folderPath}/` : '';
		return normalizePath(`${folderPrefix}${normalizedName}`);
	}

	private getSiblingExportPath(file: TFile, name: string, formatId: DocxExportFormatId) {
		const normalizedName = this.normalizeExportFileName(name, formatId);
		if (!normalizedName) {
			return null;
		}

		const folderPath = file.parent?.path;
		const folderPrefix = folderPath && folderPath !== '/' ? `${folderPath}/` : '';
		return normalizePath(`${folderPrefix}${normalizedName}`);
	}

	private async ensureParentFolders(path: string) {
		const folderPath = path.split('/').slice(0, -1).join('/');
		if (!folderPath) {
			return;
		}

		const segments = folderPath.split('/').filter(Boolean);
		let currentPath = '';

		for (const segment of segments) {
			currentPath = currentPath ? `${currentPath}/${segment}` : segment;
			const existingFile = this.app.vault.getAbstractFileByPath(currentPath);

			if (existingFile instanceof TFile) {
				throw new Error(`${currentPath} is a file.`);
			}

			if (!existingFile) {
				await this.app.vault.createFolder(currentPath);
			}
		}
	}

	private async promptToSaveIfDirty() {
		if (!this.isDirty || !this.file) {
			return;
		}

		warnLog('save', `Prompting for unsaved changes in ${this.file.path}`);
		const choice = await new Promise<UnsavedDocxChoice>((resolve) => {
			new UnsavedDocxModal(this.app, this.file?.name ?? 'Document', resolve).open();
		});

		if (choice === 'save') {
			infoLog('save', `Saving dirty document before closing ${this.file.path}`);
			await this.saveCurrentDocument();
		} else {
			warnLog('save', `Discarding unsaved changes in ${this.file.path}`);
		}
	}

	private prepareViewHost() {
		if (!this.hostEl) {
			return;
		}

		this.hostEl.setCssProps({
			'--docxidian-fixed-left-offset': '0px',
			'--docxidian-fixed-top-offset': '0px',
		});
	}

	private registerHostMetrics() {
		const updateHostMetrics = () => {
			if (!this.hostEl) {
				return;
			}

			const fixedProbe = this.hostEl.createDiv({ cls: 'docxidian-fixed-probe' });
			const fixedRect = fixedProbe.getBoundingClientRect();
			fixedProbe.remove();

			this.hostEl.setCssProps({
				'--docxidian-fixed-left-offset': `${Math.round(fixedRect.left)}px`,
				'--docxidian-fixed-top-offset': `${Math.round(fixedRect.top)}px`,
			});
		};

		updateHostMetrics();
		this.registerDomEvent(window, 'resize', updateHostMetrics);
		this.registerDomEvent(window, 'scroll', updateHostMetrics, true);
		this.hostResizeObserver = new ResizeObserver(updateHostMetrics);
		this.hostResizeObserver.observe(this.contentEl);
		this.register(() => {
			this.hostResizeObserver?.disconnect();
			this.hostResizeObserver = null;
		});
	}

	private removeNativeButtonTitles() {
		if (!this.hostEl) {
			return;
		}

		const removeTitles = () => {
			this.hostEl?.querySelectorAll('.ep-root button[title]').forEach((button) => {
				if (isHTMLElement(button)) {
					const title = button.getAttribute('title');
					if (title) {
						button.dataset.docxidianNativeTitle = title;
					}
				}
				button.removeAttribute('title');
			});
		};

		removeTitles();
		this.titleObserver = new MutationObserver(removeTitles);
		this.titleObserver.observe(this.hostEl, {
			attributes: true,
			attributeFilter: ['title'],
			childList: true,
			subtree: true,
		});
		this.register(() => {
			this.titleObserver?.disconnect();
			this.titleObserver = null;
		});
	}

	private removeEditorHelpMenu() {
		if (!this.hostEl) {
			return;
		}

		const removeHelpMenu = () => {
			const helpLabel = getEditorMenuLabels(this.getEditorLocale()).help;

			this.hostEl?.querySelectorAll('.ep-root [role="menubar"] > div').forEach((menuItem) => {
				const button = menuItem.querySelector(':scope > button');
				const label = normalizeMenuText(button?.textContent ?? '');
				if (label === helpLabel) {
					menuItem.remove();
				}
			});
		};

		removeHelpMenu();
		this.helpMenuObserver = new MutationObserver(removeHelpMenu);
		this.helpMenuObserver.observe(this.hostEl, {
			childList: true,
			subtree: true,
		});
		this.register(() => {
			this.helpMenuObserver?.disconnect();
			this.helpMenuObserver = null;
		});
	}

	private addEditorEditMenuButton() {
		if (!this.hostEl) {
			return;
		}

		const addEditButton = () => {
			if (!this.hostEl) {
				return;
			}

			const labels = getEditorMenuLabels(this.getEditorLocale());
			const menubar = this.hostEl.querySelector<HTMLElement>('.ep-root [role="menubar"]');
			if (!menubar || menubar.querySelector('[data-docxidian-edit-menu-item]')) {
				return;
			}

			const menuItems = Array.from(menubar.children).filter((child): child is HTMLElement => isHTMLElement(child));
			const findTopLevelMenu = (label: string) => menuItems.find((item) => {
				const button = item.querySelector(':scope > button');
				return normalizeMenuText(button?.textContent ?? '') === label;
			});
			const fileWrapper = findTopLevelMenu(labels.file);
			const formatWrapper = findTopLevelMenu(labels.format);
			const existingEditWrapper = findTopLevelMenu(labels.edit);
			const sourceWrapper = existingEditWrapper
				?? menuItems.find((child) => (
					!child.matches('[data-docxidian-edit-menu-item], [data-docxidian-search-menu-item], [data-docxidian-settings-menu-item]')
					&& Boolean(child.querySelector(':scope > button'))
				));
			const wrapper = existingEditWrapper
				?? (sourceWrapper ? sourceWrapper.cloneNode(true) as HTMLElement : activeDocument.createElement('div'));
			wrapper.dataset.docxidianEditMenuItem = 'true';
			wrapper.dataset.docxidianNoToolbarTooltip = 'true';
			wrapper.addClass('docxidian-edit-menu-item');
			wrapper.setCssProps({ position: 'relative' });
			wrapper.removeAttribute('id');

			let button = wrapper.querySelector<HTMLButtonElement>(':scope > button');
			Array.from(wrapper.children).forEach((child) => {
				if (child !== button) {
					child.remove();
				}
			});
			if (!button) {
				button = activeDocument.createElement('button');
				wrapper.appendChild(button);
			}

			button.type = 'button';
			button.textContent = 'Edit';
			button.dataset.docxidianNoToolbarTooltip = 'true';
			button.addClasses(['docxidian-search-menu-button', 'docxidian-edit-menu-button']);
			button.removeAttribute('aria-haspopup');
			button.removeAttribute('data-state');
			button.removeAttribute('id');
			button.setAttribute('aria-label', 'Edit');
			button.setAttribute('aria-expanded', 'false');
			button.setAttribute('role', 'menuitem');
			button.addEventListener('mousedown', (evt) => {
				evt.preventDefault();
			});
			button.addEventListener('click', (evt) => {
				evt.preventDefault();
				evt.stopImmediatePropagation();
				evt.stopPropagation();
				this.openEditorEditMenu(wrapper);
				button?.setAttribute('aria-expanded', this.editorEditPopoverEl ? 'true' : 'false');
			});

			if (fileWrapper && fileWrapper.parentElement === menubar) {
				fileWrapper.after(wrapper);
			} else if (formatWrapper && formatWrapper.parentElement === menubar) {
				menubar.insertBefore(wrapper, formatWrapper);
			} else {
				menubar.prepend(wrapper);
			}
		};

		addEditButton();
		this.editMenuObserver = new MutationObserver(addEditButton);
		this.editMenuObserver.observe(this.hostEl, {
			childList: true,
			subtree: true,
		});
		this.register(() => {
			this.editMenuObserver?.disconnect();
			this.editMenuObserver = null;
		});
	}

	private closeEditorEditMenu() {
		this.editorEditCleanup?.();
		this.editorEditCleanup = null;
		this.editorEditPopoverEl?.remove();
		this.editorEditPopoverEl = null;
		this.hostEl?.querySelector('[data-docxidian-edit-menu-item] > button')?.setAttribute('aria-expanded', 'false');
	}

	private openEditorEditMenu(anchorEl: HTMLElement) {
		if (this.editorEditPopoverEl && anchorEl.contains(this.editorEditPopoverEl)) {
			this.closeEditorEditMenu();
			return;
		}

		this.closeEditorOptionSearchMenu();
		this.closeEditorSettingsMenu();
		this.closeEditorEditMenu();

		const popoverEl = anchorEl.createDiv({ cls: 'docxidian-edit-menu docxidian-option-search-menu' });
		popoverEl.setAttribute('role', 'menu');
		this.editorEditPopoverEl = popoverEl;

		const addAction = (label: string, preserveFormatting: boolean) => {
			const button = popoverEl.createEl('button', {
				cls: 'docxidian-option-search-result docxidian-file-menu-button',
				text: label,
				type: 'button',
			});
			button.setAttribute('role', 'menuitem');
			button.addEventListener('mousedown', (evt) => {
				evt.preventDefault();
			});
			button.addEventListener('click', (evt) => {
				evt.preventDefault();
				evt.stopPropagation();
				this.closeEditorEditMenu();
				this.pasteFromClipboard(preserveFormatting);
			});
		};

		addAction('Paste', true);
		addAction('Paste without formatting', false);

		const handleOutsidePointer = (evt: MouseEvent) => {
			if (evt.target instanceof Node && !popoverEl.contains(evt.target) && !anchorEl.contains(evt.target)) {
				this.closeEditorEditMenu();
			}
		};
		const handleKeyDown = (evt: KeyboardEvent) => {
			if (evt.key === 'Escape') {
				evt.preventDefault();
				this.closeEditorEditMenu();
			}
		};

		activeDocument.addEventListener('mousedown', handleOutsidePointer, true);
		popoverEl.addEventListener('keydown', handleKeyDown);
		this.editorEditCleanup = () => {
			activeDocument.removeEventListener('mousedown', handleOutsidePointer, true);
			popoverEl.removeEventListener('keydown', handleKeyDown);
		};
	}

	private addEditorSearchMenuButton() {
		if (!this.hostEl) {
			return;
		}

		const addSearchButton = () => {
			if (!this.hostEl) {
				return;
			}

			const menubar = this.hostEl.querySelector<HTMLElement>('.ep-root [role="menubar"]');
			if (!menubar || menubar.querySelector('[data-docxidian-search-menu-item]')) {
				return;
			}

				const sourceWrapper = Array.from(menubar.children).find((child): child is HTMLElement => (
					isHTMLElement(child)
					&& !child.matches('[data-docxidian-edit-menu-item], [data-docxidian-search-menu-item], [data-docxidian-settings-menu-item]')
					&& Boolean(child.querySelector(':scope > button'))
				));
				const wrapper = sourceWrapper
					? sourceWrapper.cloneNode(true) as HTMLElement
					: activeDocument.createElement('div');
				wrapper.dataset.docxidianSearchMenuItem = 'true';
				wrapper.dataset.docxidianNoToolbarTooltip = 'true';
				wrapper.addClass('docxidian-search-menu-item');
				wrapper.setCssProps({ position: 'relative' });
				wrapper.removeAttribute('id');

				let button = wrapper.querySelector<HTMLButtonElement>(':scope > button');
				Array.from(wrapper.children).forEach((child) => {
					if (child !== button) {
						child.remove();
					}
				});
				if (!button) {
					button = activeDocument.createElement('button');
					wrapper.appendChild(button);
				}
				button.type = 'button';
				button.textContent = 'Search';
				button.dataset.docxidianNoToolbarTooltip = 'true';
				button.addClass('docxidian-search-menu-button');
				button.removeAttribute('aria-expanded');
				button.removeAttribute('aria-haspopup');
				button.removeAttribute('data-state');
				button.removeAttribute('id');
				button.setAttribute('aria-label', 'Search');
			button.setAttribute('role', 'menuitem');
			button.addEventListener('mousedown', (evt) => {
				evt.preventDefault();
			});
			button.addEventListener('click', (evt) => {
					evt.preventDefault();
					evt.stopImmediatePropagation();
					evt.stopPropagation();
					this.openEditorOptionSearchMenu(wrapper);
				});

				const settingsWrapper = menubar.querySelector('[data-docxidian-settings-menu-item]');
				if (settingsWrapper) {
					menubar.insertBefore(wrapper, settingsWrapper);
				} else {
					menubar.appendChild(wrapper);
				}
			};

		addSearchButton();
		this.searchMenuObserver = new MutationObserver(addSearchButton);
		this.searchMenuObserver.observe(this.hostEl, {
			childList: true,
			subtree: true,
		});
		this.register(() => {
			this.searchMenuObserver?.disconnect();
			this.searchMenuObserver = null;
		});
		}

		private closeEditorOptionSearchMenu() {
			this.optionSearchCleanup?.();
			this.optionSearchCleanup = null;
			this.optionSearchPopoverEl?.remove();
			this.optionSearchPopoverEl = null;
		}

		private openEditorOptionSearchMenu(anchorEl: HTMLElement) {
			if (this.optionSearchPopoverEl && anchorEl.contains(this.optionSearchPopoverEl)) {
				this.closeEditorOptionSearchMenu();
				return;
			}

			this.closeEditorOptionSearchMenu();
			this.closeEditorEditMenu();

			let activeIndex = 0;
			const popoverEl = anchorEl.createDiv({ cls: 'docxidian-option-search-menu' });
			const inputEl = popoverEl.createEl('input', {
				cls: 'docxidian-option-search-input',
				type: 'search',
			});
			inputEl.placeholder = 'Search options';
			inputEl.setAttribute('aria-label', 'Search editor options');

			const resultsEl = popoverEl.createDiv({ cls: 'docxidian-option-search-results' });
			this.optionSearchPopoverEl = popoverEl;

			const getMatches = () => {
				const items = this.getEditorOptionSearchItems();
				const query = normalizeMenuText(inputEl.value);
				if (!query) {
					return items;
				}

				return items.filter((item) => {
					const haystack = normalizeMenuText([item.label, ...item.keywords].join(' '));
					return haystack.includes(query);
				});
			};

			const chooseItem = (item: EditorOptionSearchItem) => {
				this.closeEditorOptionSearchMenu();
				this.runEditorOptionSearchItem(item);
			};

			const renderResults = () => {
				const matches = getMatches();
				activeIndex = Math.max(0, Math.min(activeIndex, matches.length - 1));
				resultsEl.empty();

				if (matches.length === 0) {
					resultsEl.createDiv({ cls: 'docxidian-option-search-empty', text: 'No options found' });
					return;
				}

				matches.forEach((item, index) => {
					const button = resultsEl.createEl('button', {
						cls: 'docxidian-option-search-result',
						text: item.label,
						type: 'button',
					});
					button.addClass('docxidian-file-menu-button');
					button.setAttribute('role', 'option');
					button.setAttribute('aria-selected', index === activeIndex ? 'true' : 'false');
					button.addEventListener('mouseenter', () => {
						activeIndex = index;
						renderResults();
					});
					button.addEventListener('click', (evt) => {
						evt.preventDefault();
						evt.stopPropagation();
						chooseItem(item);
					});
				});
			};

			const handleInput = () => {
				activeIndex = 0;
				renderResults();
			};
			const handleKeyDown = (evt: KeyboardEvent) => {
				const matches = getMatches();
				if (evt.key === 'Escape') {
					evt.preventDefault();
					this.closeEditorOptionSearchMenu();
					return;
				}
				if (evt.key === 'ArrowDown') {
					evt.preventDefault();
					activeIndex = matches.length > 0 ? (activeIndex + 1) % matches.length : 0;
					renderResults();
					return;
				}
				if (evt.key === 'ArrowUp') {
					evt.preventDefault();
					activeIndex = matches.length > 0 ? (activeIndex - 1 + matches.length) % matches.length : 0;
					renderResults();
					return;
				}
				if (evt.key === 'Enter') {
					evt.preventDefault();
					const item = matches[activeIndex];
					if (item) {
						chooseItem(item);
					}
				}
			};
			const handleOutsidePointer = (evt: MouseEvent) => {
				if (evt.target instanceof Node && !popoverEl.contains(evt.target) && !anchorEl.contains(evt.target)) {
					this.closeEditorOptionSearchMenu();
				}
			};

			inputEl.addEventListener('input', handleInput);
			inputEl.addEventListener('keydown', handleKeyDown);
			activeDocument.addEventListener('mousedown', handleOutsidePointer, true);
			this.optionSearchCleanup = () => {
				inputEl.removeEventListener('input', handleInput);
				inputEl.removeEventListener('keydown', handleKeyDown);
				activeDocument.removeEventListener('mousedown', handleOutsidePointer, true);
			};

			renderResults();
			window.setTimeout(() => inputEl.focus());
		}

		private addEditorSettingsMenuButton() {
			if (!this.hostEl) {
				return;
			}

			const addSettingsButton = () => {
				if (!this.hostEl) {
					return;
				}

				const menubar = this.hostEl.querySelector<HTMLElement>('.ep-root [role="menubar"]');
				if (!menubar || menubar.querySelector('[data-docxidian-settings-menu-item]')) {
					return;
				}

				const sourceWrapper = Array.from(menubar.children).find((child): child is HTMLElement => (
					isHTMLElement(child)
					&& !child.matches('[data-docxidian-edit-menu-item], [data-docxidian-search-menu-item], [data-docxidian-settings-menu-item]')
					&& Boolean(child.querySelector(':scope > button'))
				));
				const wrapper = sourceWrapper
					? sourceWrapper.cloneNode(true) as HTMLElement
					: activeDocument.createElement('div');
				wrapper.dataset.docxidianSettingsMenuItem = 'true';
				wrapper.dataset.docxidianNoToolbarTooltip = 'true';
				wrapper.addClass('docxidian-settings-menu-item');
				wrapper.setCssProps({ position: 'relative' });
				wrapper.removeAttribute('id');

				let button = wrapper.querySelector<HTMLButtonElement>(':scope > button');
				Array.from(wrapper.children).forEach((child) => {
					if (child !== button) {
						child.remove();
					}
				});
				if (!button) {
					button = activeDocument.createElement('button');
					wrapper.appendChild(button);
				}

				button.type = 'button';
				button.textContent = 'Settings';
				button.dataset.docxidianNoToolbarTooltip = 'true';
				button.addClasses(['docxidian-search-menu-button', 'docxidian-settings-menu-button']);
				button.removeAttribute('aria-haspopup');
				button.removeAttribute('data-state');
				button.removeAttribute('id');
				button.setAttribute('aria-label', 'Settings');
				button.setAttribute('aria-expanded', 'false');
				button.setAttribute('role', 'menuitem');
				button.addEventListener('mousedown', (evt) => {
					evt.preventDefault();
				});
				button.addEventListener('click', (evt) => {
					evt.preventDefault();
					evt.stopImmediatePropagation();
					evt.stopPropagation();
					this.openEditorSettingsMenu(wrapper);
					button?.setAttribute('aria-expanded', this.editorSettingsPopoverEl ? 'true' : 'false');
				});

				const searchWrapper = menubar.querySelector('[data-docxidian-search-menu-item]');
				if (searchWrapper) {
					searchWrapper.after(wrapper);
				} else {
					menubar.appendChild(wrapper);
				}
			};

			addSettingsButton();
			this.settingsMenuObserver = new MutationObserver(addSettingsButton);
			this.settingsMenuObserver.observe(this.hostEl, {
				childList: true,
				subtree: true,
			});
			this.register(() => {
				this.settingsMenuObserver?.disconnect();
				this.settingsMenuObserver = null;
			});
		}

		private closeEditorSettingsMenu() {
			this.editorSettingsCleanup?.();
			this.editorSettingsCleanup = null;
			this.editorSettingsPopoverEl?.remove();
			this.editorSettingsPopoverEl = null;
			this.hostEl?.querySelector('[data-docxidian-settings-menu-item] > button')?.setAttribute('aria-expanded', 'false');
		}

		private openEditorSettingsMenu(anchorEl: HTMLElement) {
			if (this.editorSettingsPopoverEl && anchorEl.contains(this.editorSettingsPopoverEl)) {
				this.closeEditorSettingsMenu();
				return;
			}

			this.closeEditorOptionSearchMenu();
			this.closeEditorEditMenu();
			this.closeEditorSettingsMenu();

			const popoverEl = anchorEl.createDiv({ cls: 'docxidian-editor-settings-menu' });
			popoverEl.setAttribute('role', 'menu');
			this.editorSettingsPopoverEl = popoverEl;
			this.renderEditorSettingsMenu(popoverEl);

			const handleOutsidePointer = (evt: MouseEvent) => {
				if (evt.target instanceof Node && !popoverEl.contains(evt.target) && !anchorEl.contains(evt.target)) {
					this.closeEditorSettingsMenu();
				}
			};
			const handleKeyDown = (evt: KeyboardEvent) => {
				if (evt.key === 'Escape') {
					evt.preventDefault();
					this.closeEditorSettingsMenu();
				}
			};

			activeDocument.addEventListener('mousedown', handleOutsidePointer, true);
			popoverEl.addEventListener('keydown', handleKeyDown);
			this.editorSettingsCleanup = () => {
				activeDocument.removeEventListener('mousedown', handleOutsidePointer, true);
				popoverEl.removeEventListener('keydown', handleKeyDown);
			};
		}

		private renderEditorSettingsMenu(menuEl: HTMLElement) {
			menuEl.empty();
			const settings = this.settingsController.getSettings();

			const addSection = (text: string) => {
				menuEl.createDiv({ cls: 'docxidian-editor-settings-section', text });
			};
			const addDescription = (parentEl: HTMLElement, text: string) => {
				parentEl.createDiv({ cls: 'docxidian-editor-settings-desc', text });
			};
			const addToggle = (
				label: string,
				description: string,
				value: boolean,
				onChange: (nextValue: boolean) => Promise<void>,
			) => {
				const row = menuEl.createEl('label', { cls: 'docxidian-editor-settings-row mod-toggle' });
				const copy = row.createDiv({ cls: 'docxidian-editor-settings-copy' });
				copy.createSpan({ cls: 'docxidian-editor-settings-label', text: label });
				addDescription(copy, description);
				const input = row.createEl('input', {
					cls: 'docxidian-editor-settings-checkbox',
					type: 'checkbox',
				});
				input.checked = value;
				input.addEventListener('change', () => {
					void onChange(input.checked);
				});
			};
			const addButtonRow = (
				label: string,
				description: string,
				buttonLabel: string,
				onClick: () => Promise<void>,
			) => {
				const row = menuEl.createDiv({ cls: 'docxidian-editor-settings-row' });
				const copy = row.createDiv({ cls: 'docxidian-editor-settings-copy' });
				copy.createSpan({ cls: 'docxidian-editor-settings-label', text: label });
				addDescription(copy, description);
				const button = row.createEl('button', {
					cls: 'docxidian-editor-settings-action',
					text: buttonLabel,
					type: 'button',
				});
				button.addEventListener('click', (evt) => {
					evt.preventDefault();
					evt.stopPropagation();
					void onClick();
				});
			};

			menuEl.addEventListener('mousedown', (evt) => evt.stopPropagation());
			menuEl.addEventListener('click', (evt) => evt.stopPropagation());

			addSection('Identity');
			const authorRow = menuEl.createDiv({ cls: 'docxidian-editor-settings-row mod-input' });
			const authorCopy = authorRow.createDiv({ cls: 'docxidian-editor-settings-copy' });
			authorCopy.createSpan({ cls: 'docxidian-editor-settings-label', text: 'Author name' });
			addDescription(authorCopy, 'Used for comments and tracked changes.');
			const authorControl = authorRow.createDiv({ cls: 'docxidian-editor-settings-inline-controls' });
			const authorInput = authorControl.createEl('input', {
				cls: 'docxidian-editor-settings-input',
				type: 'text',
			});
			authorInput.value = settings.authorName;
			authorInput.addEventListener('change', () => {
				void this.settingsController.setAuthorName(authorInput.value);
			});
			const resetAuthor = authorControl.createEl('button', {
				cls: 'docxidian-editor-settings-action',
				text: 'Reset',
				type: 'button',
			});
			resetAuthor.addEventListener('click', (evt) => {
				evt.preventDefault();
				authorInput.value = DEFAULT_SETTINGS.authorName;
				void this.settingsController.setAuthorName(DEFAULT_SETTINGS.authorName);
			});

			addSection('Editor defaults');
			const languageRow = menuEl.createDiv({ cls: 'docxidian-editor-settings-row mod-input' });
			const languageCopy = languageRow.createDiv({ cls: 'docxidian-editor-settings-copy' });
			languageCopy.createSpan({ cls: 'docxidian-editor-settings-label', text: 'Default language' });
			addDescription(languageCopy, 'English is the default language for the editor toolbar, dialogs, and messages.');
			const languageSelect = languageRow.createEl('select', { cls: 'docxidian-editor-settings-select' });
			const selectedLanguage = normalizeDocxidianLanguage(settings.editorLanguage);
			for (const option of DOCXIDIAN_LANGUAGE_OPTIONS) {
				const label = option.code === DEFAULT_LANGUAGE ? `${option.label} (default)` : option.label;
				const optionEl = languageSelect.createEl('option', { text: label, value: option.code });
				optionEl.selected = option.code === selectedLanguage;
			}
			languageSelect.addEventListener('change', () => {
				void this.settingsController.setEditorLanguage(languageSelect.value);
			});

			addToggle('Ruler', 'Show the page ruler above the document body by default.', settings.showRuler, this.settingsController.setShowRuler);
			const zoomRow = menuEl.createDiv({ cls: 'docxidian-editor-settings-row mod-input' });
			const zoomCopy = zoomRow.createDiv({ cls: 'docxidian-editor-settings-copy' });
			zoomCopy.createSpan({ cls: 'docxidian-editor-settings-label', text: 'Default zoom' });
			addDescription(zoomCopy, 'Initial zoom for DOCX files when they open.');
			const zoomSelect = zoomRow.createEl('select', { cls: 'docxidian-editor-settings-select' });
			const selectedZoom = normalizeDefaultZoom(settings.defaultZoom);
			for (const zoom of [0.5, 0.75, 1, 1.25, 1.5, 2]) {
				const optionEl = zoomSelect.createEl('option', {
					text: `${Math.round(zoom * 100)}%`,
					value: String(zoom),
				});
				optionEl.selected = zoom === selectedZoom;
			}
			zoomSelect.addEventListener('change', () => {
				void this.settingsController.setDefaultZoom(Number(zoomSelect.value));
			});

			addSection('Saving');
			addToggle('Autosave', 'Automatically save the document shortly after changes.', settings.autosave, this.settingsController.setAutosave);
			addToggle('Backups', 'Create one timestamped backup before the first overwrite in each open DOCX session.', settings.createBackupsBeforeSave, this.settingsController.setCreateBackupsBeforeSave);

			addSection('Search');
			addToggle('DOCX search index', 'Extract text from DOCX files into a local cache for vault-wide DOCX search.', settings.enableDocxSearchIndex, this.settingsController.setEnableDocxSearchIndex);
			addToggle('Auto-index DOCX changes', 'Keep the DOCX search cache updated when DOCX files are created, edited, renamed, or deleted.', settings.autoIndexDocxSearch, this.settingsController.setAutoIndexDocxSearch);
			addButtonRow('Rebuild DOCX search index', 'Refresh the searchable cache for DOCX files in this vault.', 'Rebuild', this.settingsController.rebuildDocxSearchIndex);

			addSection('File handoff');
			addToggle('Turn off for DOCX files', 'Turns off plugin specifically for DOCX files in favor of another plugin </3', settings.disableDocxFiles, this.settingsController.setDisableDocxFiles);

			addSection('Diagnostics');
			addToggle('Debug logging', 'Print Native PowerPoint Doc Editor diagnostics to the developer console.', settings.debugLogging, this.settingsController.setDebugLogging);
			addButtonRow('Copy DOCX log', 'Copy DOCX-specific Native PowerPoint Doc Editor logs to the clipboard.', 'Copy', () => this.settingsController.copyDocxLog(this.file?.path));
		}

		private addEditorFileExportAsMenuItem() {
		if (!this.hostEl) {
			return;
		}

		const retitleMenuButton = (
			button: HTMLButtonElement,
			label: string,
			sourceLabels: string[],
			options: { showChevron?: boolean } = {},
		) => {
			const labelElement = Array.from(button.children).find((child): child is HTMLElement => (
				isHTMLElement(child)
				&& sourceLabels.some(sourceLabel => textStartsWithMenuLabel(normalizeMenuText(child.textContent ?? ''), sourceLabel))
			));

			if (labelElement) {
				labelElement.textContent = label;
				Array.from(button.children).forEach((child) => {
					if (
						isHTMLElement(child)
						&& child !== labelElement
						&& /^(?:ctrl|cmd|⌘)/.test(normalizeMenuText(child.textContent ?? ''))
					) {
						child.remove();
					}
				});
			} else {
				button.textContent = label;
			}

			button.querySelectorAll('[data-docxidian-export-chevron]').forEach(chevron => chevron.remove());
				if (options.showChevron) {
					const chevron = activeDocument.createElement('span');
					chevron.dataset.docxidianExportChevron = 'true';
					chevron.textContent = '›';
					chevron.addClass('docxidian-export-chevron');
					button.appendChild(chevron);
				}
			};

		const addExportAsItem = () => {
			if (!this.hostEl) {
				return;
			}

			const labels = getEditorMenuLabels(this.getEditorLocale());
			this.hostEl.querySelectorAll<HTMLElement>('.ep-root [role="menubar"] > div').forEach((menuItem) => {
				const menuButton = menuItem.querySelector(':scope > button');
				const menuLabel = normalizeMenuText(menuButton?.textContent ?? '');
				if (menuLabel !== labels.file) {
					return;
				}

				const dropdown = Array.from(menuItem.children).find((child): child is HTMLElement => (
					isHTMLElement(child)
					&& child !== menuButton
					&& Boolean(child.querySelector(':scope > div > button, :scope > button'))
				));
				if (!dropdown) {
					return;
				}

				const itemWrappers = Array.from(dropdown.children).filter((child): child is HTMLElement => isHTMLElement(child));
				const saveWrapper = itemWrappers.find((itemWrapper) => {
					const button = itemWrapper.querySelector(':scope > button');
					const text = normalizeMenuText(button?.textContent ?? '');
					return labels.save.some((label) => textStartsWithMenuLabel(text, label));
				});
				const sourceWrapper = saveWrapper ?? itemWrappers.find((itemWrapper) => itemWrapper.querySelector(':scope > button'));

				let duplicateWrapper = dropdown.querySelector<HTMLElement>('[data-docxidian-duplicate-menu-item]');
				if (!duplicateWrapper) {
					duplicateWrapper = sourceWrapper
						? sourceWrapper.cloneNode(true) as HTMLElement
						: activeDocument.createElement('div');
					duplicateWrapper.dataset.docxidianDuplicateMenuItem = 'true';
					duplicateWrapper.addClasses(['docxidian-file-menu-item', 'docxidian-duplicate-menu-item']);

					const duplicateButton = duplicateWrapper.querySelector('button') ?? duplicateWrapper.createEl('button');
					duplicateButton.type = 'button';
					duplicateButton.addClass('docxidian-file-menu-button');
					retitleMenuButton(duplicateButton, 'Duplicate current DOCX', labels.save);
					duplicateButton.removeAttribute('disabled');
					duplicateButton.removeAttribute('aria-disabled');
					duplicateButton.addEventListener('mousedown', (evt) => {
						evt.preventDefault();
					});
					duplicateButton.addEventListener('click', (evt) => {
						evt.preventDefault();
						evt.stopImmediatePropagation();
						evt.stopPropagation();
						activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
						void this.duplicateCurrentDocument();
					});

					if (saveWrapper) {
						saveWrapper.after(duplicateWrapper);
					} else {
						dropdown.prepend(duplicateWrapper);
					}
				}

				let exportWrapper = dropdown.querySelector<HTMLElement>('[data-docxidian-export-as-menu-item]');
				if (!exportWrapper) {
					exportWrapper = sourceWrapper
						? sourceWrapper.cloneNode(true) as HTMLElement
						: activeDocument.createElement('div');
						exportWrapper.dataset.docxidianExportAsMenuItem = 'true';

						let exportButton = exportWrapper.querySelector('button');
					if (!exportButton) {
						exportButton = activeDocument.createElement('button');
						exportWrapper.appendChild(exportButton);
					}
						exportWrapper.addClasses(['docxidian-file-menu-item', 'docxidian-export-menu-item']);
						exportButton.type = 'button';
						exportButton.addClass('docxidian-file-menu-button');
						retitleMenuButton(exportButton, 'Export as...', labels.save, { showChevron: true });
					exportButton.removeAttribute('disabled');
					exportButton.removeAttribute('aria-disabled');
					exportButton.addEventListener('mousedown', (evt) => {
						evt.preventDefault();
					});
					exportButton.addEventListener('click', (evt) => {
						evt.preventDefault();
						evt.stopImmediatePropagation();
						evt.stopPropagation();
						activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
						void this.exportCurrentDocumentAs();
					});

					const exportSubmenu = exportWrapper.createDiv({ cls: 'docxidian-export-submenu' });
					for (const format of DOCX_EXPORT_FORMATS) {
						const optionWrapper = sourceWrapper
							? sourceWrapper.cloneNode(true) as HTMLElement
							: activeDocument.createElement('div');
						optionWrapper.removeAttribute('data-docxidian-export-as-menu-item');
							const optionButton = optionWrapper.querySelector('button') ?? optionWrapper.createEl('button');
							optionButton.type = 'button';
							optionButton.addClass('docxidian-file-menu-button');
							retitleMenuButton(optionButton, format.label, labels.save);
						optionButton.removeAttribute('disabled');
						optionButton.removeAttribute('aria-disabled');
						optionButton.addEventListener('mousedown', (evt) => {
							evt.preventDefault();
						});
						optionButton.addEventListener('click', (evt) => {
							evt.preventDefault();
							evt.stopImmediatePropagation();
							evt.stopPropagation();
							activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
							void this.exportCurrentDocumentAs(format.id);
						});
						exportSubmenu.appendChild(optionWrapper);
					}

					if (duplicateWrapper.parentElement === dropdown) {
						duplicateWrapper.after(exportWrapper);
					} else if (saveWrapper) {
						saveWrapper.after(exportWrapper);
					} else {
						dropdown.prepend(exportWrapper);
					}
				}

				if (!dropdown.querySelector('[data-docxidian-find-hidden-text-menu-item]')) {
					const hiddenTextWrapper = sourceWrapper
						? sourceWrapper.cloneNode(true) as HTMLElement
						: activeDocument.createElement('div');
						hiddenTextWrapper.dataset.docxidianFindHiddenTextMenuItem = 'true';
						hiddenTextWrapper.addClasses(['docxidian-file-menu-item', 'docxidian-find-hidden-text-menu-item']);
						const hiddenTextButton = hiddenTextWrapper.querySelector('button') ?? hiddenTextWrapper.createEl('button');
						hiddenTextButton.type = 'button';
						hiddenTextButton.addClass('docxidian-file-menu-button');
						retitleMenuButton(hiddenTextButton, 'Find hidden text...', labels.save);
					hiddenTextButton.removeAttribute('disabled');
					hiddenTextButton.removeAttribute('aria-disabled');
					hiddenTextButton.addEventListener('mousedown', (evt) => {
						evt.preventDefault();
					});
					hiddenTextButton.addEventListener('click', (evt) => {
						evt.preventDefault();
						evt.stopImmediatePropagation();
						evt.stopPropagation();
						activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
						void this.findHiddenText();
					});

					if (exportWrapper.parentElement === dropdown) {
						exportWrapper.after(hiddenTextWrapper);
					} else if (saveWrapper) {
						saveWrapper.after(hiddenTextWrapper);
					} else {
						dropdown.prepend(hiddenTextWrapper);
					}
				}
			});
		};

		addExportAsItem();
		this.fileMenuObserver = new MutationObserver(addExportAsItem);
		this.fileMenuObserver.observe(this.hostEl, {
			childList: true,
			subtree: true,
		});
		this.register(() => {
			this.fileMenuObserver?.disconnect();
			this.fileMenuObserver = null;
		});
		}

		private addEditorInsertMenuItems() {
			if (!this.hostEl) {
				return;
			}

			const addInsertImageItem = () => {
				if (!this.hostEl) {
					return;
				}

				const labels = getEditorMenuLabels(this.getEditorLocale());
				this.hostEl.querySelectorAll<HTMLElement>('.ep-root [role="menubar"] > div').forEach((menuItem) => {
					const menuButton = menuItem.querySelector(':scope > button');
					const menuLabel = normalizeMenuText(menuButton?.textContent ?? '');
					if (menuLabel !== labels.insert) {
						return;
					}

					const dropdown = Array.from(menuItem.children).find((child): child is HTMLElement => (
						isHTMLElement(child)
						&& child !== menuButton
						&& Boolean(child.querySelector(':scope > div > button, :scope > button'))
					));
					if (!dropdown || dropdown.querySelector('[data-docxidian-insert-image-menu-item]')) {
						return;
					}

					const itemWrappers = Array.from(dropdown.children).filter((child): child is HTMLElement => isHTMLElement(child));
					const sourceWrapper = itemWrappers.find((itemWrapper) => itemWrapper.querySelector(':scope > button'));
					const insertImageWrapper = sourceWrapper
						? sourceWrapper.cloneNode(true) as HTMLElement
						: activeDocument.createElement('div');
					insertImageWrapper.dataset.docxidianInsertImageMenuItem = 'true';
					insertImageWrapper.addClasses(['docxidian-file-menu-item', 'docxidian-insert-image-menu-item']);

					const insertImageButton = insertImageWrapper.querySelector('button') ?? insertImageWrapper.createEl('button');
					insertImageButton.type = 'button';
					insertImageButton.textContent = 'Insert image...';
					insertImageButton.addClass('docxidian-file-menu-button');
					insertImageButton.removeAttribute('disabled');
					insertImageButton.removeAttribute('aria-disabled');
					insertImageButton.addEventListener('mousedown', (evt) => {
						evt.preventDefault();
					});
					insertImageButton.addEventListener('click', (evt) => {
						evt.preventDefault();
						evt.stopImmediatePropagation();
						evt.stopPropagation();
						activeDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
						this.openImagePicker();
					});

					const imageLikeWrapper = itemWrappers.find((itemWrapper) => {
						const button = itemWrapper.querySelector(':scope > button');
						const text = normalizeMenuText(button?.textContent ?? '');
						return text.includes('image') || text.includes('picture');
					});
					if (imageLikeWrapper) {
						imageLikeWrapper.replaceWith(insertImageWrapper);
					} else {
						dropdown.appendChild(insertImageWrapper);
					}
				});
			};

			addInsertImageItem();
			this.insertMenuObserver = new MutationObserver(addInsertImageItem);
			this.insertMenuObserver.observe(this.hostEl, {
				childList: true,
				subtree: true,
			});
			this.register(() => {
				this.insertMenuObserver?.disconnect();
				this.insertMenuObserver = null;
			});
		}

		private normalizeNativeEditorMenuActionItems() {
			if (!this.hostEl) {
				return;
			}

			const normalizeMenuItems = () => {
				if (!this.hostEl) {
					return;
				}

				const labels = getEditorMenuLabels(this.getEditorLocale());
				const menuSpecs = [
					{ menuLabel: labels.file, optionLabels: labels.pageSetup },
					{ menuLabel: labels.format, optionLabels: labels.rightToLeft },
					{ menuLabel: labels.insert, optionLabels: labels.pageBreak },
					{ menuLabel: labels.insert, optionLabels: labels.tableOfContents },
				];

				this.hostEl.querySelectorAll<HTMLElement>('.ep-root [role="menubar"] > div').forEach((menuItem) => {
					const menuButton = menuItem.querySelector(':scope > button');
					const menuLabel = normalizeMenuText(menuButton?.textContent ?? '');
					const specs = menuSpecs.filter(spec => spec.menuLabel === menuLabel);
					if (specs.length === 0) {
						return;
					}

					const dropdown = Array.from(menuItem.children).find((child): child is HTMLElement => (
						isHTMLElement(child)
						&& child !== menuButton
						&& Boolean(child.querySelector(':scope > div > button, :scope > button'))
					));
					if (!dropdown) {
						return;
					}

					for (const button of Array.from(dropdown.querySelectorAll<HTMLButtonElement>('button'))) {
						const buttonLabel = normalizeMenuText(cleanEditorOptionLabel(button.textContent ?? ''));
						const matchesTarget = specs.some(spec => spec.optionLabels.some(label => (
							textStartsWithMenuLabel(buttonLabel, label)
							|| textStartsWithMenuLabel(normalizeMenuText(label), buttonLabel)
						)));
						if (!matchesTarget) {
							continue;
						}

						const wrapper = Array.from(dropdown.children).find((child): child is HTMLElement => (
							isHTMLElement(child) && child.contains(button)
						));
						wrapper?.addClasses(['docxidian-file-menu-item', 'docxidian-native-menu-action-item']);
						button.addClasses(['docxidian-file-menu-button', 'docxidian-native-menu-action-button']);
					}
				});
			};

			normalizeMenuItems();
			this.nativeMenuStyleObserver = new MutationObserver(normalizeMenuItems);
			this.nativeMenuStyleObserver.observe(this.hostEl, {
				childList: true,
				subtree: true,
			});
			this.register(() => {
				this.nativeMenuStyleObserver?.disconnect();
				this.nativeMenuStyleObserver = null;
			});
		}

		private trackEditorHoverState() {
		if (!this.hostEl) {
			return;
		}

			const markEditorHovering = () => activeDocument.body.classList.add('docxidian-editor-hovering');
			const clearEditorHovering = () => activeDocument.body.classList.remove('docxidian-editor-hovering');

		this.registerDomEvent(this.hostEl, 'pointerenter', markEditorHovering);
		this.registerDomEvent(this.hostEl, 'pointerleave', clearEditorHovering);
		this.register(clearEditorHovering);
	}

	private registerEditorSaveInterceptor() {
		this.registerDomEvent(activeDocument, 'click', (evt) => {
			if (
				!this.hostEl
				|| this.app.workspace.getActiveViewOfType(DocxView) !== this
				|| (evt.target instanceof Element && !!evt.target.closest('.modal'))
				|| !shouldHandleEditorSaveClick(evt.target, getEditorMenuLabels(this.getEditorLocale()).save)
			) {
				return;
			}

			evt.preventDefault();
			evt.stopImmediatePropagation();
			void this.saveCurrentDocument();
		}, true);
	}

	private registerEditorListAwareCopyInterceptor() {
		this.registerDomEvent(activeDocument, 'copy', (evt) => {
			if (!this.hostEl) {
				return;
			}

			const targetInsideHost = evt.target instanceof Node && this.hostEl.contains(evt.target);
			const activeInsideHost = activeDocument.activeElement instanceof Node && this.hostEl.contains(activeDocument.activeElement);
			const selection = activeDocument.getSelection();
			const selectionInsideHost = Boolean(
				selection
				&& !selection.isCollapsed
				&& (
					(selection.anchorNode && this.hostEl.contains(selection.anchorNode))
					|| (selection.focusNode && this.hostEl.contains(selection.focusNode))
				),
			);
			if (!targetInsideHost && !activeInsideHost && !selectionInsideHost) {
				return;
			}

			const handle = this.getReactHandle();
			if (handle) {
				window.setTimeout(() => {
					void handle.rewriteClipboardTextWithListMarkers();
				}, 0);
			}
		}, true);
	}

	private registerSaveShortcut() {
		this.registerDomEvent(activeDocument, 'keydown', (evt) => {
			if (
				!this.hostEl
				|| evt.key.toLowerCase() !== 's'
				|| (!evt.metaKey && !evt.ctrlKey)
				|| !(activeDocument.activeElement instanceof Node)
				|| !this.hostEl.contains(activeDocument.activeElement)
			) {
				return;
			}

			evt.preventDefault();
			evt.stopImmediatePropagation();
			void this.saveCurrentDocument();
		}, true);
	}

	private registerFindShortcut() {
		const handleFindShortcut = (evt: KeyboardEvent) => {
			if (!this.hostEl || !isPrimaryFindShortcut(evt) || !this.isActiveDocxView()) {
				return;
			}

			const target = evt.target instanceof Element ? evt.target : null;
			if (target?.closest('.modal') && !target.closest('.docxidian-find-dialog')) {
				return;
			}

			evt.preventDefault();
			evt.stopImmediatePropagation();
			this.openFindDialog();
		};

		this.registerDomEvent(window, 'keydown', handleFindShortcut, true);
		this.registerDomEvent(activeDocument, 'keydown', handleFindShortcut, true);
	}

	private isActiveDocxView() {
		if (this.app.workspace.getActiveViewOfType(DocxView) === this) {
			return true;
		}

		const workspace = this.app.workspace as typeof this.app.workspace & { activeLeaf?: WorkspaceLeaf | null };
		if (workspace.activeLeaf === this.leaf) {
			return true;
		}

		if (this.contentEl.closest('.workspace-leaf.mod-active')) {
			return true;
		}

			const activeElement = activeDocument.activeElement;
		return Boolean(activeElement instanceof Node && this.hostEl?.contains(activeElement));
	}

	private registerEditorDropdownScrollGuard() {
		const keepEditorListboxOpen = (evt: Event) => {
			if (!this.hostEl || this.app.workspace.getActiveViewOfType(DocxView) !== this || !(evt.target instanceof Element)) {
				return;
			}

			const listbox = evt.target.closest('[role="listbox"]');
			if (listbox && this.hostEl.contains(listbox)) {
				evt.stopImmediatePropagation();
				evt.stopPropagation();
			}
		};

		this.registerDomEvent(window, 'scroll', keepEditorListboxOpen, true);
	}

	private getReactHandle(): DocxReactViewHandle | null {
		return this.reactMount?.getHandle() ?? null;
	}

	private getReactProps(): DocxReactViewProps {
		return {
			file: this.file,
			buffer: this.buffer,
			documentKey: this.file ? `${this.file.path}:${this.documentSession}` : 'docxidian-empty',
			error: this.error,
			isLoading: this.isLoading,
			authorName: this.getAuthorName(),
			i18n: this.getEditorLocale(),
			showRuler: this.getShowRuler(),
			autosave: this.getAutosave(),
			defaultZoom: this.getDefaultZoom(),
			reserveReviewSidebar: this.reserveReviewSidebar,
			onDirtyChange: (isDirty) => {
				this.isDirty = isDirty;
			},
			onSave: (buffer) => this.saveFile(buffer),
			onDocumentNameChange: (name) => this.renameFile(name),
		};
	}

	private async ensureReactMount() {
		if (this.reactMount || this.reactMountLoading || !this.hostEl) {
			debugLog('editor', 'Skipping React mount request', {
				hasMount: Boolean(this.reactMount),
				isLoading: this.reactMountLoading,
				hasHost: Boolean(this.hostEl),
			});
			return;
		}

		this.reactMountLoading = true;
		try {
			infoLog('editor', 'Loading DOCX editor UI');
			const { createDocxReactMount } = await loadDocxEditorChunk();
			if (!this.hostEl) {
				debugLog('editor', 'Aborted DOCX editor mount because host was removed');
				return;
			}

			this.hostEl.empty();
			this.reactMount = createDocxReactMount(this.hostEl);
			this.reactMount.render(this.getReactProps());
			infoLog('editor', 'Mounted DOCX editor UI', { file: this.file?.path });
		} catch (loadError) {
			const message = loadError instanceof Error ? loadError.message : 'Unknown load error';
			this.error = `Could not load DOCX editor: ${message}`;
			errorLog('editor', this.error, loadError);
			new Notice(this.error);
			if (this.hostEl) {
				this.hostEl.empty();
				this.hostEl.createDiv({ cls: 'docxidian-editor-load-error', text: this.error });
			}
		} finally {
			this.reactMountLoading = false;
		}
	}

	private render() {
		if (!this.hostEl) {
			return;
		}

		if (!this.reactMount) {
			this.hostEl.empty();
			this.hostEl.createDiv({ cls: 'docxidian-editor-loading', text: 'Loading DOCX editor...' });
			void this.ensureReactMount();
			return;
		}

		this.reactMount.render(this.getReactProps());
	}
}
