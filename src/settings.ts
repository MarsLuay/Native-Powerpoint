import { App, Notice, PluginSettingTab, Setting } from 'obsidian';
import type DocxidianPlugin from './main';
import { configureDocxidianLogger, infoLog } from './logger';
import { DOCXIDIAN_LANGUAGE_OPTIONS, DEFAULT_LANGUAGE, normalizeDocxidianLanguage, type DocxidianLanguage } from './locales';

const DEFAULT_ZOOM = 1;
const MIN_DEFAULT_ZOOM = 0.5;
const MAX_DEFAULT_ZOOM = 2;
const DEFAULT_ZOOM_STEP = 0.05;

const FREE_TEMPLATE_LINKS = [
	{ name: 'Microsoft Create', url: 'https://create.microsoft.com/en-us/templates/presentations' },
	{ name: 'Slidesgo', url: 'https://slidesgo.com/' },
	{ name: 'SlidesCarnival', url: 'https://www.slidescarnival.com/' },
	{ name: 'SlidesMania', url: 'https://slidesmania.com/free-templates/presentation-templates/' },
	{ name: 'Canva templates', url: 'https://www.canva.com/presentations/templates/slides/' },
];

export interface NativePowerPointSettings {
	autosaveEnabled: boolean;
	hideUnsupportedSvgContent: boolean;
	openWithYoloMode: boolean;
	showInspector: boolean;
	setOpenWithYoloMode: (value: boolean) => Promise<void>;
}

export interface DocxidianSettings {
	authorName: string;
	editorLanguage: DocxidianLanguage;
	showRuler: boolean;
	autosave: boolean;
	createBackupsBeforeSave: boolean;
	defaultZoom: number;
	enableDocxSearchIndex: boolean;
	autoIndexDocxSearch: boolean;
	debugLogging: boolean;
	powerPointAutosaveEnabled: boolean;
	powerPointHideUnsupportedSvgContent: boolean;
	powerPointOpenWithYoloMode: boolean;
	powerPointShowInspector: boolean;
	disableDocxFiles: boolean;
	disablePowerPointFiles: boolean;
}

export const DEFAULT_SETTINGS: DocxidianSettings = {
	authorName: 'Mars',
	editorLanguage: DEFAULT_LANGUAGE,
	showRuler: false,
	autosave: true,
	createBackupsBeforeSave: false,
	defaultZoom: DEFAULT_ZOOM,
	enableDocxSearchIndex: true,
	autoIndexDocxSearch: true,
	debugLogging: false,
	powerPointAutosaveEnabled: true,
	powerPointHideUnsupportedSvgContent: false,
	powerPointOpenWithYoloMode: false,
	powerPointShowInspector: false,
	disableDocxFiles: false,
	disablePowerPointFiles: false,
};

export function getNativePowerPointSettings(
	settings: DocxidianSettings,
	setOpenWithYoloMode: (value: boolean) => Promise<void> = async () => {}
): NativePowerPointSettings {
	return {
		autosaveEnabled: settings.powerPointAutosaveEnabled,
		hideUnsupportedSvgContent: settings.powerPointHideUnsupportedSvgContent,
		openWithYoloMode: settings.powerPointOpenWithYoloMode,
		showInspector: settings.powerPointShowInspector,
		setOpenWithYoloMode,
	};
}

export function normalizeDefaultZoom(value: unknown): number {
	const numericValue = typeof value === 'number' ? value : Number(value);
	if (!Number.isFinite(numericValue)) {
		return DEFAULT_ZOOM;
	}

	const clampedValue = Math.min(MAX_DEFAULT_ZOOM, Math.max(MIN_DEFAULT_ZOOM, numericValue));
	return Math.round(clampedValue / DEFAULT_ZOOM_STEP) * DEFAULT_ZOOM_STEP;
}

function formatZoom(value: number): string {
	return `${Math.round(value * 100)}%`;
}

export class DocxidianSettingTab extends PluginSettingTab {
	plugin: DocxidianPlugin;

	constructor(app: App, plugin: DocxidianPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		this.renderSettings();
	}

	private renderSettings(): void {
		const { containerEl } = this;
		const selectedLanguage = normalizeDocxidianLanguage(this.plugin.settings.editorLanguage);
		const selectedZoom = normalizeDefaultZoom(this.plugin.settings.defaultZoom);

		this.plugin.settings.editorLanguage = selectedLanguage;
		this.plugin.settings.defaultZoom = selectedZoom;

		containerEl.empty();

		new Setting(containerEl)
			.setName('Identity')
			.setHeading();

		new Setting(containerEl)
			.setName('Author name')
			.setDesc('Used for comments and tracked changes.')
			.addText(text => text
				.setPlaceholder('Mars')
				.setValue(this.plugin.settings.authorName)
				.onChange(async (value) => {
					this.plugin.settings.authorName = value.trim() || DEFAULT_SETTINGS.authorName;
					await this.plugin.saveSettings();
				}))
			.addButton(button => button
				.setButtonText('Reset')
				.onClick(async () => {
					this.plugin.settings.authorName = DEFAULT_SETTINGS.authorName;
					await this.plugin.saveSettings();
					this.renderSettings();
				}));

		new Setting(containerEl)
			.setName('File type handoff')
			.setHeading();

		new Setting(containerEl)
			.setName('Turn off for DOCX files')
			.setDesc('Turns off plugin specifically for DOCX files in favor of another plugin </3')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.disableDocxFiles)
				.onChange(async (value) => {
					this.plugin.settings.disableDocxFiles = value;
					await this.plugin.saveSettings();
					new Notice('Reload Obsidian or disable/re-enable this plugin to update DOCX file handling.');
				}));

		new Setting(containerEl)
			.setName('Turn off for PPTX files')
			.setDesc('Turns off plugin specifically for PPTX files in favor of another plugin </3')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.disablePowerPointFiles)
				.onChange(async (value) => {
					this.plugin.settings.disablePowerPointFiles = value;
					await this.plugin.saveSettings();
					new Notice('Reload Obsidian or disable/re-enable this plugin to update PPTX file handling.');
				}));

		new Setting(containerEl)
			.setName('Editor defaults')
			.setHeading();

		new Setting(containerEl)
			.setName('Default language')
			.setDesc('English is the default language for the editor toolbar, dialogs, and messages.')
			.addDropdown(dropdown => {
				for (const option of DOCXIDIAN_LANGUAGE_OPTIONS) {
					const label = option.code === DEFAULT_LANGUAGE ? `${option.label} (default)` : option.label;
					dropdown.addOption(option.code, label);
				}

				dropdown
					.setValue(selectedLanguage)
					.onChange(async (value) => {
						this.plugin.settings.editorLanguage = normalizeDocxidianLanguage(value);
						await this.plugin.saveSettings();
						this.plugin.refreshDocxViews();
					});
			})
			.addButton(button => button
				.setButtonText('Use English')
				.onClick(async () => {
					this.plugin.settings.editorLanguage = DEFAULT_LANGUAGE;
					await this.plugin.saveSettings();
					this.plugin.refreshDocxViews();
					this.renderSettings();
				}));

		new Setting(containerEl)
			.setName('Ruler')
			.setDesc('Show the page ruler above the document body by default.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.showRuler)
				.onChange(async (value) => {
					this.plugin.settings.showRuler = value;
					await this.plugin.saveSettings();
					this.plugin.refreshDocxViews();
				}));

		const zoomSetting = new Setting(containerEl)
			.setName('Default zoom')
			.setDesc('Initial zoom for DOCX files when they open.');
		const zoomValueEl = zoomSetting.controlEl.createSpan({
			cls: 'docxidian-setting-value',
			text: formatZoom(selectedZoom),
		});

		zoomSetting
			.addSlider(slider => slider
				.setLimits(MIN_DEFAULT_ZOOM, MAX_DEFAULT_ZOOM, DEFAULT_ZOOM_STEP)
				.setValue(selectedZoom)
				.onChange(async (value) => {
					const zoom = normalizeDefaultZoom(value);
					this.plugin.settings.defaultZoom = zoom;
					zoomValueEl.setText(formatZoom(zoom));
					await this.plugin.saveSettings();
				}))
			.addButton(button => button
				.setButtonText('Reset')
				.onClick(async () => {
					this.plugin.settings.defaultZoom = DEFAULT_SETTINGS.defaultZoom;
					await this.plugin.saveSettings();
					this.renderSettings();
				}));

		new Setting(containerEl)
			.setName('Saving')
			.setHeading();

		new Setting(containerEl)
			.setName('Autosave')
			.setDesc('Automatically save the document shortly after changes.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.autosave)
				.onChange(async (value) => {
					this.plugin.settings.autosave = value;
					await this.plugin.saveSettings();
					this.plugin.refreshDocxViews();
				}));

		new Setting(containerEl)
			.setName('Backups')
			.setDesc('Create one timestamped backup before the first overwrite in each open DOCX session.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.createBackupsBeforeSave)
				.onChange(async (value) => {
					this.plugin.settings.createBackupsBeforeSave = value;
					await this.plugin.saveSettings();
					this.plugin.refreshDocxViews();
				}));

		new Setting(containerEl)
			.setName('PowerPoint')
			.setHeading();

		new Setting(containerEl)
			.setName('Autosave edits')
			.setDesc('Save editable PowerPoint files 1500 ms after edits settle. When disabled, closing or switching files writes unsaved edits to a recovery copy without overwriting the original.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.powerPointAutosaveEnabled)
				.onChange(async (value) => {
					this.plugin.settings.powerPointAutosaveEnabled = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Show inspector panel')
			.setDesc('Show the object inspector panel on the right side of the PowerPoint editor. Off by default.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.powerPointShowInspector)
				.onChange(async (value) => {
					this.plugin.settings.powerPointShowInspector = value;
					await this.plugin.saveSettings();
					this.plugin.refreshPowerPointViews();
				}));

		new Setting(containerEl)
			.setName('Hide unsupported SVG details')
			.setDesc('Temporarily hide unsupported PowerPoint SVG details in the Obsidian preview. Off by default. This does not delete them from the PPTX file.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.powerPointHideUnsupportedSvgContent)
				.onChange(async (value) => {
					this.plugin.settings.powerPointHideUnsupportedSvgContent = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Open PowerPoints with YOLO mode')
			.setDesc('Remember YOLO mode for future PPTX files so the original slide SVG is shown without preview cleanup.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.powerPointOpenWithYoloMode)
				.onChange(async (value) => {
					this.plugin.settings.powerPointOpenWithYoloMode = value;
					await this.plugin.saveSettings();
				}));

		const templateBox = containerEl.createDiv({ cls: 'native-powerpoint-template-box' });
		templateBox.createDiv({ cls: 'native-powerpoint-template-title', text: 'Free presentation templates' });
		templateBox.createDiv({
			cls: 'native-powerpoint-template-desc',
			text: 'A few places to find starter decks you can download or customize.',
		});

		const templateLinks = templateBox.createDiv({ cls: 'native-powerpoint-template-links' });
		for (const link of FREE_TEMPLATE_LINKS) {
			templateLinks.createEl('a', {
				cls: 'native-powerpoint-template-link',
				text: link.name,
				attr: {
					href: link.url,
					rel: 'noopener noreferrer',
					target: '_blank',
				},
			});
		}

		new Setting(containerEl)
			.setName('Search')
			.setHeading();

		new Setting(containerEl)
			.setName('DOCX search index')
			.setDesc('Extract text from DOCX files into a local cache so the vault-wide DOCX search command can find them.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.enableDocxSearchIndex)
				.onChange(async (value) => {
					this.plugin.settings.enableDocxSearchIndex = value;
					await this.plugin.saveSettings();
					if (value) {
						await this.plugin.rebuildDocxSearchIndex(false);
					}
				}));

		new Setting(containerEl)
			.setName('Auto-index DOCX changes')
			.setDesc('Keep the DOCX search cache updated when DOCX files are created, edited, renamed, or deleted.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.autoIndexDocxSearch)
				.onChange(async (value) => {
					this.plugin.settings.autoIndexDocxSearch = value;
					await this.plugin.saveSettings();
					if (value && this.plugin.settings.enableDocxSearchIndex) {
						await this.plugin.rebuildDocxSearchIndex(false);
					}
				}));

		new Setting(containerEl)
			.setName('Diagnostics')
			.setHeading();

		new Setting(containerEl)
			.setName('Debug logging')
			.setDesc('Log DOCX and PowerPoint diagnostics to the developer console and the in-memory Native PowerPoint Doc Editor log.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.debugLogging)
				.onChange(async (value) => {
					this.plugin.settings.debugLogging = value;
					configureDocxidianLogger(value);
					infoLog('settings', `Debug logging ${value ? 'enabled' : 'disabled'}`);
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Copy full debug log')
			.setDesc('Copy all Native PowerPoint Doc Editor logs produced for DOCX and PowerPoint to the clipboard.')
			.addButton(button => button
				.setButtonText('Copy log')
				.onClick(async () => {
					await this.plugin.copyDebugLog('all');
				}));

		const reportBugBox = containerEl.createDiv({ cls: 'native-powerpoint-report-bug' });
		reportBugBox.createEl('a', {
			cls: 'native-powerpoint-report-bug-link',
			text: 'Report bug',
			attr: {
				href: 'https://github.com/MarsLuay/NativePowerPointDocEditor/issues',
				rel: 'noopener noreferrer',
				target: '_blank',
			},
		});
	}
}
