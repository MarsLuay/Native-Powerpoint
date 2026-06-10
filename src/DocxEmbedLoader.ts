import { App, Component, MarkdownPostProcessorContext, MarkdownRenderChild, Plugin, TFile } from 'obsidian';
import type { Translations } from '@eigenpal/docx-editor-i18n';
import { loadDocxEditorChunk } from './docxEditorLoader';
import { debugLog, errorLog, infoLog, warnLog } from './logger';

const DOCX_EMBED_SELECTOR = '.internal-embed[src], .internal-embed[data-src]';

interface EmbedInfo {
	containerEl: HTMLElement;
}

type DocxFileEmbedCreator = (info: EmbedInfo, file: TFile, subpath: string) => Component;

interface EmbedRegistry {
	registerExtension?: (extension: string, creator: DocxFileEmbedCreator) => void;
	registerExtensions?: (extensions: string[], creator: DocxFileEmbedCreator) => void;
	unregisterExtension?: (extension: string) => void;
	unregisterExtensions?: (extensions: string[]) => void;
}

function getEmbedRegistry(app: App) {
	return (app as App & { embedRegistry?: EmbedRegistry }).embedRegistry;
}

function getEmbedLinkPath(embedEl: Element) {
	return embedEl.getAttribute('src') ?? embedEl.getAttribute('data-src') ?? '';
}

function stripSubpath(linkPath: string) {
	return linkPath.split('#')[0] ?? '';
}

function isDocxLink(linkPath: string) {
	return stripSubpath(linkPath).toLowerCase().endsWith('.docx');
}

function resolveDocxEmbed(app: App, linkPath: string, sourcePath: string) {
	const file = app.metadataCache.getFirstLinkpathDest(stripSubpath(linkPath), sourcePath);
	return file instanceof TFile && file.extension.toLowerCase() === 'docx' ? file : null;
}

function collectEmbedElements(el: HTMLElement) {
	const embeds = Array.from(el.querySelectorAll(DOCX_EMBED_SELECTOR));
	if (el.matches(DOCX_EMBED_SELECTOR)) {
		embeds.unshift(el);
	}

	return embeds;
}

function hasRenderableDocxEmbed(app: App, el: HTMLElement, ctx: MarkdownPostProcessorContext) {
	return collectEmbedElements(el).some((embedEl) => {
		if (!(embedEl instanceof HTMLElement) || embedEl.dataset.docxidianEmbed === 'true') {
			return false;
		}

		const linkPath = getEmbedLinkPath(embedEl);
		return isDocxLink(linkPath) && !!resolveDocxEmbed(app, linkPath, ctx.sourcePath);
	});
}

class LazyDocxFileEmbed extends Component {
	private unloaded = false;

	constructor(
		private info: EmbedInfo,
		private app: App,
		private file: TFile,
		private getEditorLocale: () => Translations | undefined,
		private subpath = '',
	) {
		super();
		this.info.containerEl.addClasses(['docxidian-embed', 'docxidian-native-embed']);
		this.registerDomEvent(this.info.containerEl, 'click', (evt) => {
			evt.stopImmediatePropagation();
		});
	}

	onload() {
		super.onload();
		void this.loadEmbed();
	}

	onunload() {
		this.unloaded = true;
		super.onunload();
	}

	private async loadEmbed() {
		const { containerEl } = this.info;
		containerEl.empty();
		containerEl.createDiv({ cls: 'docxidian-embed-loading', text: `Loading ${this.file.name}...` });
		debugLog('embed', `Loading DOCX embed ${this.file.path}`);

		try {
			const { DocxFileEmbed } = await loadDocxEditorChunk();
			if (this.unloaded) {
				debugLog('embed', `Discarded loaded embed because it was unloaded: ${this.file.path}`);
				return;
			}

			containerEl.empty();
			this.addChild(new DocxFileEmbed(this.info, this.app, this.file, this.getEditorLocale, this.subpath));
			infoLog('embed', `Loaded DOCX embed ${this.file.path}`);
		} catch (error) {
			if (this.unloaded) {
				return;
			}

			const message = error instanceof Error ? error.message : 'Unknown error';
			errorLog('embed', `Could not load DOCX embed ${this.file.path}`, error);
			containerEl.empty();
			containerEl.createDiv({
				cls: 'docxidian-embed-error',
				text: `Could not load the DOCX editor: ${message}`,
			});
		}
	}
}

class DocxEmbedScanChild extends MarkdownRenderChild {
	private scanTimeout: number | null = null;
	private observer: MutationObserver | null = null;

	constructor(
		containerEl: HTMLElement,
		private app: App,
		private ctx: MarkdownPostProcessorContext,
		private getEditorLocale: () => Translations | undefined,
	) {
		super(containerEl);
	}

	onload() {
		this.scan();
		this.queueScan(0);
		this.queueScan(100);
		this.observer = new MutationObserver(() => this.queueScan(25));
		this.observer.observe(this.containerEl, {
			attributes: true,
			attributeFilter: ['data-src', 'src'],
			childList: true,
			subtree: true,
		});
	}

	onunload() {
		if (this.scanTimeout !== null) {
			window.clearTimeout(this.scanTimeout);
			this.scanTimeout = null;
		}
		this.observer?.disconnect();
		this.observer = null;
		super.onunload();
	}

	private queueScan(delay: number) {
		if (this.scanTimeout !== null) {
			return;
		}

		this.scanTimeout = window.setTimeout(() => {
			this.scanTimeout = null;
			this.scan();
		}, delay);
	}

	private scan() {
		if (!hasRenderableDocxEmbed(this.app, this.containerEl, this.ctx)) {
			return;
		}

		void loadDocxEditorChunk().then(({ renderDocxEmbeds }) => {
			debugLog('embed', `Rendering DOCX embeds in ${this.ctx.sourcePath}`);
			renderDocxEmbeds(this.app, this.containerEl, this.ctx, this.getEditorLocale);
		}).catch((error) => {
			errorLog('embed', `Could not scan DOCX embeds in ${this.ctx.sourcePath}`, error);
		});
	}
}

export function processDocxEmbeds(
	app: App,
	el: HTMLElement,
	ctx: MarkdownPostProcessorContext,
	getEditorLocale: () => Translations | undefined,
) {
	ctx.addChild(new DocxEmbedScanChild(el, app, ctx, getEditorLocale));
}

export function registerDocxFileEmbed(plugin: Plugin, getEditorLocale: () => Translations | undefined) {
	const registry = getEmbedRegistry(plugin.app);
	if (!registry) {
		warnLog('embed', 'Obsidian embed registry is unavailable');
		return false;
	}

	const createEmbed: DocxFileEmbedCreator = (info, file, subpath) => new LazyDocxFileEmbed(info, plugin.app, file, getEditorLocale, subpath);

	try {
		if (typeof registry.registerExtension === 'function') {
			registry.registerExtension('docx', createEmbed);
			plugin.register(() => registry.unregisterExtension?.('docx'));
			infoLog('embed', 'Registered DOCX file embed extension');
			return true;
		}

		if (typeof registry.registerExtensions === 'function') {
			registry.registerExtensions(['docx'], createEmbed);
			plugin.register(() => registry.unregisterExtensions?.(['docx']));
			infoLog('embed', 'Registered DOCX file embed extensions');
			return true;
		}
	} catch (error) {
		errorLog('embed', 'Could not register DOCX file embed extension', error);
		return false;
	}

	warnLog('embed', 'Obsidian embed registry does not expose a DOCX registration method');
	return false;
}
