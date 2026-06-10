import { App, Modal, Notice, TFile } from 'obsidian';
import { DocxSearchIndex, type DocxSearchIndexStats, type DocxSearchResult } from './docxSearchIndex';
import { errorLog } from './logger';

const SEARCH_DEBOUNCE_MS = 150;

function formatIndexStatus(stats: DocxSearchIndexStats): string {
	const indexedPart = stats.indexed > 0 ? `${stats.indexed} updated` : 'up to date';
	const removedPart = stats.removed > 0 ? `, ${stats.removed} removed` : '';
	const errorPart = stats.errors > 0 ? `, ${stats.errors} failed` : '';

	return `Indexed ${stats.total} DOCX files (${indexedPart}${removedPart}${errorPart}).`;
}

export class DocxSearchModal extends Modal {
	private inputEl: HTMLInputElement | null = null;
	private statusEl: HTMLElement | null = null;
	private resultsEl: HTMLElement | null = null;
	private debounceTimer: number | null = null;
	private readyPromise: Promise<void> | null = null;

	constructor(
		app: App,
		private searchIndex: DocxSearchIndex,
	) {
		super(app);
	}

	onOpen(): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.addClass('docxidian-search-modal');
		contentEl.createEl('h2', { text: 'Search DOCX files' });

		const toolbarEl = contentEl.createDiv({ cls: 'docxidian-search-toolbar' });
		this.inputEl = toolbarEl.createEl('input', {
			cls: 'docxidian-search-input',
			type: 'search',
		});
		this.inputEl.placeholder = 'Search indexed DOCX text...';
		this.inputEl.setAttribute('spellcheck', 'false');

		const rebuildButton = toolbarEl.createEl('button', { text: 'Rebuild' });
		rebuildButton.addEventListener('click', () => {
			void this.rebuildIndex(true);
		});

		this.statusEl = contentEl.createDiv({
			cls: 'docxidian-search-status',
			text: 'Preparing DOCX search index...',
		});
		this.resultsEl = contentEl.createDiv({ cls: 'docxidian-search-results' });

		this.inputEl.addEventListener('input', () => this.queueSearch());
		this.readyPromise = this.rebuildIndex(false);

		window.setTimeout(() => this.inputEl?.focus());
	}

	onClose(): void {
		if (this.debounceTimer !== null) {
			window.clearTimeout(this.debounceTimer);
			this.debounceTimer = null;
		}

		this.contentEl.empty();
	}

	private async rebuildIndex(force: boolean): Promise<void> {
		this.setStatus(force ? 'Rebuilding DOCX search index...' : 'Updating DOCX search index...');

		try {
			const stats = await this.searchIndex.rebuild({ force });
			this.setStatus(formatIndexStatus(stats));
			this.renderSearch();
		} catch (error) {
			errorLog('search', 'Could not rebuild DOCX search index', error);
			this.setStatus('Could not rebuild DOCX search index.');
			new Notice('Could not rebuild DOCX search index. Check the Native PowerPoint Doc Editor debug log.');
		}
	}

	private queueSearch(): void {
		if (this.debounceTimer !== null) {
			window.clearTimeout(this.debounceTimer);
		}

		this.debounceTimer = window.setTimeout(() => {
			this.debounceTimer = null;
			void this.readyPromise?.then(() => this.renderSearch());
		}, SEARCH_DEBOUNCE_MS);
	}

	private renderSearch(): void {
		const query = this.inputEl?.value.trim() ?? '';
		if (!this.resultsEl) {
			return;
		}

		this.resultsEl.empty();

		if (!query) {
			const stats = this.searchIndex.getStats();
			this.resultsEl.createDiv({
				cls: 'docxidian-search-empty',
				text: stats.files > 0 ? 'Type to search DOCX files.' : 'No DOCX files are indexed yet.',
			});
			return;
		}

		const results = this.searchIndex.search(query);
		if (results.length === 0) {
			this.resultsEl.createDiv({ cls: 'docxidian-search-empty', text: 'No DOCX matches found.' });
			return;
		}

		for (const result of results) {
			this.renderResult(result);
		}
	}

	private renderResult(result: DocxSearchResult): void {
		if (!this.resultsEl) {
			return;
		}

		const resultEl = this.resultsEl.createDiv({ cls: 'docxidian-search-result' });
		resultEl.setAttribute('role', 'button');
		resultEl.setAttribute('tabindex', '0');

		const titleRow = resultEl.createDiv({ cls: 'docxidian-search-result-title-row' });
		titleRow.createSpan({ cls: 'docxidian-search-result-title', text: result.name });
		titleRow.createSpan({
			cls: 'docxidian-search-result-count',
			text: `${result.matchCount} ${result.matchCount === 1 ? 'match' : 'matches'}`,
		});

		resultEl.createDiv({ cls: 'docxidian-search-result-path', text: result.path });

		for (const snippet of result.snippets) {
			resultEl.createDiv({ cls: 'docxidian-search-snippet', text: snippet });
		}

		resultEl.addEventListener('click', () => {
			void this.openResult(result.path);
		});
		resultEl.addEventListener('keydown', (evt) => {
			if (evt.key !== 'Enter' && evt.key !== ' ') {
				return;
			}

			evt.preventDefault();
			void this.openResult(result.path);
		});
	}

	private async openResult(path: string): Promise<void> {
		const file = this.app.vault.getAbstractFileByPath(path);
		if (!(file instanceof TFile)) {
			new Notice('That DOCX file is no longer in the vault.');
			await this.searchIndex.removePath(path);
			this.renderSearch();
			return;
		}

		await this.app.workspace.getLeaf(false).openFile(file);
		this.close();
	}

	private setStatus(message: string): void {
		this.statusEl?.setText(message);
	}
}
