import { App, TFile, normalizePath } from 'obsidian';
import { extractDocxText } from './docxTextExtractor';
import { debugLog, errorLog, infoLog, warnLog } from './logger';

const INDEX_FILE_NAME = 'docx-search-index.json';
const SEARCH_SNIPPET_RADIUS = 90;

export interface DocxSearchIndexStats {
	total: number;
	indexed: number;
	skipped: number;
	removed: number;
	errors: number;
}

interface DocxSearchCacheFile {
	version: 1;
	generatedAt: string;
	entries: Record<string, DocxSearchCacheEntry>;
}

interface DocxSearchCacheEntry {
	path: string;
	mtime: number;
	size: number;
	text: string;
	indexedAt: string;
	error?: string;
}

export interface DocxSearchResult {
	path: string;
	name: string;
	snippets: string[];
	matchCount: number;
}

function createEmptyStats(): DocxSearchIndexStats {
	return {
		total: 0,
		indexed: 0,
		skipped: 0,
		removed: 0,
		errors: 0,
	};
}

function isDocxPath(path: string): boolean {
	return path.toLowerCase().endsWith('.docx');
}

function createSnippet(text: string, lowerText: string, lowerQuery: string, index: number): string {
	const start = Math.max(0, index - SEARCH_SNIPPET_RADIUS);
	const end = Math.min(text.length, index + lowerQuery.length + SEARCH_SNIPPET_RADIUS);
	const prefix = start > 0 ? '...' : '';
	const suffix = end < text.length ? '...' : '';
	const snippet = text.slice(start, end).replace(/\s+/g, ' ').trim();

	return `${prefix}${snippet}${suffix}`;
}

function countMatches(lowerText: string, lowerQuery: string): number {
	let count = 0;
	let index = lowerText.indexOf(lowerQuery);

	while (index !== -1) {
		count += 1;
		index = lowerText.indexOf(lowerQuery, index + lowerQuery.length);
	}

	return count;
}

function getResultName(path: string): string {
	const parts = path.split('/');
	return parts[parts.length - 1] ?? path;
}

export class DocxSearchIndex {
	private entries = new Map<string, DocxSearchCacheEntry>();
	private loaded = false;
	private writePromise: Promise<void> = Promise.resolve();

	constructor(
		private app: App,
		private pluginDir: string | undefined,
	) {}

	async load(): Promise<void> {
		if (this.loaded) {
			return;
		}

		const indexPath = this.getIndexPath();
		this.loaded = true;

		try {
			const exists = await this.app.vault.adapter.exists(indexPath);
			if (!exists) {
				return;
			}

			const rawIndex = await this.app.vault.adapter.read(indexPath);
			const parsedIndex = JSON.parse(rawIndex) as Partial<DocxSearchCacheFile>;
			const entries = parsedIndex.entries ?? {};

			for (const [path, entry] of Object.entries(entries)) {
				if (!isDocxPath(path) || typeof entry.text !== 'string') {
					continue;
				}

				this.entries.set(path, {
					path,
					mtime: Number(entry.mtime) || 0,
					size: Number(entry.size) || 0,
					text: entry.text,
					indexedAt: typeof entry.indexedAt === 'string' ? entry.indexedAt : '',
					error: typeof entry.error === 'string' ? entry.error : undefined,
				});
			}

			infoLog('search', 'Loaded DOCX search index', {
				path: indexPath,
				entries: this.entries.size,
			});
		} catch (error) {
			this.entries.clear();
			warnLog('search', 'Could not load DOCX search index cache', error);
		}
	}

	async save(): Promise<void> {
		await this.load();

		const entries: Record<string, DocxSearchCacheEntry> = {};
		for (const [path, entry] of this.entries) {
			entries[path] = entry;
		}

		const indexFile: DocxSearchCacheFile = {
			version: 1,
			generatedAt: new Date().toISOString(),
			entries,
		};

		this.writePromise = this.writePromise
			.catch(() => undefined)
			.then(() => this.app.vault.adapter.write(this.getIndexPath(), JSON.stringify(indexFile, null, 2)));

		await this.writePromise;
	}

	async rebuild(options: { force?: boolean } = {}): Promise<DocxSearchIndexStats> {
		await this.load();

		const stats = createEmptyStats();
		const docxFiles = this.app.vault.getFiles().filter(file => this.isDocxFile(file));
		const livePaths = new Set(docxFiles.map(file => file.path));

		stats.total = docxFiles.length;

		for (const path of Array.from(this.entries.keys())) {
			if (!livePaths.has(path)) {
				this.entries.delete(path);
				stats.removed += 1;
			}
		}

		for (const file of docxFiles) {
			const result = await this.indexFile(file, { force: options.force === true, save: false });
			stats[result] += 1;
		}

		await this.save();
		infoLog('search', 'Rebuilt DOCX search index', stats);
		return stats;
	}

	async indexFile(file: TFile, options: { force?: boolean; save?: boolean } = {}): Promise<'indexed' | 'skipped' | 'errors'> {
		await this.load();

		if (!this.isDocxFile(file)) {
			return 'skipped';
		}

		const existingEntry = this.entries.get(file.path);
		if (
			options.force !== true
			&& existingEntry
			&& existingEntry.mtime === file.stat.mtime
			&& existingEntry.size === file.stat.size
		) {
			return 'skipped';
		}

		try {
			const buffer = await this.app.vault.readBinary(file);
			const text = await extractDocxText(buffer);

			this.entries.set(file.path, {
				path: file.path,
				mtime: file.stat.mtime,
				size: file.stat.size,
				text,
				indexedAt: new Date().toISOString(),
			});

			debugLog('search', `Indexed ${file.path}`, {
				characters: text.length,
			});

			if (options.save !== false) {
				await this.save();
			}

			return 'indexed';
		} catch (error) {
			const message = error instanceof Error ? error.message : String(error);

			this.entries.set(file.path, {
				path: file.path,
				mtime: file.stat.mtime,
				size: file.stat.size,
				text: '',
				indexedAt: new Date().toISOString(),
				error: message,
			});
			errorLog('search', `Could not index ${file.path}`, error);

			if (options.save !== false) {
				await this.save();
			}

			return 'errors';
		}
	}

	async removePath(path: string): Promise<void> {
		await this.load();

		if (this.entries.delete(path)) {
			await this.save();
			debugLog('search', `Removed ${path} from DOCX search index`);
		}
	}

	search(query: string, limit = 50): DocxSearchResult[] {
		const normalizedQuery = query.trim().toLowerCase();
		if (!normalizedQuery) {
			return [];
		}

		const results: DocxSearchResult[] = [];

		for (const entry of this.entries.values()) {
			if (!entry.text || entry.error) {
				continue;
			}

			const lowerText = entry.text.toLowerCase();
			const firstMatchIndex = lowerText.indexOf(normalizedQuery);
			if (firstMatchIndex === -1) {
				continue;
			}

			const snippets = [createSnippet(entry.text, lowerText, normalizedQuery, firstMatchIndex)];
			const secondMatchIndex = lowerText.indexOf(normalizedQuery, firstMatchIndex + normalizedQuery.length);
			if (secondMatchIndex !== -1) {
				snippets.push(createSnippet(entry.text, lowerText, normalizedQuery, secondMatchIndex));
			}

			results.push({
				path: entry.path,
				name: getResultName(entry.path),
				snippets,
				matchCount: countMatches(lowerText, normalizedQuery),
			});
		}

		return results
			.sort((left, right) => right.matchCount - left.matchCount || left.path.localeCompare(right.path))
			.slice(0, limit);
	}

	getStats(): { files: number; errors: number } {
		let errors = 0;

		for (const entry of this.entries.values()) {
			if (entry.error) {
				errors += 1;
			}
		}

		return {
			files: this.entries.size,
			errors,
		};
	}

	isDocxFile(file: TFile): boolean {
		return file.extension.toLowerCase() === 'docx' || isDocxPath(file.path);
	}

	private getIndexPath(): string {
		const dir = this.pluginDir || `${this.app.vault.configDir}/plugins/native-powerpoint-doc-editor`;
		return normalizePath(`${dir}/${INDEX_FILE_NAME}`);
	}
}
