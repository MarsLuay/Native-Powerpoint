import { debugLog, infoLog } from './logger';
import * as docxEditorChunk from './docxEditorChunk';

export type DocxEditorChunkModule = typeof import('./docxEditorChunk');

let editorChunkPromise: Promise<DocxEditorChunkModule> | null = null;

export function configureDocxEditorChunkPaths(paths: string[]) {
	debugLog('chunk', 'DOCX editor is bundled into main.js', { ignoredPaths: paths });
}

export function loadDocxEditorChunk(): Promise<DocxEditorChunkModule> {
	if (!editorChunkPromise) {
		infoLog('chunk', 'Loaded bundled DOCX editor from main.js');
		editorChunkPromise = Promise.resolve(docxEditorChunk);
	}

	return editorChunkPromise;
}
