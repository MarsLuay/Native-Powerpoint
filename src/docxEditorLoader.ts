import { debugLog, errorLog, infoLog } from './logger';

export type DocxEditorChunkModule = typeof import('./docxEditorChunk');

let editorChunkPromise: Promise<DocxEditorChunkModule> | null = null;
let editorChunkPaths = ['./docx-editor.js'];

function getErrorMessage(error: unknown) {
	return error instanceof Error ? error.message : String(error);
}

function requireDocxEditorChunk() {
	const errors: string[] = [];

	for (const chunkPath of editorChunkPaths) {
		try {
			debugLog('chunk', `Trying DOCX editor chunk at ${chunkPath}`);
			const chunk = require(chunkPath) as Partial<DocxEditorChunkModule>;
			if (typeof chunk.createDocxReactMount === 'function') {
				infoLog('chunk', `Loaded DOCX editor chunk from ${chunkPath}`);
				return chunk as DocxEditorChunkModule;
			}

			errors.push(`${chunkPath}: missing expected exports`);
			debugLog('chunk', `DOCX editor chunk missing expected exports at ${chunkPath}`, {
				exports: Object.keys(chunk),
			});
		} catch (error) {
			errors.push(`${chunkPath}: ${getErrorMessage(error)}`);
			debugLog('chunk', `Failed to load DOCX editor chunk from ${chunkPath}`, error);
		}
	}

	const failure = new Error(`Unable to load docx-editor.js. Tried ${errors.join('; ')}`);
	errorLog('chunk', 'Unable to load DOCX editor chunk from any configured path', failure);
	throw failure;
}

export function configureDocxEditorChunkPaths(paths: string[]) {
	editorChunkPaths = Array.from(new Set([...paths, './docx-editor.js']));
	editorChunkPromise = null;
	debugLog('chunk', 'Configured DOCX editor chunk paths', editorChunkPaths);
}

export function loadDocxEditorChunk(): Promise<DocxEditorChunkModule> {
	if (!editorChunkPromise) {
		editorChunkPromise = Promise.resolve()
			.then(requireDocxEditorChunk)
			.catch((error) => {
				editorChunkPromise = null;
				throw error;
			});
	}

	return editorChunkPromise;
}
