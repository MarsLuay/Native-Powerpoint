export type DocxidianLogLevel = 'debug' | 'info' | 'warn' | 'error';

export interface DocxidianLogEntry {
	time: string;
	level: DocxidianLogLevel;
	area: string;
	message: string;
	data?: unknown;
}

const MAX_LOG_ENTRIES = 200;
const LOG_PREFIX = '[Native PowerPoint/Doc Editor]';
const DEFAULT_LOG_FILE = 'Library/Logs/native-powerpoint-doc-editor.log';
const LOG_STATE_KEY = '__nativePowerPointDocEditorLogState';

interface DocxidianLogState {
	debugLoggingEnabled: boolean;
	entries: DocxidianLogEntry[];
}

function getGlobalLogHost() {
	return globalThis as typeof globalThis & {
		[LOG_STATE_KEY]?: DocxidianLogState;
	};
}

function getWindowWithLogs() {
	if (typeof window === 'undefined') {
		return null;
	}

	return window as Window & {
		docxidianDebugLogs?: DocxidianLogEntry[];
		docxidianDebugLogging?: boolean;
	};
}

function getLogState() {
	const host = getGlobalLogHost();
	if (!host[LOG_STATE_KEY]) {
		const logsWindow = getWindowWithLogs();
		host[LOG_STATE_KEY] = {
			debugLoggingEnabled: logsWindow?.docxidianDebugLogging === true,
			entries: Array.isArray(logsWindow?.docxidianDebugLogs) ? logsWindow.docxidianDebugLogs : [],
		};
	}

	return host[LOG_STATE_KEY];
}

function syncWindowLogState() {
	const logsWindow = getWindowWithLogs();
	if (!logsWindow) {
		return;
	}

	const state = getLogState();
	logsWindow.docxidianDebugLogging = state.debugLoggingEnabled;
	logsWindow.docxidianDebugLogs = state.entries;
}

function shouldPrint(level: DocxidianLogLevel) {
	return getLogState().debugLoggingEnabled || level === 'warn' || level === 'error';
}

function writeConsole(level: DocxidianLogLevel, area: string, message: string, data?: unknown) {
	if (!shouldPrint(level)) {
		return;
	}

	const consoleMethod = level === 'debug' ? console.debug
		: level === 'info' ? console.info
			: level === 'warn' ? console.warn
				: console.error;

	if (data === undefined) {
		consoleMethod.call(console, `${LOG_PREFIX} ${area}: ${message}`);
	} else {
		consoleMethod.call(console, `${LOG_PREFIX} ${area}: ${message}`, data);
	}
}

function normalizeLogData(data: unknown): unknown {
	if (data instanceof Error) {
		return {
			name: data.name,
			message: data.message,
			stack: data.stack,
		};
	}

	return data;
}

function getLogFilePath() {
	const home = typeof process !== 'undefined' ? process.env?.HOME : undefined;
	return home ? `${home}/${DEFAULT_LOG_FILE}` : null;
}

function appendFileLog(entry: DocxidianLogEntry) {
	const logFile = getLogFilePath();
	if (!logFile) {
		return;
	}

	try {
		const fs = require('fs') as typeof import('fs');
		const path = require('path') as typeof import('path');
		fs.mkdirSync(path.dirname(logFile), { recursive: true });
		fs.appendFileSync(logFile, `${JSON.stringify(entry)}\n`);
	} catch {
		// Console/in-memory logs are still available if file logging is blocked.
	}
}

export function configureDocxidianLogger(enabled: boolean) {
	getLogState().debugLoggingEnabled = enabled;
	syncWindowLogState();
}

export function logDocxidian(level: DocxidianLogLevel, area: string, message: string, data?: unknown) {
	const state = getLogState();
	const entry: DocxidianLogEntry = {
		time: new Date().toISOString(),
		level,
		area,
		message,
		data: normalizeLogData(data),
	};

	state.entries.push(entry);
	if (state.entries.length > MAX_LOG_ENTRIES) {
		state.entries.splice(0, state.entries.length - MAX_LOG_ENTRIES);
	}
	appendFileLog(entry);
	syncWindowLogState();

	writeConsole(level, area, message, data);
}

export function debugLog(area: string, message: string, data?: unknown) {
	logDocxidian('debug', area, message, data);
}

export function infoLog(area: string, message: string, data?: unknown) {
	logDocxidian('info', area, message, data);
}

export function warnLog(area: string, message: string, data?: unknown) {
	logDocxidian('warn', area, message, data);
}

export function errorLog(area: string, message: string, data?: unknown) {
	logDocxidian('error', area, message, data);
}

export function getDocxidianLogSnapshot() {
	return getLogState().entries.slice();
}
