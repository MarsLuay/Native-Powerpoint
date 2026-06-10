const assert = require('node:assert/strict');
const fs = require('node:fs');
const Module = require('node:module');
const os = require('node:os');
const path = require('node:path');

const projectRoot = path.resolve(__dirname, '..');
const inferredVaultRoot = path.resolve(projectRoot, '..', '..');
const vaultRoot = process.env.DOCXIDIAN_VAULT_ROOT
	? path.resolve(process.env.DOCXIDIAN_VAULT_ROOT)
	: fs.existsSync(path.join(inferredVaultRoot, '.obsidian'))
		? inferredVaultRoot
		: projectRoot;
let temporaryPluginDir = null;
const installedPluginDir = path.resolve(vaultRoot, '.obsidian', 'plugins', 'native-powerpoint-doc-editor');
const pluginDir = resolvePluginDir();
const maxDocxFiles = Number(process.env.DOCXIDIAN_SMOKE_DOCX_LIMIT ?? 5);

const originalLoad = Module._load;
const originalConsole = {
	debug: console.debug,
	info: console.info,
	warn: console.warn,
	error: console.error,
};
const originalNavigatorDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
const originalWindowDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'window');
const capturedLogs = [];
let copiedClipboardText = '';
const DEFAULT_PLUGIN_DATA = {
	autosave: true,
	createBackupsBeforeSave: false,
	debugLogging: true,
	defaultZoom: 1,
	disableDocxFiles: false,
	disablePowerPointFiles: false,
	editorLanguage: 'en',
	enableDocxSearchIndex: true,
	autoIndexDocxSearch: true,
	powerPointAutosaveEnabled: true,
	powerPointHideUnsupportedSvgContent: false,
	powerPointOpenWithYoloMode: false,
	showRuler: false,
};
let pluginData = { ...DEFAULT_PLUGIN_DATA };

function resolvePluginDir() {
	if (process.env.DOCXIDIAN_PLUGIN_DIR) {
		return path.resolve(process.env.DOCXIDIAN_PLUGIN_DIR);
	}

	if (fs.existsSync(path.join(installedPluginDir, 'main.js'))) {
		return installedPluginDir;
	}

	temporaryPluginDir = fs.mkdtempSync(path.join(os.tmpdir(), 'native-powerpoint-doc-editor-smoke-'));
	for (const fileName of ['main.js', 'manifest.json']) {
		fs.copyFileSync(path.join(projectRoot, fileName), path.join(temporaryPluginDir, fileName));
	}
	return temporaryPluginDir;
}

function captureConsole(level) {
	return (...args) => {
		capturedLogs.push({ level, args: args.map(String) });
	};
}

function createObsidianStub() {
	class Component {
		addChild(child) {
			return child;
		}

		register() {}

		registerDomEvent() {}

		registerEvent(eventRef) {
			return eventRef;
		}

		load() {}

		unload() {}

		onload() {}

		onunload() {}
	}

	class Plugin extends Component {
		constructor(app, manifest) {
			super();
			this.app = app;
			this.manifest = manifest;
			this.registeredViews = [];
			this.registeredExtensions = [];
			this.commands = [];
		}

		addCommand(command) {
			this.commands.push(command);
			return command;
		}

		addSettingTab(settingTab) {
			this.settingTab = settingTab;
		}

		loadData() {
			return Promise.resolve({ ...pluginData });
		}

		registerExtensions(extensions, viewType) {
			this.registeredExtensions.push({ extensions, viewType });
		}

		registerMarkdownPostProcessor(callback, sortOrder) {
			this.markdownPostProcessor = { callback, sortOrder };
		}

		registerView(viewType, factory) {
			this.registeredViews.push({ viewType, factory });
		}

		saveData() {
			return Promise.resolve();
		}
	}

	class PluginSettingTab {
		constructor(app, plugin) {
			this.app = app;
			this.plugin = plugin;
			this.containerEl = {};
		}
	}

	class Setting {}
	class Modal extends Component {}
	class FileView extends Component {
		constructor(leaf) {
			super();
			this.leaf = leaf;
			this.app = leaf?.app;
		}
	}
	class MarkdownRenderChild extends Component {
		constructor(containerEl) {
			super();
			this.containerEl = containerEl;
		}
	}
	class TFile {}

	return {
		App: class App {},
		Component,
		FileView,
		MarkdownRenderChild,
		Modal,
		Notice: class Notice {
			constructor(message) {
				this.message = message;
			}
		},
		Platform: { isMacOS: true },
		Plugin,
		PluginSettingTab,
		Setting,
		TFile,
		WorkspaceLeaf: class WorkspaceLeaf {},
		normalizePath: (value) => value.replace(/\\/g, '/').replace(/\/+/g, '/'),
	};
}

function createAppStub() {
	const createdVaultFiles = [];
	const embedRegistry = {
		registerExtension(extension, creator) {
			embedRegistry.extension = extension;
			embedRegistry.creator = creator;
		},
		unregisterExtension() {},
	};

	const app = {
		createdVaultFiles,
		embedRegistry,
		metadataCache: {
			getFirstLinkpathDest() {
				return null;
			},
		},
		vault: {
			adapter: {
				basePath: vaultRoot,
				exists() {
					return Promise.resolve(false);
				},
				read() {
					return Promise.resolve('');
				},
				write() {
					return Promise.resolve();
				},
			},
			getAbstractFileByPath() {
				return null;
			},
			getFiles() {
				return [];
			},
			create(path, contents) {
				const file = { path };
				createdVaultFiles.push({ path, binary: false, size: contents.length });
				return Promise.resolve(file);
			},
			createBinary(path, contents) {
				const file = { path };
				createdVaultFiles.push({ path, binary: true, size: contents.byteLength });
				return Promise.resolve(file);
			},
			on() {
				return {};
			},
		},
		workspace: {
			getActiveViewOfType() {
				return null;
			},
			getLeavesOfType() {
				return [];
			},
		},
	};
	return app;
}

function installObsidianStub() {
	const obsidianStub = createObsidianStub();
	Module._load = function loadWithObsidianStub(request, parent, isMain) {
		if (request === 'obsidian') {
			return obsidianStub;
		}

		return originalLoad.call(this, request, parent, isMain);
	};
}

function restoreEnvironment() {
	Module._load = originalLoad;
	console.debug = originalConsole.debug;
	console.info = originalConsole.info;
	console.warn = originalConsole.warn;
	console.error = originalConsole.error;
	if (originalNavigatorDescriptor) {
		Object.defineProperty(globalThis, 'navigator', originalNavigatorDescriptor);
	} else {
		delete globalThis.navigator;
	}
	if (originalWindowDescriptor) {
		Object.defineProperty(globalThis, 'window', originalWindowDescriptor);
	} else {
		delete globalThis.window;
	}
	if (temporaryPluginDir) {
		fs.rmSync(temporaryPluginDir, { recursive: true, force: true });
		temporaryPluginDir = null;
	}
}

function assertPluginFiles() {
	for (const fileName of ['main.js', 'manifest.json']) {
		const filePath = path.join(pluginDir, fileName);
		assert.ok(fs.existsSync(filePath), `Missing installed plugin file: ${filePath}`);
	}
}

function toArrayBuffer(buffer) {
	return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function findDocxFiles(dir, limit, results = []) {
	if (results.length >= limit) {
		return results;
	}

	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		if (results.length >= limit) {
			break;
		}

		if (entry.name === '.obsidian' || entry.name === 'node_modules' || entry.name === '.git') {
			continue;
		}

		const entryPath = path.join(dir, entry.name);
		if (entry.isDirectory()) {
			findDocxFiles(entryPath, limit, results);
		} else if (entry.isFile() && entry.name.toLowerCase().endsWith('.docx')) {
			results.push(entryPath);
		}
	}

	return results;
}

async function runSmoke() {
	console.debug = captureConsole('debug');
	console.info = captureConsole('info');
	console.warn = captureConsole('warn');
	console.error = captureConsole('error');
	Object.defineProperty(globalThis, 'navigator', {
		configurable: true,
		value: {
			clipboard: {
				writeText(text) {
					copiedClipboardText = text;
					return Promise.resolve();
				},
			},
		},
	});
	Object.defineProperty(globalThis, 'window', {
		configurable: true,
		value: {
			clearTimeout() {},
			setTimeout() {
				return 0;
			},
		},
	});

	assertPluginFiles();
	installObsidianStub();
	const manifest = JSON.parse(fs.readFileSync(path.join(pluginDir, 'manifest.json'), 'utf8'));

	const pluginModule = require(path.join(pluginDir, 'main.js'));
	const PluginCtor = pluginModule.default;
	assert.equal(typeof PluginCtor, 'function', 'Installed main.js should export a plugin class.');

	const app = createAppStub();
	const plugin = new PluginCtor(app, { ...manifest, dir: '.obsidian/plugins/native-powerpoint-doc-editor' });

	await plugin.onload();
	assert.equal(plugin.registeredViews.length, 2, 'Plugin should register the DOCX and PowerPoint views.');
	assert.ok(plugin.registeredViews.some((view) => view.viewType === 'docxidian-docx-view'), 'Plugin should register the DOCX view.');
	assert.ok(plugin.registeredViews.some((view) => view.viewType === 'native-powerpoint-view'), 'Plugin should register the PowerPoint view.');
	assert.ok(
		plugin.registeredExtensions.some((entry) => entry.viewType === 'docxidian-docx-view' && entry.extensions.includes('docx')),
		'Plugin should register DOCX file extensions.'
	);
	assert.ok(
		plugin.registeredExtensions.some((entry) => entry.viewType === 'native-powerpoint-view' && entry.extensions.includes('pptx')),
		'Plugin should register PowerPoint file extensions.'
	);
	const copyLogCommand = plugin.commands.find((command) => command.id === 'copy-docxidian-debug-log');
	assert.ok(copyLogCommand, 'Plugin should register the copy debug log command.');
	assert.ok(plugin.commands.some((command) => command.id === 'search-docx-files'), 'Plugin should register the vault-wide DOCX search command.');
	assert.ok(plugin.commands.some((command) => command.id === 'rebuild-docx-search-index'), 'Plugin should register the DOCX search rebuild command.');
	assert.ok(plugin.commands.some((command) => command.id === 'save-current-powerpoint-file'), 'Plugin should register the PowerPoint save command.');
	assert.ok(capturedLogs.some((entry) => entry.args.join(' ').includes('[Native PowerPoint Doc Editor] plugin: Plugin loaded')), 'Debug logging should emit a plugin loaded entry.');
	assert.ok(capturedLogs.some((entry) => entry.args.join(' ').includes('[Native PowerPoint Doc Editor] chunk: DOCX editor is bundled into main.js')), 'Debug logging should emit bundled DOCX editor mode.');

	const docxViewFactory = plugin.registeredViews.find((view) => view.viewType === 'docxidian-docx-view')?.factory;
	assert.equal(typeof docxViewFactory, 'function', 'Plugin should expose a DOCX view factory.');
	const docxView = docxViewFactory({ app });
	const sourceDocx = {
		path: 'Folder/Original.docx',
		parent: { path: 'Folder' },
		basename: 'Original',
		name: 'Original.docx',
	};
	assert.equal(docxView.getAvailableExportPath(sourceDocx, 'pdf'), 'Folder/Original.pdf', 'PDF exports should default to the original title.');
	assert.equal(docxView.getAvailableExportPath(sourceDocx, 'docx'), 'Folder/Original 2.docx', 'DOCX exports should not default to overwriting the open file.');
	docxView.file = { path: 'Original.docx', parent: { path: '/' }, basename: 'Original', name: 'Original.docx' };
	docxView.getReactHandle = () => ({
		exportBuffer() {
			return Promise.resolve(toArrayBuffer(Buffer.from('docx')));
		},
		exportRenderedPdf() {
			return Promise.resolve(toArrayBuffer(Buffer.from('%PDF-1.7\n')));
		},
	});
	const didExportPdf = await docxView.createCurrentDocumentExport('Original.pdf', 'pdf');
	assert.equal(didExportPdf, true, 'PDF export writer should report success.');
	assert.deepEqual(app.createdVaultFiles.at(-1), { path: 'Original.pdf', binary: true, size: 9 }, 'PDF export should create the selected output file.');
	docxView.getReactHandle = () => ({
		exportBuffer() {
			return Promise.resolve(toArrayBuffer(Buffer.from('docx')));
		},
		exportRenderedPdf() {
			return Promise.resolve(null);
		},
	});
	const createdFileCountBeforeFailedPdf = app.createdVaultFiles.length;
	const didExportFailedPdf = await docxView.createCurrentDocumentExport('Fallback.pdf', 'pdf');
	assert.equal(didExportFailedPdf, false, 'PDF export should fail when formatted rendering fails.');
	assert.equal(app.createdVaultFiles.length, createdFileCountBeforeFailedPdf, 'PDF export should not create a text-only fallback file.');

	pluginData = {
		...DEFAULT_PLUGIN_DATA,
		disableDocxFiles: true,
		disablePowerPointFiles: true,
	};
	const disabledPlugin = new PluginCtor(createAppStub(), { ...manifest, dir: '.obsidian/plugins/native-powerpoint-doc-editor' });
	await disabledPlugin.onload();
	assert.equal(disabledPlugin.registeredViews.length, 0, 'Disabled file handoff should not register DOCX or PowerPoint views.');
	assert.equal(disabledPlugin.registeredExtensions.length, 0, 'Disabled file handoff should not register DOCX or PowerPoint extensions.');
	assert.ok(!disabledPlugin.commands.some((command) => command.id === 'save-current-docx'), 'Disabled DOCX handoff should skip DOCX commands.');
	assert.ok(!disabledPlugin.commands.some((command) => command.id === 'save-current-powerpoint-file'), 'Disabled PPTX handoff should skip PowerPoint commands.');
	assert.ok(disabledPlugin.commands.some((command) => command.id === 'copy-docxidian-debug-log'), 'Disabled file handoff should keep diagnostics available.');
	pluginData = { ...DEFAULT_PLUGIN_DATA };

	await copyLogCommand.callback();
	const copiedDiagnostics = JSON.parse(copiedClipboardText);
	assert.equal(copiedDiagnostics.plugin.id, 'native-powerpoint-doc-editor', 'Copied diagnostics should include plugin metadata.');
	assert.ok(Array.isArray(copiedDiagnostics.logs), 'Copied diagnostics should include log entries.');
	assert.ok(copiedDiagnostics.logs.length > 0, 'Copied diagnostics should include at least one log entry.');

	const chunk = pluginModule;
	for (const exportName of ['createDocxReactMount', 'DocxFileEmbed', 'renderDocxEmbeds', 'hasReviewMarkup']) {
		assert.equal(typeof chunk[exportName], 'function', `main.js should export bundled ${exportName}.`);
	}

	const docxFiles = process.env.DOCXIDIAN_SMOKE_DOCX
		? [path.resolve(process.env.DOCXIDIAN_SMOKE_DOCX)]
		: findDocxFiles(vaultRoot, maxDocxFiles);

	for (const docxFile of docxFiles) {
		const buffer = fs.readFileSync(docxFile);
		const result = await chunk.hasReviewMarkup(toArrayBuffer(buffer));
		assert.equal(typeof result, 'boolean', `hasReviewMarkup should return a boolean for ${docxFile}.`);
	}

	return {
		logCount: capturedLogs.length,
		docxCount: docxFiles.length,
		docxFiles: docxFiles.map((filePath) => path.relative(vaultRoot, filePath)),
	};
}

runSmoke()
	.then((result) => {
		restoreEnvironment();
	console.log(`Native PowerPoint Doc Editor smoke passed: ${result.logCount} logs captured; ${result.docxCount} DOCX file(s) inspected.`);
	if (result.docxCount === 0) {
		console.log('- DOCX review-markup scan skipped: set DOCXIDIAN_SMOKE_DOCX or DOCXIDIAN_VAULT_ROOT to include sample documents.');
	}
	for (const docxFile of result.docxFiles) {
			console.log(`- ${docxFile}`);
		}
	})
	.catch((error) => {
		restoreEnvironment();
		console.error(error);
		process.exitCode = 1;
	});
