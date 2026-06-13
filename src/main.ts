import { Notice, Platform, Plugin, TAbstractFile, TFile, setIcon } from 'obsidian';
import {
	DEFAULT_SETTINGS,
	DocxidianSettings,
	DocxidianSettingTab,
	getNativePowerPointSettings,
	normalizeDefaultZoom,
	type NativePowerPointSettings,
} from './settings';
import { processDocxEmbeds, registerDocxFileEmbed } from './DocxEmbedLoader';
import { DocxSearchModal } from './DocxSearchModal';
import { DocxView, VIEW_TYPE_DOCX, type DocxEditorSettingsController, type DocxEditorSettingsSnapshot } from './DocxView';
import {
	NativePowerPointView,
	NATIVE_POWERPOINT_VIEW_TYPE,
	POWERPOINT_EXTENSIONS,
	isPowerPointExtension,
} from './NativePowerPointView';
import { configureDocxEditorChunkPaths } from './docxEditorLoader';
import { DocxSearchIndex } from './docxSearchIndex';
import { configureDocxidianLogger, errorLog, getDocxidianLogSnapshot, infoLog, setDocxidianLogSink } from './logger';
import { getDocxEditorLocale, normalizeDocxidianLanguage } from './locales';
import { configureObsidianRuntime } from './obsidianRuntime';

export { createDocxReactMount, DocxFileEmbed, renderDocxEmbeds, hasReviewMarkup } from './docxEditorChunk';

const DOCX_EXTENSIONS = ['docx'];
const DOCX_LOG_AREAS = new Set([
	'chunk',
	'clipboard',
	'copy',
	'diagnostics',
	'editor',
	'export',
	'file',
	'plugin',
	'render',
	'review',
	'save',
	'search',
	'security',
	'settings',
	'view',
]);
type DebugLogScope = 'all' | 'docx';

export default class DocxidianPlugin extends Plugin {
	settings: DocxidianSettings;
	private docxSearchIndex: DocxSearchIndex | null = null;

	async onload() {
		await this.loadSettings();
		configureDocxidianLogger(this.settings.debugLogging);
		infoLog('plugin', 'Plugin loaded', {
			version: this.manifest.version,
			debugLogging: this.settings.debugLogging,
			editorLanguage: this.settings.editorLanguage,
		});
		configureObsidianRuntime({ Notice, Platform, setIcon });

		if (!this.settings.disableDocxFiles) {
			await this.loadDocxSupport();
		} else {
			infoLog('plugin', 'DOCX support disabled by settings');
		}

		if (!this.settings.disablePowerPointFiles) {
			this.loadPowerPointSupport();
		} else {
			infoLog('plugin', 'PowerPoint support disabled by settings');
		}

		this.addCommand({
			id: 'copy-docxidian-debug-log',
			name: 'Copy debug log',
			callback: async () => {
				await this.copyDebugLog();
			},
		});

		this.addSettingTab(new DocxidianSettingTab(this.app, this));

		void this.setupDevHotReload();
	}

	onunload() {
		setDocxidianLogSink(null);
		infoLog('plugin', 'Plugin unloaded');
	}

	/**
	 * Dev-only: mirror every log entry to `dev-debug.log` inside the plugin
	 * folder so the running diagnostics can be inspected from outside Obsidian.
	 * Only used when the `.hotreload` marker is present.
	 */
	private async setupDevFileLog(pluginDir: string): Promise<void> {
		const adapter = this.app.vault.adapter;
		const logPath = `${pluginDir}/dev-debug.log`;
		const appendable = adapter as typeof adapter & {
			append?: (path: string, data: string) => Promise<void>;
		};

		try {
			await adapter.write(logPath, `# session ${new Date().toISOString()}\n`);
		} catch {
			return;
		}

		let queue: Promise<void> = Promise.resolve();
		setDocxidianLogSink((entry) => {
			const line = `${JSON.stringify(entry)}\n`;
			queue = queue
				.then(async () => {
					if (appendable.append) {
						await appendable.append(logPath, line);
					}
				})
				.catch(() => undefined);
		});
		infoLog('plugin', 'Dev file log enabled', { logPath });
	}

	/**
	 * Dev-only self reload. When a `.hotreload` marker file exists in this
	 * plugin's folder, poll the built `main.js` modification time and reload the
	 * plugin whenever it changes (i.e. after a rebuild). This is completely inert
	 * for normal users because the marker file is never shipped — it is only
	 * created in a development vault. Mirrors the convention used by the
	 * community "Hot Reload" plugin without requiring it to be installed.
	 */
	private async setupDevHotReload(): Promise<void> {
		const adapter = this.app.vault.adapter;
		const pluginDir = this.manifest.dir;
		if (!pluginDir) {
			return;
		}

		const markerPath = `${pluginDir}/.hotreload`;
		try {
			if (!(await adapter.exists(markerPath))) {
				return;
			}
		} catch {
			return;
		}

		await this.setupDevFileLog(pluginDir);

		const mainPath = `${pluginDir}/main.js`;
		const stampPath = `${pluginDir}/.build-stamp`;
		// Prefer a content-based build stamp (written by the esbuild deploy step);
		// fall back to main.js mtime when the stamp file is absent.
		const readStamp = async (): Promise<string | null> => {
			try {
				if (await adapter.exists(stampPath)) {
					return (await adapter.read(stampPath)).trim();
				}
			} catch {
				// fall through to mtime
			}
			try {
				const stat = await adapter.stat(mainPath);
				return stat?.mtime != null ? String(stat.mtime) : null;
			} catch {
				return null;
			}
		};

		let lastStamp = await readStamp();
		let reloading = false;
		infoLog('plugin', 'Dev hot reload enabled', { mainPath, stampPath, lastStamp });

		const interval = window.setInterval(async () => {
			if (reloading) {
				return;
			}
			const stamp = await readStamp();
			if (stamp === null || lastStamp === null) {
				lastStamp = stamp;
				return;
			}
			if (stamp === lastStamp) {
				return;
			}

			lastStamp = stamp;
			reloading = true;
			const pluginId = this.manifest.id;
			infoLog('plugin', 'Detected rebuild; reloading plugin', { pluginId });
			try {
				const plugins = (this.app as unknown as {
					plugins?: {
						disablePlugin(id: string): Promise<void>;
						enablePlugin(id: string): Promise<void>;
					};
				}).plugins;
				if (plugins) {
					await plugins.disablePlugin(pluginId);
					await plugins.enablePlugin(pluginId);
					new Notice('Native PowerPoint Doc Editor reloaded.');
				}
			} catch (error) {
				reloading = false;
				errorLog('plugin', 'Hot reload failed', error);
			}
		}, 1000);

		this.registerInterval(interval);
	}

	async loadSettings() {
		const savedSettings = await this.loadData() as Partial<DocxidianSettings> | null;
		const legacySettings = savedSettings as Partial<DocxidianSettings> & {
			powerPointRemoveUnsupportedSvgContent?: unknown;
			powerPointYoloMode?: unknown;
		} | null;
		this.settings = Object.assign({}, DEFAULT_SETTINGS, savedSettings);

		const normalizedLanguage = normalizeDocxidianLanguage(this.settings.editorLanguage);
		const normalizedDefaultZoom = normalizeDefaultZoom(this.settings.defaultZoom);
		const normalizedDebugLogging = this.settings.debugLogging === true;
		const normalizedEnableDocxSearchIndex = this.settings.enableDocxSearchIndex !== false;
		const normalizedAutoIndexDocxSearch = this.settings.autoIndexDocxSearch !== false;
		const normalizedPowerPointAutosaveEnabled = this.settings.powerPointAutosaveEnabled !== false;
		const normalizedPowerPointHideUnsupportedSvgContent =
			typeof savedSettings?.powerPointHideUnsupportedSvgContent === 'boolean'
				? savedSettings.powerPointHideUnsupportedSvgContent
				: legacySettings?.powerPointRemoveUnsupportedSvgContent === true;
		const normalizedPowerPointOpenWithYoloMode =
			typeof savedSettings?.powerPointOpenWithYoloMode === 'boolean'
				? savedSettings.powerPointOpenWithYoloMode
				: legacySettings?.powerPointYoloMode === true;
		const normalizedPowerPointShowInspector = this.settings.powerPointShowInspector === true;
		const normalizedDisableDocxFiles = this.settings.disableDocxFiles === true;
		const normalizedDisablePowerPointFiles = this.settings.disablePowerPointFiles === true;
		const shouldPersistSettings = savedSettings?.editorLanguage !== normalizedLanguage
			|| savedSettings?.defaultZoom !== normalizedDefaultZoom
			|| savedSettings?.debugLogging !== normalizedDebugLogging
			|| savedSettings?.enableDocxSearchIndex !== normalizedEnableDocxSearchIndex
			|| savedSettings?.autoIndexDocxSearch !== normalizedAutoIndexDocxSearch
			|| savedSettings?.powerPointAutosaveEnabled !== normalizedPowerPointAutosaveEnabled
			|| savedSettings?.powerPointHideUnsupportedSvgContent !== normalizedPowerPointHideUnsupportedSvgContent
			|| savedSettings?.powerPointOpenWithYoloMode !== normalizedPowerPointOpenWithYoloMode
			|| savedSettings?.powerPointShowInspector !== normalizedPowerPointShowInspector
			|| legacySettings?.powerPointRemoveUnsupportedSvgContent !== undefined
			|| legacySettings?.powerPointYoloMode !== undefined
			|| savedSettings?.disableDocxFiles !== normalizedDisableDocxFiles
			|| savedSettings?.disablePowerPointFiles !== normalizedDisablePowerPointFiles;

		this.settings.editorLanguage = normalizedLanguage;
		this.settings.defaultZoom = normalizedDefaultZoom;
		this.settings.debugLogging = normalizedDebugLogging;
		this.settings.enableDocxSearchIndex = normalizedEnableDocxSearchIndex;
		this.settings.autoIndexDocxSearch = normalizedAutoIndexDocxSearch;
		this.settings.powerPointAutosaveEnabled = normalizedPowerPointAutosaveEnabled;
		this.settings.powerPointHideUnsupportedSvgContent = normalizedPowerPointHideUnsupportedSvgContent;
		this.settings.powerPointOpenWithYoloMode = normalizedPowerPointOpenWithYoloMode;
		this.settings.powerPointShowInspector = normalizedPowerPointShowInspector;
		this.settings.disableDocxFiles = normalizedDisableDocxFiles;
		this.settings.disablePowerPointFiles = normalizedDisablePowerPointFiles;
		delete (this.settings as unknown as Record<string, unknown>).powerPointRemoveUnsupportedSvgContent;
		delete (this.settings as unknown as Record<string, unknown>).powerPointYoloMode;

		if (shouldPersistSettings) {
			await this.saveSettings();
		}
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	private getDocxSettingsSnapshot(): DocxEditorSettingsSnapshot {
		return {
			authorName: this.settings.authorName,
			editorLanguage: this.settings.editorLanguage,
			showRuler: this.settings.showRuler,
			autosave: this.settings.autosave,
			createBackupsBeforeSave: this.settings.createBackupsBeforeSave,
			defaultZoom: this.settings.defaultZoom,
			enableDocxSearchIndex: this.settings.enableDocxSearchIndex,
			autoIndexDocxSearch: this.settings.autoIndexDocxSearch,
			debugLogging: this.settings.debugLogging,
			disableDocxFiles: this.settings.disableDocxFiles,
		};
	}

	private createDocxSettingsController(): DocxEditorSettingsController {
		const saveDocxSettings = async (refreshViews = false) => {
			await this.saveSettings();
			if (refreshViews) {
				this.refreshDocxViews();
			}
		};

		return {
			getSettings: () => this.getDocxSettingsSnapshot(),
			setAuthorName: async (value) => {
				this.settings.authorName = value.trim() || DEFAULT_SETTINGS.authorName;
				await saveDocxSettings();
			},
			setEditorLanguage: async (value) => {
				this.settings.editorLanguage = normalizeDocxidianLanguage(value);
				await saveDocxSettings(true);
			},
			setShowRuler: async (value) => {
				this.settings.showRuler = value;
				await saveDocxSettings(true);
			},
			setAutosave: async (value) => {
				this.settings.autosave = value;
				await saveDocxSettings(true);
			},
			setCreateBackupsBeforeSave: async (value) => {
				this.settings.createBackupsBeforeSave = value;
				await saveDocxSettings();
			},
			setDefaultZoom: async (value) => {
				this.settings.defaultZoom = normalizeDefaultZoom(value);
				await saveDocxSettings();
			},
			setEnableDocxSearchIndex: async (value) => {
				this.settings.enableDocxSearchIndex = value;
				await saveDocxSettings(false);
				if (value) {
					await this.rebuildDocxSearchIndex(false);
				}
			},
			setAutoIndexDocxSearch: async (value) => {
				this.settings.autoIndexDocxSearch = value;
				await saveDocxSettings(false);
				if (value && this.settings.enableDocxSearchIndex) {
					await this.rebuildDocxSearchIndex(false);
				}
			},
			setDebugLogging: async (value) => {
				this.settings.debugLogging = value;
				configureDocxidianLogger(value);
				infoLog('settings', `Debug logging ${value ? 'enabled' : 'disabled'}`);
				await saveDocxSettings(false);
			},
			setDisableDocxFiles: async (value) => {
				this.settings.disableDocxFiles = value;
				await saveDocxSettings(false);
				new Notice('Reload Obsidian or disable/re-enable this plugin to update DOCX file handling.');
			},
			rebuildDocxSearchIndex: async () => {
				await this.rebuildDocxSearchIndex(true);
			},
			copyDocxLog: async (filePath) => {
				await this.copyDebugLog('docx', filePath);
			},
		};
	}

	private getDebugLogEntries(scope: DebugLogScope) {
		const logs = getDocxidianLogSnapshot();
		if (scope === 'all') {
			return logs;
		}

		return logs.filter((entry) => {
			let serializedData = '';
			try {
				serializedData = entry.data === undefined ? '' : JSON.stringify(entry.data).toLowerCase();
			} catch {
				serializedData = String(entry.data).toLowerCase();
			}
			return DOCX_LOG_AREAS.has(entry.area)
				|| entry.message.toLowerCase().includes('docx')
				|| serializedData.includes('.docx');
		});
	}

	async copyDebugLog(scope: DebugLogScope = 'all', activeDocxPath?: string) {
		const logs = this.getDebugLogEntries(scope);
		const payload = {
			generatedAt: new Date().toISOString(),
			scope,
			activeDocxPath,
			plugin: {
				id: this.manifest.id,
				version: this.manifest.version,
				dir: this.manifest.dir,
			},
			settings: {
				editorLanguage: this.settings.editorLanguage,
				showRuler: this.settings.showRuler,
				autosave: this.settings.autosave,
				createBackupsBeforeSave: this.settings.createBackupsBeforeSave,
				defaultZoom: this.settings.defaultZoom,
				debugLogging: this.settings.debugLogging,
				enableDocxSearchIndex: this.settings.enableDocxSearchIndex,
				autoIndexDocxSearch: this.settings.autoIndexDocxSearch,
				powerPointAutosaveEnabled: this.settings.powerPointAutosaveEnabled,
				powerPointHideUnsupportedSvgContent: this.settings.powerPointHideUnsupportedSvgContent,
				powerPointOpenWithYoloMode: this.settings.powerPointOpenWithYoloMode,
				disableDocxFiles: this.settings.disableDocxFiles,
				disablePowerPointFiles: this.settings.disablePowerPointFiles,
			},
			docxEditorBundle: 'main.js',
			logs,
		};

		try {
			await navigator.clipboard.writeText(JSON.stringify(payload, null, 2));
			const label = scope === 'docx' ? 'DOCX' : 'Native PowerPoint Doc Editor';
			new Notice(`Copied ${payload.logs.length} ${label} log entries.`);
		} catch (error) {
			errorLog('diagnostics', 'Could not copy Native PowerPoint Doc Editor debug log', error);
			new Notice('Could not copy Native PowerPoint Doc Editor debug log. Open the developer console and check window.docxidianDebugLogs.');
		}
	}

	async rebuildDocxSearchIndex(force = false, showNotice = true) {
		if (this.settings.disableDocxFiles) {
			if (showNotice) {
				new Notice('DOCX support is turned off for this plugin. Reload after turning it back on.');
			}
			return;
		}

		if (!this.settings.enableDocxSearchIndex) {
			return;
		}

		if (!this.docxSearchIndex) {
			if (showNotice) {
				new Notice('DOCX search index is not ready yet.');
			}
			return;
		}

		try {
			const stats = await this.docxSearchIndex.rebuild({ force });
			if (showNotice) {
				new Notice(`DOCX search index ready: ${stats.total} files, ${stats.errors} errors.`);
			}
		} catch (error) {
			errorLog('search', 'Could not rebuild DOCX search index', error);
			if (showNotice) {
					new Notice('Could not rebuild DOCX search index. Check the Native PowerPoint Doc Editor debug log.');
			}
		}
	}

	refreshDocxViews() {
		for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)) {
			const view = leaf.view;
			if (view instanceof DocxView) {
				view.refreshSettings();
			}
		}
	}

	refreshPowerPointViews() {
		for (const leaf of this.app.workspace.getLeavesOfType(NATIVE_POWERPOINT_VIEW_TYPE)) {
			const view = leaf.view;
			if (view instanceof NativePowerPointView) {
				view.refreshSettings();
			}
		}
	}

	getPowerPointSettings(): NativePowerPointSettings {
		return getNativePowerPointSettings(this.settings, async (value) => {
			this.settings.powerPointOpenWithYoloMode = value;
			await this.saveSettings();
		});
	}

	private async loadDocxSupport() {
		configureDocxEditorChunkPaths([]);
		this.docxSearchIndex = new DocxSearchIndex(this.app, this.manifest.dir);
		await this.docxSearchIndex.load();

		this.registerView(
			VIEW_TYPE_DOCX,
			(leaf) => new DocxView(
				leaf,
				() => this.settings.authorName,
				() => getDocxEditorLocale(this.settings.editorLanguage),
				() => this.settings.showRuler,
				() => this.settings.autosave,
				() => this.settings.createBackupsBeforeSave,
				() => this.settings.defaultZoom,
				this.createDocxSettingsController(),
			),
		);
		this.registerExtensions(DOCX_EXTENSIONS, VIEW_TYPE_DOCX);

		registerDocxFileEmbed(this, () => getDocxEditorLocale(this.settings.editorLanguage));
		this.registerMarkdownPostProcessor((el, ctx) => {
			processDocxEmbeds(this.app, el, ctx, () => getDocxEditorLocale(this.settings.editorLanguage));
		}, 1000);

		this.addCommand({
			id: 'save-current-docx',
			name: 'Save current docx',
			callback: async () => {
				const docxView = this.app.workspace.getActiveViewOfType(DocxView);
				if (!docxView) {
					new Notice('Open a docx file to save it.');
					return;
				}

				await docxView.saveCurrentDocument();
			},
		});
		this.addCommand({
			id: 'save-current-docx-as',
			name: 'Save current DOCX as...',
			callback: async () => {
				const docxView = this.app.workspace.getActiveViewOfType(DocxView);
				if (!docxView) {
					new Notice('Open a docx file to save a copy.');
					return;
				}

				await docxView.saveCurrentDocumentAs();
			},
		});
		this.addCommand({
			id: 'duplicate-current-docx',
			name: 'Duplicate current DOCX',
			callback: async () => {
				const docxView = this.app.workspace.getActiveViewOfType(DocxView);
				if (!docxView) {
					new Notice('Open a docx file to duplicate it.');
					return;
				}

				await docxView.duplicateCurrentDocument();
			},
		});
		this.addCommand({
			id: 'find-in-current-docx',
			name: 'Find in current docx',
			callback: () => {
				const docxView = this.app.workspace.getActiveViewOfType(DocxView);
				if (!docxView) {
					new Notice('Open a docx file to search it.');
					return;
				}

				docxView.openFindDialog();
			},
		});
		this.addCommand({
			id: 'find-replace-in-current-docx',
			name: 'Find and replace in current docx',
			callback: () => {
				const docxView = this.app.workspace.getActiveViewOfType(DocxView);
				if (!docxView) {
					new Notice('Open a docx file to search it.');
					return;
				}

				docxView.openFindReplaceDialog();
			},
		});
		this.addCommand({
			id: 'search-docx-files',
			name: 'Search DOCX files in vault',
			callback: () => {
				if (!this.settings.enableDocxSearchIndex) {
					new Notice('Turn on the DOCX search index in Native PowerPoint Doc Editor settings first.');
					return;
				}

				if (!this.docxSearchIndex) {
					new Notice('DOCX search index is not ready yet.');
					return;
				}

				new DocxSearchModal(this.app, this.docxSearchIndex).open();
			},
		});
		this.addCommand({
			id: 'rebuild-docx-search-index',
			name: 'Rebuild DOCX search index',
			callback: async () => {
				await this.rebuildDocxSearchIndex(true);
			},
		});

		this.registerDocxSearchEvents();
		this.queueInitialDocxSearchIndex();
	}

	private loadPowerPointSupport() {
		this.registerView(
			NATIVE_POWERPOINT_VIEW_TYPE,
			(leaf) => new NativePowerPointView(leaf, () => this.getPowerPointSettings()),
		);
		this.registerExtensions(POWERPOINT_EXTENSIONS, NATIVE_POWERPOINT_VIEW_TYPE);

		this.addCommand({
			id: 'open-powerpoint-file',
			name: 'Open PowerPoint file',
			callback: () => {
				const file = this.app.workspace.getActiveFile();
				if (!file || !isPowerPointExtension(file.extension)) {
					new Notice('Select a PowerPoint file to open it.');
					return;
				}

				const leaf = this.app.workspace.getLeaf('tab');
				void leaf.openFile(file, { active: true });
			},
		});
		this.addCommand({
			id: 'save-current-powerpoint-file',
			name: 'Save current PowerPoint file',
			callback: async () => {
				const view = this.app.workspace.getActiveViewOfType(NativePowerPointView);
				if (!view) {
					new Notice('Open a PowerPoint file to save it.');
					return;
				}

				await view.saveCurrentPresentation();
			},
		});
	}

	private registerDocxSearchEvents() {
		this.registerEvent(this.app.vault.on('create', file => this.handleDocxSearchFileChanged(file)));
		this.registerEvent(this.app.vault.on('modify', file => this.handleDocxSearchFileChanged(file)));
		this.registerEvent(this.app.vault.on('delete', file => this.handleDocxSearchFileDeleted(file)));
		this.registerEvent(this.app.vault.on('rename', (file, oldPath) => {
			this.handleDocxSearchFileDeleted(oldPath);
			this.handleDocxSearchFileChanged(file);
		}));
	}

	private queueInitialDocxSearchIndex() {
		if (!this.settings.enableDocxSearchIndex || !this.settings.autoIndexDocxSearch) {
			return;
		}

		const timeoutId = window.setTimeout(() => {
			void this.rebuildDocxSearchIndex(false, false);
		}, 1500);

		this.register(() => window.clearTimeout(timeoutId));
	}

	private handleDocxSearchFileChanged(file: TAbstractFile) {
		if (!this.settings.enableDocxSearchIndex || !this.settings.autoIndexDocxSearch || !this.docxSearchIndex) {
			return;
		}

		if (!(file instanceof TFile) || !this.docxSearchIndex.isDocxFile(file)) {
			return;
		}

		void this.docxSearchIndex.indexFile(file);
	}

	private handleDocxSearchFileDeleted(fileOrPath: TAbstractFile | string) {
		if (!this.settings.enableDocxSearchIndex || !this.docxSearchIndex) {
			return;
		}

		const path = typeof fileOrPath === 'string' ? fileOrPath : fileOrPath.path;
		if (!path.toLowerCase().endsWith('.docx')) {
			return;
		}

		void this.docxSearchIndex.removePath(path);
	}

}
