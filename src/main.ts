import { Notice, Platform, Plugin, setIcon } from 'obsidian';
import {
	DEFAULT_SETTINGS,
	DocxidianSettingTab,
	getNativePowerPointSettings,
	normalizeDefaultZoom,
	type NativePowerPointSettings,
	type DocxidianSettings,
} from './settings';
import type { DocxEditorSettingsController, DocxEditorSettingsSnapshot } from './DocxView';
import { DocxSearchIndex } from './docxSearchIndex';
import {
	configureDocxidianLogger,
	errorLog,
	getDocxidianLogSnapshot,
	getDocxidianLogStats,
	infoLog,
	setDocxidianLogSink,
} from './logger';
import { configureObsidianRuntime, configureChromiumVersionReader } from './obsidianRuntime';
import { loadDocxEditorLocale, normalizeDocxidianLanguage, preloadDocxEditorLocale } from './locales';
import { configureForceJsBackendOverrideReader } from './powerpoint/forceJsBackend';

type DocxSupportModule = typeof import('./docxSupport');
type PptxSupportModule = typeof import('./pptxSupport');

const DOCX_LOG_AREAS = new Set([
	'backup',
	'chunk',
	'clipboard',
	'copy',
	'diagnostics',
	'editor',
	'embed',
	'export',
	'file',
	'font-preservation',
	'ime',
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

let docxSupportModule: DocxSupportModule | null = null;
let pptxSupportModule: PptxSupportModule | null = null;
let docxSupportModulePromise: Promise<DocxSupportModule> | null = null;
let pptxSupportModulePromise: Promise<PptxSupportModule> | null = null;

function loadDocxSupportModule(): Promise<DocxSupportModule> {
	if (docxSupportModule) {
		return Promise.resolve(docxSupportModule);
	}

	docxSupportModulePromise ??= import('./docxSupport').then((module) => {
		docxSupportModule = module;
		return module;
	});
	return docxSupportModulePromise;
}

function loadPptxSupportModule(): Promise<PptxSupportModule> {
	if (pptxSupportModule) {
		return Promise.resolve(pptxSupportModule);
	}

	pptxSupportModulePromise ??= import('./pptxSupport').then((module) => {
		pptxSupportModule = module;
		return module;
	});
	return pptxSupportModulePromise;
}

export default class DocxidianPlugin extends Plugin {
	settings: DocxidianSettings;
	private docxSearchIndex: DocxSearchIndex | null = null;
	private forceJsBackendDevOverride = false;

	setDocxSearchIndex(index: DocxSearchIndex) {
		this.docxSearchIndex = index;
	}

	async onload() {
		await this.loadSettings();
		configureDocxidianLogger(this.settings.debugLogging);
		infoLog('plugin', 'Plugin loaded', {
			version: this.manifest.version,
			debugLogging: this.settings.debugLogging,
			editorLanguage: this.settings.editorLanguage,
		});
		configureObsidianRuntime({ Notice, Platform, setIcon });
		configureChromiumVersionReader(() => {
			if (!Platform.isDesktop) {
				return null;
			}
			try {
				return typeof process !== 'undefined' ? process.versions?.chrome ?? null : null;
			} catch {
				return null;
			}
		});
		configureForceJsBackendOverrideReader(() => this.forceJsBackendDevOverride);

		if (!this.settings.disableDocxFiles) {
			await this.loadDocxSupport();
		} else {
			infoLog('plugin', 'DOCX support disabled by settings');
		}

		if (!this.settings.disablePowerPointFiles) {
			await this.loadPowerPointSupport();
		} else {
			infoLog('plugin', 'PowerPoint support disabled by settings');
		}

		infoLog('diagnostics', 'Feature diagnostics initialized', {
			schemaVersion: 1,
			docx: [
				'lifecycle', 'load', 'save', 'autosave', 'export', 'copy', 'rename',
				'mode', 'zoom', 'find-replace', 'clipboard', 'table', 'image',
				'font', 'font-size', 'review', 'hidden-text', 'search', 'embed',
				'pagination', 'table-cell-font-preservation', 'japanese-ime-anchor',
			],
			powerPoint: [
				'lifecycle', 'load', 'save', 'autosave', 'recovery', 'render',
				'navigation', 'slide-operations', 'history', 'find-replace',
				'clipboard', 'selection', 'arrange', 'insert', 'text-editing',
				'text-formatting', 'inspector', 'export', 'security',
			],
		});

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
		infoLog('plugin', 'Plugin unloaded');
		setDocxidianLogSink(null);
	}

	private async setupDevFileLog(pluginDir: string): Promise<void> {
		const adapter = this.app.vault.adapter;
		const logPath = `${pluginDir}/dev-debug.log`;
		const appendable = adapter as typeof adapter & {
			append?: (path: string, data: string) => Promise<void>;
		};

		try {
			const retainedEntries = getDocxidianLogSnapshot();
			const initialLines = retainedEntries.map((entry) => `${JSON.stringify(entry)}\n`).join('');
			await adapter.write(
				logPath,
				`# session ${new Date().toISOString()}\n${initialLines}`,
			);
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
						return;
					}
					const existing = await adapter.read(logPath);
					await adapter.write(logPath, `${existing}${line}`);
				})
				.catch(() => undefined);
		});
		infoLog('plugin', 'Dev file log enabled', {
			logPath,
			startupEntriesCopied: getDocxidianLogSnapshot().length,
			logStats: getDocxidianLogStats(),
		});
	}

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

		const pollForRebuild = async (): Promise<void> => {
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
		};

		const interval = window.setInterval(() => {
			void pollForRebuild();
		}, 1000);

		this.registerInterval(interval);
	}

	setForceJsBackendDevOverride(enabled: boolean): void {
		this.forceJsBackendDevOverride = enabled;
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
		const normalizedEnableDocxSearchIndex = this.settings.enableDocxSearchIndex === true;
		const normalizedAutoIndexDocxSearch = this.settings.autoIndexDocxSearch === true;
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
				preloadDocxEditorLocale(this.settings.editorLanguage);
				await loadDocxEditorLocale(this.settings.editorLanguage);
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
			logStats: getDocxidianLogStats(),
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
		if (!this.docxSearchIndex) {
			if (showNotice) {
				new Notice('DOCX search index is not ready yet.');
			}
			return;
		}

		const docx = docxSupportModule ?? await loadDocxSupportModule();
		await docx.rebuildDocxSearchIndex(this, this.docxSearchIndex, force, showNotice);
	}

	refreshDocxViews() {
		if (!docxSupportModule) {
			return;
		}

		docxSupportModule.refreshDocxViews(this);
	}

	refreshPowerPointViews() {
		if (!pptxSupportModule) {
			return;
		}

		pptxSupportModule.refreshPowerPointViews(this);
	}

	getPowerPointSettings(): NativePowerPointSettings {
		return getNativePowerPointSettings(this.settings, async (value) => {
			this.settings.powerPointOpenWithYoloMode = value;
			await this.saveSettings();
		});
	}

	private async loadDocxSupport() {
		const docx = await loadDocxSupportModule();
		await docx.registerDocxSupport(this, () => this.createDocxSettingsController());
	}

	private async loadPowerPointSupport() {
		const pptx = await loadPptxSupportModule();
		pptx.registerPowerPointSupport(this, () => this.getPowerPointSettings());
	}
}
