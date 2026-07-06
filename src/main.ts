import { Notice, Platform, Plugin, setIcon } from 'obsidian';
import {
	NativePowerPointDocEditorSettingTab,
	getNativePowerPointSettings,
	normalizeEditorThemePreference,
	readNativePowerPointDocEditorSettings,
	resolveEditorThemePreference,
	type NativePowerPointSettings,
	type NativePowerPointDocEditorSettings,
	type EditorThemeResolution,
} from './settings';
import { DocxSearchIndex } from './docxSearchIndex';
import {
	configureNativePowerPointDocEditorLogger,
	errorLog,
	getNativePowerPointDocEditorLogSnapshot,
	getNativePowerPointDocEditorLogStats,
	infoLog,
	setNativePowerPointDocEditorLogSink,
} from './logger';
import { configureObsidianRuntime, configureChromiumVersionReader } from './obsidianRuntime';
import { loadDocxEditorLocale, preloadDocxEditorLocale, resolveAutomaticDocxEditorLanguage, type NativePowerPointDocEditorLanguage } from './locales';
import { initPluginI18n, resolvePluginLocale } from './i18n/pluginI18n';
import { getObsidianLocale } from './i18n/obsidianLocale';
import type { PluginI18nService } from './i18n/I18nService';
import { showI18nNotice } from './i18n/notify';
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
	'lifecycle',
	'load',
	'observer',
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
const EDITOR_THEME_CLASSES = [
	'native-powerpoint-doc-editor-theme-system',
	'native-powerpoint-doc-editor-theme-light',
	'native-powerpoint-doc-editor-theme-dark',
];
const RESOLVED_EDITOR_THEME_CLASSES = [
	'native-powerpoint-doc-editor-theme-resolved-light',
	'native-powerpoint-doc-editor-theme-resolved-dark',
];

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

export default class NativePowerPointDocEditorPlugin extends Plugin {
	settings: NativePowerPointDocEditorSettings;
	i18n: PluginI18nService | null = null;
	private docxSearchIndex: DocxSearchIndex | null = null;
	private forceJsBackendDevOverride = false;
	private lastAppliedResolvedEditorTheme?: EditorThemeResolution;
	private editorThemeObserver: MutationObserver | null = null;

	setDocxSearchIndex(index: DocxSearchIndex) {
		this.docxSearchIndex = index;
	}

	getI18n(): PluginI18nService | null {
		return this.i18n;
	}

	getResolvedLocale(): string {
		return this.i18n?.locale ?? 'en';
	}

	getResolvedDocxEditorLanguage(): NativePowerPointDocEditorLanguage {
		return resolveAutomaticDocxEditorLanguage(getObsidianLocale());
	}

	async onload() {
		await this.loadSettings();
		await initPluginI18n(this, await resolvePluginLocale(this));
		const docxLanguage = this.getResolvedDocxEditorLanguage();
		preloadDocxEditorLocale(docxLanguage);
		void loadDocxEditorLocale(docxLanguage);
		this.applyEditorThemePreference();
		configureNativePowerPointDocEditorLogger(this.settings.debugLogging);
		infoLog('plugin', 'Plugin loaded', {
			version: this.manifest.version,
			debugLogging: this.settings.debugLogging,
			locale: this.getResolvedLocale(),
			editorTheme: this.settings.editorTheme,
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
				'lifecycle', 'load', 'observer', 'save', 'autosave', 'export', 'copy', 'rename',
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
			id: 'copy-debug-log',
			name: 'Copy debug log',
			callback: async () => {
				await this.copyDebugLog();
			},
		});

		this.addSettingTab(new NativePowerPointDocEditorSettingTab(this.app, this));
		this.registerEditorThemeObserver();
		void this.setupDevHotReload();
	}

	onunload() {
		infoLog('plugin', 'Plugin unloaded');
		this.editorThemeObserver?.disconnect();
		this.editorThemeObserver = null;
		activeDocument.body.removeClasses([...EDITOR_THEME_CLASSES, ...RESOLVED_EDITOR_THEME_CLASSES]);
		activeDocument.body.removeAttribute('data-native-powerpoint-doc-editor-theme');
		activeDocument.body.removeAttribute('data-native-powerpoint-doc-editor-resolved-theme');
		setNativePowerPointDocEditorLogSink(null);
	}

	private async setupDevFileLog(pluginDir: string): Promise<void> {
		const adapter = this.app.vault.adapter;
		const logPath = `${pluginDir}/dev-debug.log`;
		const appendable = adapter as typeof adapter & {
			append?: (path: string, data: string) => Promise<void>;
		};

		try {
			const retainedEntries = getNativePowerPointDocEditorLogSnapshot();
			const initialLines = retainedEntries.map((entry) => `${JSON.stringify(entry)}\n`).join('');
			await adapter.write(
				logPath,
				`# session ${new Date().toISOString()}\n${initialLines}`,
			);
		} catch {
			return;
		}

		let queue: Promise<void> = Promise.resolve();
		setNativePowerPointDocEditorLogSink((entry) => {
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
			startupEntriesCopied: getNativePowerPointDocEditorLogSnapshot().length,
			logStats: getNativePowerPointDocEditorLogStats(),
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
					showI18nNotice(this.getI18n(), 'settings:plugin.reloaded');
				}
			} catch (error) {
				reloading = false;
				errorLog('plugin', 'Hot reload failed', error);
			}
		};

		this.registerInterval(window.setInterval(() => {
			void pollForRebuild();
		}, 1000));
	}

	setForceJsBackendDevOverride(enabled: boolean): void {
		this.forceJsBackendDevOverride = enabled;
	}

	async loadSettings() {
		const savedSettings = await this.loadData() as Record<string, unknown> | null;
		const { settings, shouldPersistSettings } = readNativePowerPointDocEditorSettings(
			savedSettings,
			this.getObsidianThemeResolution(),
		);
		this.settings = settings;

		if (shouldPersistSettings) {
			await this.saveSettings();
		}
	}

	async saveSettings() {
		this.applyEditorThemePreference();
		await this.saveData(this.settings);
	}

	private applyEditorThemePreference(): boolean {
		const editorTheme = normalizeEditorThemePreference(this.settings.editorTheme);
		const resolvedTheme = this.resolveCurrentEditorTheme();
		const themeClass = `native-powerpoint-doc-editor-theme-${editorTheme}`;
		const resolvedThemeClass = `native-powerpoint-doc-editor-theme-resolved-${resolvedTheme}`;
		for (const className of EDITOR_THEME_CLASSES) {
			activeDocument.body.classList.toggle(className, className === themeClass);
		}
		for (const className of RESOLVED_EDITOR_THEME_CLASSES) {
			activeDocument.body.classList.toggle(className, className === resolvedThemeClass);
		}
		activeDocument.body.setAttribute('data-native-powerpoint-doc-editor-theme', editorTheme);
		activeDocument.body.setAttribute('data-native-powerpoint-doc-editor-resolved-theme', resolvedTheme);

		const resolvedThemeChanged = this.lastAppliedResolvedEditorTheme !== undefined
			&& this.lastAppliedResolvedEditorTheme !== resolvedTheme;
		this.lastAppliedResolvedEditorTheme = resolvedTheme;
		return resolvedThemeChanged;
	}

	getResolvedEditorTheme(): EditorThemeResolution {
		return this.lastAppliedResolvedEditorTheme ?? this.resolveCurrentEditorTheme();
	}

	private registerEditorThemeObserver() {
		if (typeof MutationObserver === 'undefined') {
			return;
		}
		this.editorThemeObserver = new MutationObserver(() => {
			const previousResolvedTheme = this.lastAppliedResolvedEditorTheme;
			this.applyEditorThemePreference();
			if (
				this.settings.editorTheme === 'system'
				&& previousResolvedTheme !== undefined
				&& previousResolvedTheme !== this.lastAppliedResolvedEditorTheme
			) {
				this.refreshDocxViews();
				this.refreshPowerPointViews();
			}
		});
		this.editorThemeObserver.observe(activeDocument.body, { attributes: true, attributeFilter: ['class'] });
		this.register(() => {
			this.editorThemeObserver?.disconnect();
			this.editorThemeObserver = null;
		});
	}

	private getObsidianThemeResolution(): EditorThemeResolution {
		const bodyClassList = activeDocument.body.classList;
		if (bodyClassList?.contains('theme-dark')) {
			return 'dark';
		}
		return 'light';
	}

	private resolveCurrentEditorTheme(): EditorThemeResolution {
		return resolveEditorThemePreference(
			normalizeEditorThemePreference(this.settings.editorTheme),
			this.getObsidianThemeResolution(),
		);
	}

	private getDebugLogEntries(scope: DebugLogScope) {
		const logs = getNativePowerPointDocEditorLogSnapshot();
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
				locale: this.getResolvedLocale(),
				editorTheme: this.settings.editorTheme,
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
			logStats: getNativePowerPointDocEditorLogStats(),
			logs,
		};

		try {
			await navigator.clipboard.writeText(JSON.stringify(payload, null, 2));
			const label = scope === 'docx' ? 'DOCX' : 'Native PowerPoint Doc Editor';
			showI18nNotice(this.getI18n(), 'settings:debug.logCopied', { count: payload.logs.length, label });
		} catch (error) {
			errorLog('diagnostics', 'Could not copy Native PowerPoint Doc Editor debug log', error);
			showI18nNotice(this.getI18n(), 'settings:debug.logCopyFailed');
		}
	}

	async rebuildDocxSearchIndex(force = false, showNotice = true) {
		if (!this.docxSearchIndex) {
			if (showNotice) {
				showI18nNotice(this.getI18n(), 'docx:notice.searchIndexNotReady');
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
		return getNativePowerPointSettings(
			this.settings,
			this.getResolvedEditorTheme(),
			async (value) => {
				this.settings.powerPointOpenWithYoloMode = value;
				await this.saveSettings();
			},
		);
	}

	private async loadDocxSupport() {
		const docx = await loadDocxSupportModule();
		await docx.registerDocxSupport(this);
	}

	private async loadPowerPointSupport() {
		const pptx = await loadPptxSupportModule();
		pptx.registerPowerPointSupport(this, () => this.getPowerPointSettings());
	}
}
