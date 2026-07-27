import { Notice, Platform, Plugin, normalizePath, setIcon, type WorkspaceLeaf } from 'obsidian';
import {
	NativePowerPointDocEditorSettingTab,
	getNativePowerPointSettings,
	normalizeEditorThemePreference,
	mergeNativePowerPointDocEditorSettings,
	resolveEditorThemePreference,
	type NativePowerPointSettings,
	type NativePowerPointDocEditorSettings,
	type EditorThemeResolution,
} from './settings';
import { DocxSearchIndex } from './docxSearchIndex';
import {
	configureNativePowerPointDocEditorLogger,
	debugLog,
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
import { configurePptxRuntimeArtifactLoader } from './powerpoint/runtimeArtifactLoader';
import { formatDocumentWordCount, type DocumentWordCount } from './documentWordCount';
import {
	AiCore,
	createNpdeAiApi,
	createAiRuntime,
	listOpDefinitions,
	registerAiCommands,
	removeCapabilitiesManifest,
	writeCapabilitiesManifest,
	type NpdeAiApi,
} from './ai';

type DocxSupportModule = typeof import('./docxSupport');
type PptxSupportModule = typeof import('./pptxSupport');

const DOCX_LOG_AREAS = new Set([
	'backup',
	'chunk',
	'clipboard',
	'comments',
	'copy',
	'diagnostics',
	'editor',
	'embed',
	'export',
	'file',
	'font-preservation',
	'history',
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
	'agent',
]);
const PPTX_FILE_EXTENSIONS = ['.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm'];
type DebugLogScope = 'all' | 'docx' | 'pptx';
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
	pluginSettings!: NativePowerPointDocEditorSettings;
	i18n: PluginI18nService | null = null;
	private docxSearchIndex: DocxSearchIndex | null = null;
	private forceJsBackendDevOverride = false;
	private lastAppliedResolvedEditorTheme?: EditorThemeResolution;
	private editorThemeObserver: MutationObserver | null = null;
	private applyingEditorThemePreference = false;
	private aiCore: AiCore | null = null;
	private docxWordCountStatusBarItem: HTMLElement | null = null;
	private readonly documentWordCounts = new Map<WorkspaceLeaf, DocumentWordCount>();
	private activeWordCountLeaf: WorkspaceLeaf | null = null;
	/** Agent API surface. Undefined when AI-Interfacing is disabled in settings. */
	ai: NpdeAiApi | undefined;

	setDocxSearchIndex(index: DocxSearchIndex) {
		this.docxSearchIndex = index;
	}

	updateDocumentWordCount(leaf: WorkspaceLeaf, wordCount: DocumentWordCount) {
		this.documentWordCounts.set(leaf, wordCount);
		this.refreshDocumentWordCountStatus();
	}

	clearDocumentWordCount(leaf: WorkspaceLeaf) {
		this.documentWordCounts.delete(leaf);
		this.refreshDocumentWordCountStatus();
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
		this.initializeDocumentWordCountStatus();
		await initPluginI18n(this, await resolvePluginLocale(this));
		const docxLanguage = this.getResolvedDocxEditorLanguage();
		preloadDocxEditorLocale(docxLanguage);
		void loadDocxEditorLocale(docxLanguage);
		this.applyEditorThemePreference();
		configureNativePowerPointDocEditorLogger(this.pluginSettings.debugLogging);
		// Debug logging is an explicit user setting, not a hot-reload-only feature.
		// Without this, installed development builds retain an old dev-debug.log but
		// never attach its file sink unless a separate .hotreload marker exists.
		if (this.pluginSettings.debugLogging && this.manifest.dir) {
			await this.setupDevFileLog(this.manifest.dir);
		}
		this.aiCore = new AiCore({
			getEnabled: () => this.pluginSettings.enableAiInterfacing,
			pluginVersion: this.manifest.version,
			runtime: createAiRuntime({
				vault: this.app.vault,
				normalizePath: (path) => normalizePath(path),
				findOpenDocxView: (path) => {
					if (!docxSupportModule) return null;
					return docxSupportModule.findDocxViewForPath(this.app, path);
				},
				findOpenPptxView: (path) => {
					if (!pptxSupportModule) return null;
					return pptxSupportModule.findPptxViewForPath(this.app, path);
				},
			}),
		});
		await this.syncAiInterfacing();
		infoLog('plugin', 'Plugin loaded', {
			version: this.manifest.version,
			debugLogging: this.pluginSettings.debugLogging,
			locale: this.getResolvedLocale(),
			editorTheme: this.pluginSettings.editorTheme,
		});
		configureObsidianRuntime({ Notice, Platform, setIcon });
		configureChromiumVersionReader(() => {
			if (!Platform.isDesktop) {
				return null;
			}
			try {
				const electronProcess = (window as unknown as {
					process?: { versions?: { chrome?: string } };
				}).process;
				return electronProcess?.versions?.chrome ?? null;
			} catch {
				return null;
			}
		});
		configureForceJsBackendOverrideReader(() => this.forceJsBackendDevOverride);
		this.configurePowerPointRuntimeArtifactLoader();

		if (!this.pluginSettings.disableDocxFiles) {
			await this.loadDocxSupport();
		} else {
			infoLog('plugin', 'DOCX support disabled by settings');
		}

		if (!this.pluginSettings.disablePowerPointFiles) {
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

		this.registerAiCommands();

		this.addSettingTab(new NativePowerPointDocEditorSettingTab(this.app, this));
		this.registerEditorThemeObserver();
		this.registerEvent(this.app.workspace.on('active-leaf-change', (leaf) => {
			this.activeWordCountLeaf = leaf;
			this.refreshDocumentWordCountStatus();
		}));
		void this.setupDevHotReload();
	}

	private configurePowerPointRuntimeArtifactLoader(): void {
		const pluginDir = this.manifest.dir;
		if (!pluginDir) {
			throw new Error('Could not resolve the Native PowerPoint Doc Editor plugin directory.');
		}

		configurePptxRuntimeArtifactLoader((artifact) => {
			const path = normalizePath(`${pluginDir}/${artifact}`);
			return {
				path,
				resourceUrl: this.app.vault.adapter.getResourcePath(path),
			};
		});
	}

	onunload(): void {
		void this.flushOpenDocumentViewsBeforeUnload().finally(() => {
			this.cleanupAfterUnload();
		});
	}

	private async flushOpenDocumentViewsBeforeUnload(): Promise<void> {
		const flushes: Array<{ kind: 'DOCX' | 'PPTX'; save: () => Promise<boolean> }> = [];
		if (pptxSupportModule) {
			flushes.push({
				kind: 'PPTX',
				save: () => pptxSupportModule!.savePowerPointViewsBeforePluginReload(this),
			});
		}
		if (docxSupportModule) {
			flushes.push({
				kind: 'DOCX',
				save: () => docxSupportModule!.saveDocxViewsBeforePluginReload(this),
			});
		}

		infoLog('plugin', 'Plugin unloading — flushing open document views', {
			formats: flushes.map(({ kind }) => kind),
		});
		const results = await Promise.allSettled(flushes.map(async ({ kind, save }) => ({ kind, ok: await save() })));
		for (const result of results) {
			if (result.status === 'fulfilled') {
				const { kind, ok } = result.value;
				(ok ? infoLog : errorLog)('plugin', `${kind} unload flush settled`, { ok });
			} else {
				errorLog('plugin', 'Unload flush failed', result.reason);
			}
		}
	}

	private cleanupAfterUnload(): void {
		this.editorThemeObserver?.disconnect();
		this.documentWordCounts.clear();
		this.activeWordCountLeaf = null;
		this.docxWordCountStatusBarItem = null;
		this.editorThemeObserver = null;
		const activeDocument = this.app.workspace.containerEl.ownerDocument;
		activeDocument.body.removeClasses([...EDITOR_THEME_CLASSES, ...RESOLVED_EDITOR_THEME_CLASSES]);
		activeDocument.body.removeAttribute('data-native-powerpoint-doc-editor-theme');
		activeDocument.body.removeAttribute('data-native-powerpoint-doc-editor-resolved-theme');
		setNativePowerPointDocEditorLogSink(null);
	}

	private initializeDocumentWordCountStatus() {
		this.activeWordCountLeaf = this.app.workspace.getMostRecentLeaf();
		this.docxWordCountStatusBarItem = this.addStatusBarItem();
		this.docxWordCountStatusBarItem.addClass('native-powerpoint-doc-editor-word-count');
		this.docxWordCountStatusBarItem.setAttribute('aria-live', 'polite');
		this.refreshDocumentWordCountStatus();
	}

	private refreshDocumentWordCountStatus() {
		const statusBarItem = this.docxWordCountStatusBarItem;
		if (!statusBarItem) {
			return;
		}

		const wordCount = this.activeWordCountLeaf
			? this.documentWordCounts.get(this.activeWordCountLeaf)
			: undefined;
		statusBarItem.toggleClass('is-hidden', !wordCount);
		if (wordCount) {
			const text = formatDocumentWordCount(wordCount);
			statusBarItem.setText(text);
			statusBarItem.setAttribute('aria-label', text);
		}
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

		if (!this.pluginSettings.debugLogging) {
			await this.setupDevFileLog(pluginDir);
		}

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
				if (pptxSupportModule && !await pptxSupportModule.savePowerPointViewsBeforePluginReload(this)) {
					reloading = false;
					errorLog('plugin', 'Hot reload aborted because an open PowerPoint file could not be saved', {
						pluginId,
					});
					return;
				}
				if (docxSupportModule && !await docxSupportModule.saveDocxViewsBeforePluginReload(this)) {
					reloading = false;
					errorLog('plugin', 'Hot reload aborted because an open DOCX file could not be saved', {
						pluginId,
					});
					return;
				}

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
		const { settings, shouldPersistSettings } = mergeNativePowerPointDocEditorSettings(
			savedSettings,
			this.getObsidianThemeResolution(),
		);
		this.pluginSettings = settings;

		if (shouldPersistSettings) {
			await this.saveSettings();
		}
	}

	async saveSettings() {
		this.applyEditorThemePreference();
		await this.saveData(this.pluginSettings);
	}

	async syncAiInterfacing(): Promise<void> {
		if (!this.aiCore) {
			this.aiCore = new AiCore({
				getEnabled: () => this.pluginSettings.enableAiInterfacing,
				pluginVersion: this.manifest.version,
				runtime: createAiRuntime({
					vault: this.app.vault,
					normalizePath: (path) => normalizePath(path),
					findOpenDocxView: (path) => {
						if (!docxSupportModule) return null;
						return docxSupportModule.findDocxViewForPath(this.app, path);
					},
					findOpenPptxView: (path) => {
						if (!pptxSupportModule) return null;
						return pptxSupportModule.findPptxViewForPath(this.app, path);
					},
				}),
			});
		}

		if (this.pluginSettings.enableAiInterfacing) {
			this.ai = createNpdeAiApi(this.aiCore);
			const manifestPath = await writeCapabilitiesManifest(
				this.app.vault.adapter,
				this.manifest.dir,
				this.manifest.version,
				true,
			);
			infoLog('agent', 'AI interfacing enabled', {
				manifestPath,
				operationCount: listOpDefinitions().length,
			});
			return;
		}

		this.ai = undefined;
		await removeCapabilitiesManifest(this.app.vault.adapter, this.manifest.dir);
		infoLog('agent', 'AI interfacing disabled');
	}

	private registerAiCommands(): void {
		registerAiCommands({
			plugin: this,
			getI18n: () => this.getI18n(),
			getAi: () => this.ai,
		});
	}

	private applyEditorThemePreference(): boolean {
		if (this.applyingEditorThemePreference) {
			return false;
		}

		this.applyingEditorThemePreference = true;
		try {
			const activeDocument = this.app.workspace.containerEl.ownerDocument;
			const editorTheme = normalizeEditorThemePreference(this.pluginSettings.editorTheme);
			const obsidianTheme = this.getObsidianThemeResolution();
			const resolvedTheme = resolveEditorThemePreference(editorTheme, obsidianTheme);
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

			debugLog('settings', 'Applied editor theme preference', {
				editorTheme,
				obsidianTheme,
				resolvedTheme,
				resolvedThemeChanged,
			});

			return resolvedThemeChanged;
		} finally {
			this.applyingEditorThemePreference = false;
		}
	}

	getResolvedEditorTheme(): EditorThemeResolution {
		// Always live-resolve. Cached lastApplied alone went stale across PPTX↔DOCX
		// leaf switches when system mode re-sampled Obsidian dark/light.
		return this.resolveCurrentEditorTheme();
	}

	private registerEditorThemeObserver() {
		if (typeof MutationObserver === 'undefined') {
			return;
		}
		this.editorThemeObserver = new MutationObserver(() => {
			if (this.applyingEditorThemePreference) {
				return;
			}
			const previousResolvedTheme = this.lastAppliedResolvedEditorTheme;
			this.applyEditorThemePreference();
			if (
				this.pluginSettings.editorTheme === 'system'
				&& previousResolvedTheme !== undefined
				&& previousResolvedTheme !== this.lastAppliedResolvedEditorTheme
			) {
				this.refreshDocxViews();
				this.refreshPowerPointViews();
			}
		});
		const activeDocument = this.app.workspace.containerEl.ownerDocument;
		this.editorThemeObserver.observe(activeDocument.body, { attributes: true, attributeFilter: ['class'] });
		this.editorThemeObserver.observe(activeDocument.documentElement, { attributes: true, attributeFilter: ['class'] });
		this.register(() => {
			this.editorThemeObserver?.disconnect();
			this.editorThemeObserver = null;
		});
	}

	private getObsidianThemeResolution(): EditorThemeResolution {
		const activeDocument = this.app.workspace.containerEl.ownerDocument;
		const bodyClassList = activeDocument.body?.classList;
		const rootClassList = activeDocument.documentElement?.classList;
		if (bodyClassList?.contains('theme-dark') || rootClassList?.contains('theme-dark')) {
			return 'dark';
		}
		if (bodyClassList?.contains('theme-light') || rootClassList?.contains('theme-light')) {
			return 'light';
		}
		// Obsidian "Adapt to system" can leave body without theme-* briefly; prefer OS.
		if (typeof window !== 'undefined' && typeof window.matchMedia === 'function'
			&& window.matchMedia('(prefers-color-scheme: dark)').matches) {
			return 'dark';
		}
		return 'light';
	}

	private resolveCurrentEditorTheme(): EditorThemeResolution {
		return resolveEditorThemePreference(
			normalizeEditorThemePreference(this.pluginSettings.editorTheme),
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

			if (scope === 'pptx') {
				const message = entry.message.toLowerCase();
				return entry.area === 'agent'
					|| message.includes('powerpoint')
					|| message.includes('pptx')
					|| message.includes('ai ')
					|| message.startsWith('ai ')
					|| PPTX_FILE_EXTENSIONS.some((extension) => serializedData.includes(extension));
			}

			return entry.area === 'agent'
				|| DOCX_LOG_AREAS.has(entry.area)
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
				editorTheme: this.pluginSettings.editorTheme,
				showRuler: this.pluginSettings.showRuler,
				autosave: this.pluginSettings.autosave,
				createBackupsBeforeSave: this.pluginSettings.createBackupsBeforeSave,
				defaultZoom: this.pluginSettings.defaultZoom,
				debugLogging: this.pluginSettings.debugLogging,
				enableDocxSearchIndex: this.pluginSettings.enableDocxSearchIndex,
				autoIndexDocxSearch: this.pluginSettings.autoIndexDocxSearch,
				powerPointAutosaveEnabled: this.pluginSettings.powerPointAutosaveEnabled,
				powerPointHideUnsupportedSvgContent: this.pluginSettings.powerPointHideUnsupportedSvgContent,
				powerPointOpenWithYoloMode: this.pluginSettings.powerPointOpenWithYoloMode,
				disableDocxFiles: this.pluginSettings.disableDocxFiles,
				disablePowerPointFiles: this.pluginSettings.disablePowerPointFiles,
				enableAiInterfacing: this.pluginSettings.enableAiInterfacing,
			},
			docxEditorBundle: 'main.js',
			logStats: getNativePowerPointDocEditorLogStats(),
			logs,
		};

		try {
			await navigator.clipboard.writeText(JSON.stringify(payload, null, 2));
			const label = scope === 'docx' ? 'DOCX' : scope === 'pptx' ? 'PPTX' : 'Native PowerPoint Doc Editor';
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
			this.pluginSettings,
			this.getResolvedEditorTheme(),
			async (value) => {
				this.pluginSettings.powerPointOpenWithYoloMode = value;
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
