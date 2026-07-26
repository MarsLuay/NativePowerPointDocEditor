import { App, PluginSettingTab, Setting, type SettingDefinitionItem } from 'obsidian';
import type NativePowerPointDocEditorPlugin from './main';
import { configureNativePowerPointDocEditorLogger, infoLog } from './logger';
import type { I18nService } from './i18n/I18nService';
import {
	getNativePowerPointDocEditorSettingDescriptors,
	getNativePowerPointDocEditorSettingsTabSections,
	getEditorThemeSettingOptions as getEditorThemeSettingOptionsI18n,
} from './i18n/settingsCatalog';
import { showI18nNotice } from './i18n/notify';

export type {
	DocxEditorSettingDescriptor,
	NativePowerPointDocEditorSettingDescriptor,
} from './i18n/settingsCatalog';
export {
	getDocxEditorSettingDescriptors,
	getDocxEditorSettingSectionLabels,
	getDocxEditorSettingsMenuSections,
} from './i18n/docxSettingsCatalog';
export {
	getNativePowerPointDocEditorSettingDescriptors,
	getNativePowerPointDocEditorSettingSectionLabels,
	getNativePowerPointDocEditorSettingsTabSections,
} from './i18n/settingsCatalog';

export const DEFAULT_ZOOM = 1;
export const MIN_DEFAULT_ZOOM = 0.5;
export const MAX_DEFAULT_ZOOM = 2;
export const DEFAULT_ZOOM_STEP = 0.05;
export const DEFAULT_ZOOM_OPTIONS = [0.5, 0.75, 1, 1.25, 1.5, 2] as const;

export type EditorThemePreference = 'system' | 'light' | 'dark';
export type EditorThemeResolution = Exclude<EditorThemePreference, 'system'>;

const ENGLISH_EDITOR_THEME_OPTIONS: Array<{ value: EditorThemePreference; label: string }> = [
	{ value: 'system', label: 'System' },
	{ value: 'light', label: 'Light' },
	{ value: 'dark', label: 'Dark' },
];

const FREE_TEMPLATE_LINKS = [
	{ name: 'Microsoft Create', url: 'https://create.microsoft.com/en-us/templates/presentations' },
	{ name: 'Slidesgo', url: 'https://slidesgo.com/' },
	{ name: 'SlidesCarnival', url: 'https://www.slidescarnival.com/' },
	{ name: 'SlidesMania', url: 'https://slidesmania.com/free-templates/presentation-templates/' },
	{ name: 'Canva templates', url: 'https://www.canva.com/presentations/templates/slides/' },
];

export interface NativePowerPointSettings {
	editorTheme: EditorThemePreference;
	resolvedEditorTheme: EditorThemeResolution;
	autosaveEnabled: boolean;
	hideUnsupportedSvgContent: boolean;
	openWithYoloMode: boolean;
	showInspector: boolean;
	setOpenWithYoloMode: (value: boolean) => Promise<void>;
}

export interface NativePowerPointDocEditorSettings {
	authorName: string;
	editorTheme: EditorThemePreference;
	showRuler: boolean;
	autosave: boolean;
	createBackupsBeforeSave: boolean;
	defaultZoom: number;
	enableDocxSearchIndex: boolean;
	autoIndexDocxSearch: boolean;
	debugLogging: boolean;
	powerPointAutosaveEnabled: boolean;
	powerPointHideUnsupportedSvgContent: boolean;
	powerPointOpenWithYoloMode: boolean;
	powerPointShowInspector: boolean;
	disableDocxFiles: boolean;
	disablePowerPointFiles: boolean;
	enableAiInterfacing: boolean;
}

export const DEFAULT_SETTINGS: NativePowerPointDocEditorSettings = {
	authorName: 'Obsidian',
	editorTheme: 'system',
	showRuler: false,
	autosave: true,
	createBackupsBeforeSave: false,
	defaultZoom: DEFAULT_ZOOM,
	enableDocxSearchIndex: false,
	autoIndexDocxSearch: false,
	debugLogging: false,
	powerPointAutosaveEnabled: true,
	powerPointHideUnsupportedSvgContent: false,
	powerPointOpenWithYoloMode: false,
	powerPointShowInspector: false,
	disableDocxFiles: false,
	disablePowerPointFiles: false,
	enableAiInterfacing: false,
};

export type NativePowerPointDocEditorSettingSectionId =
	| 'identity'
	| 'fileHandoff'
	| 'editorDefaults'
	| 'saving'
	| 'powerpoint'
	| 'search'
	| 'ai'
	| 'diagnostics';

export type DocxEditorSettingSectionId = Exclude<NativePowerPointDocEditorSettingSectionId, 'powerpoint' | 'ai'>;

export type NativePowerPointDocEditorSettingId =
	| 'authorName'
	| 'disableDocxFiles'
	| 'disablePowerPointFiles'
	| 'editorTheme'
	| 'showRuler'
	| 'defaultZoom'
	| 'autosave'
	| 'createBackupsBeforeSave'
	| 'powerPointAutosaveEnabled'
	| 'powerPointShowInspector'
	| 'powerPointHideUnsupportedSvgContent'
	| 'powerPointOpenWithYoloMode'
	| 'enableDocxSearchIndex'
	| 'autoIndexDocxSearch'
	| 'rebuildDocxSearchIndex'
	| 'enableAiInterfacing'
	| 'debugLogging'
	| 'copyDocxLog'
	| 'copyPptxLog'
	| 'copyFullLog';

export type DocxEditorSettingId = Extract<
	NativePowerPointDocEditorSettingId,
	| 'authorName'
	| 'editorTheme'
	| 'showRuler'
	| 'defaultZoom'
	| 'autosave'
	| 'createBackupsBeforeSave'
	| 'enableDocxSearchIndex'
	| 'autoIndexDocxSearch'
	| 'rebuildDocxSearchIndex'
	| 'disableDocxFiles'
	| 'debugLogging'
	| 'copyDocxLog'
>;

export interface SettingsOption<TValue extends string = string> {
	value: TValue;
	label: string;
}

function withDefaultLabel<TValue extends string>(
	options: ReadonlyArray<{ value: TValue; label: string }>,
	defaultValue: TValue,
	defaultSuffix = '(default)',
): SettingsOption<TValue>[] {
	return options.map(option => ({
		value: option.value,
		label: option.value === defaultValue ? `${option.label} ${defaultSuffix}` : option.label,
	}));
}

export function getEditorThemeSettingOptions(
	i18n?: I18nService,
	defaultTheme: EditorThemePreference = DEFAULT_SETTINGS.editorTheme,
): SettingsOption<EditorThemePreference>[] {
	const theme = normalizeEditorThemePreference(defaultTheme);
	if (i18n) {
		return getEditorThemeSettingOptionsI18n(i18n, theme);
	}
	return withDefaultLabel(ENGLISH_EDITOR_THEME_OPTIONS, theme);
}

function readString(value: unknown, fallback: string): string {
	return typeof value === 'string' && value.trim().length > 0 ? value.trim() : fallback;
}

function hasSavedSetting(
	saved: Record<string, unknown> | null | undefined,
	key: keyof NativePowerPointDocEditorSettings,
): boolean {
	return saved != null && Object.prototype.hasOwnProperty.call(saved, key);
}

export interface ReadNativePowerPointDocEditorSettingsResult {
	settings: NativePowerPointDocEditorSettings;
	hadLegacyEditorLanguage: boolean;
	shouldPersistSettings: boolean;
}

export function mergeNativePowerPointDocEditorSettings(
	saved: Record<string, unknown> | null | undefined,
	systemTheme: EditorThemeResolution = 'light',
): ReadNativePowerPointDocEditorSettingsResult {
	const raw = saved ?? {};
	const hadLegacyEditorLanguage = Object.hasOwn(raw, 'editorLanguage');

	const normalizedEditorTheme = hasSavedSetting(raw, 'editorTheme')
		? normalizeEditorThemePreference(raw.editorTheme)
		: saved
			? systemTheme
			: DEFAULT_SETTINGS.editorTheme;
	const normalizedDefaultZoom = normalizeDefaultZoom(raw.defaultZoom);
	const normalizedDebugLogging = raw.debugLogging === true;
	const normalizedEnableDocxSearchIndex = raw.enableDocxSearchIndex === true;
	const normalizedAutoIndexDocxSearch = raw.autoIndexDocxSearch === true;
	const normalizedPowerPointAutosaveEnabled = raw.powerPointAutosaveEnabled !== false;
	const normalizedPowerPointHideUnsupportedSvgContent =
		typeof raw.powerPointHideUnsupportedSvgContent === 'boolean'
			? raw.powerPointHideUnsupportedSvgContent
			: raw.powerPointRemoveUnsupportedSvgContent === true;
	const normalizedPowerPointOpenWithYoloMode =
		typeof raw.powerPointOpenWithYoloMode === 'boolean'
			? raw.powerPointOpenWithYoloMode
			: raw.powerPointYoloMode === true;
	const normalizedPowerPointShowInspector = raw.powerPointShowInspector === true;
	const normalizedDisableDocxFiles = raw.disableDocxFiles === true;
	const normalizedDisablePowerPointFiles = raw.disablePowerPointFiles === true;
	const normalizedEnableAiInterfacing = raw.enableAiInterfacing === true;

	const settings: NativePowerPointDocEditorSettings = {
		authorName: readString(raw.authorName, DEFAULT_SETTINGS.authorName),
		editorTheme: normalizedEditorTheme,
		showRuler: raw.showRuler === true,
		autosave: raw.autosave !== false,
		createBackupsBeforeSave: raw.createBackupsBeforeSave === true,
		defaultZoom: normalizedDefaultZoom,
		enableDocxSearchIndex: normalizedEnableDocxSearchIndex,
		autoIndexDocxSearch: normalizedAutoIndexDocxSearch,
		debugLogging: normalizedDebugLogging,
		powerPointAutosaveEnabled: normalizedPowerPointAutosaveEnabled,
		powerPointHideUnsupportedSvgContent: normalizedPowerPointHideUnsupportedSvgContent,
		powerPointOpenWithYoloMode: normalizedPowerPointOpenWithYoloMode,
		powerPointShowInspector: normalizedPowerPointShowInspector,
		disableDocxFiles: normalizedDisableDocxFiles,
		disablePowerPointFiles: normalizedDisablePowerPointFiles,
		enableAiInterfacing: normalizedEnableAiInterfacing,
	};

	const shouldPersistSettings = hadLegacyEditorLanguage
		|| raw.editorTheme !== normalizedEditorTheme
		|| raw.defaultZoom !== normalizedDefaultZoom
		|| raw.debugLogging !== normalizedDebugLogging
		|| raw.enableDocxSearchIndex !== normalizedEnableDocxSearchIndex
		|| raw.autoIndexDocxSearch !== normalizedAutoIndexDocxSearch
		|| raw.powerPointAutosaveEnabled !== normalizedPowerPointAutosaveEnabled
		|| raw.powerPointHideUnsupportedSvgContent !== normalizedPowerPointHideUnsupportedSvgContent
		|| raw.powerPointOpenWithYoloMode !== normalizedPowerPointOpenWithYoloMode
		|| raw.powerPointShowInspector !== normalizedPowerPointShowInspector
		|| raw.powerPointRemoveUnsupportedSvgContent !== undefined
		|| raw.powerPointYoloMode !== undefined
		|| raw.disableDocxFiles !== normalizedDisableDocxFiles
		|| raw.disablePowerPointFiles !== normalizedDisablePowerPointFiles
		|| raw.enableAiInterfacing !== normalizedEnableAiInterfacing;

	return {
		settings,
		hadLegacyEditorLanguage,
		shouldPersistSettings,
	};
}

export function getDefaultZoomSettingOptions(): SettingsOption[] {
	return DEFAULT_ZOOM_OPTIONS.map(zoom => ({
		value: String(zoom),
		label: formatZoom(zoom),
	}));
}

export function getNativePowerPointSettings(
	settings: NativePowerPointDocEditorSettings,
	resolvedEditorThemeOrSetOpenWithYoloMode: EditorThemeResolution | ((value: boolean) => Promise<void>) = resolveEditorThemePreference(settings.editorTheme),
	setOpenWithYoloMode: (value: boolean) => Promise<void> = async () => {}
): NativePowerPointSettings {
	const resolvedEditorTheme = typeof resolvedEditorThemeOrSetOpenWithYoloMode === 'function'
		? resolveEditorThemePreference(settings.editorTheme)
		: resolvedEditorThemeOrSetOpenWithYoloMode;
	const setOpenWithYoloModeHandler = typeof resolvedEditorThemeOrSetOpenWithYoloMode === 'function'
		? resolvedEditorThemeOrSetOpenWithYoloMode
		: setOpenWithYoloMode;

	return {
		editorTheme: settings.editorTheme,
		resolvedEditorTheme,
		autosaveEnabled: settings.powerPointAutosaveEnabled,
		hideUnsupportedSvgContent: settings.powerPointHideUnsupportedSvgContent,
		openWithYoloMode: settings.powerPointOpenWithYoloMode,
		showInspector: settings.powerPointShowInspector,
		setOpenWithYoloMode: setOpenWithYoloModeHandler,
	};
}

export function normalizeDefaultZoom(value: unknown): number {
	const numericValue = typeof value === 'number' ? value : Number(value);
	if (!Number.isFinite(numericValue)) {
		return DEFAULT_ZOOM;
	}

	const clampedValue = Math.min(MAX_DEFAULT_ZOOM, Math.max(MIN_DEFAULT_ZOOM, numericValue));
	return Math.round(clampedValue / DEFAULT_ZOOM_STEP) * DEFAULT_ZOOM_STEP;
}

export function normalizeEditorThemePreference(value: unknown): EditorThemePreference {
	return value === 'light' || value === 'dark' || value === 'system' ? value : DEFAULT_SETTINGS.editorTheme;
}

export function resolveEditorThemePreference(
	value: unknown,
	systemTheme: EditorThemeResolution = 'light',
): EditorThemeResolution {
	const normalizedTheme = normalizeEditorThemePreference(value);
	if (normalizedTheme === 'light' || normalizedTheme === 'dark') {
		return normalizedTheme;
	}
	return systemTheme === 'dark' ? 'dark' : 'light';
}

export function formatZoom(value: number): string {
	return `${Math.round(value * 100)}%`;
}

export class NativePowerPointDocEditorSettingTab extends PluginSettingTab {
	plugin: NativePowerPointDocEditorPlugin;

	constructor(app: App, plugin: NativePowerPointDocEditorPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		this.renderSettings();
	}

	/**
	 * Obsidian 1.13+: indexes settings for global search and renders this tree
	 * (skips {@link display}). Pre-1.13 keeps imperative {@link display}.
	 */
	getSettingDefinitions(): SettingDefinitionItem[] {
		const i18n = this.plugin.getI18n()!;
		const descriptors = getNativePowerPointDocEditorSettingDescriptors(i18n);
		const sectionLabels = Object.fromEntries(
			getNativePowerPointDocEditorSettingsTabSections(i18n).map(({ id, label }) => [id, label]),
		) as Record<NativePowerPointDocEditorSettingSectionId, string>;
		const defaultAuthorName = String(descriptors.authorName.defaultValue ?? DEFAULT_SETTINGS.authorName);
		const defaultTheme = normalizeEditorThemePreference(descriptors.editorTheme.defaultValue);
		const defaultZoom = normalizeDefaultZoom(descriptors.defaultZoom.defaultValue);
		const themeOptions = Object.fromEntries(
			getEditorThemeSettingOptions(i18n, defaultTheme).map((option) => [option.value, option.label]),
		);

		return [
			{
				type: 'group',
				heading: sectionLabels.identity,
				items: [
					{
						name: descriptors.authorName.name,
						desc: descriptors.authorName.description,
						render: (setting) => {
							setting
								.addText((text) =>
									text
										.setPlaceholder(descriptors.authorName.placeholder ?? defaultAuthorName)
										.setValue(this.plugin.pluginSettings.authorName)
										.onChange(async (value) => {
											this.plugin.pluginSettings.authorName = value.trim() || defaultAuthorName;
											await this.plugin.saveSettings();
										}),
								)
								.addButton((button) =>
									button
										.setButtonText(descriptors.authorName.resetLabel ?? i18n.t('common:actions.reset'))
										.onClick(async () => {
											this.plugin.pluginSettings.authorName = defaultAuthorName;
											await this.plugin.saveSettings();
											this.refreshSettingsUi();
										}),
								);
						},
					},
				],
			},
			{
				type: 'group',
				heading: i18n.t('settings:fileHandoff.section'),
				items: [
					{
						name: descriptors.disableDocxFiles.name,
						desc: descriptors.disableDocxFiles.description,
						control: { type: 'toggle', key: 'disableDocxFiles' },
					},
					{
						name: descriptors.disablePowerPointFiles.name,
						desc: descriptors.disablePowerPointFiles.description,
						control: { type: 'toggle', key: 'disablePowerPointFiles' },
					},
				],
			},
			{
				type: 'group',
				heading: sectionLabels.editorDefaults,
				items: [
					{
						name: descriptors.editorTheme.name,
						desc: descriptors.editorTheme.description,
						control: {
							type: 'dropdown',
							key: 'editorTheme',
							defaultValue: defaultTheme,
							options: themeOptions,
						},
					},
					{
						name: descriptors.showRuler.name,
						desc: descriptors.showRuler.description,
						control: { type: 'toggle', key: 'showRuler' },
					},
					{
						name: descriptors.defaultZoom.name,
						desc: descriptors.defaultZoom.description,
						render: (setting) => {
							const selectedZoom = normalizeDefaultZoom(this.plugin.pluginSettings.defaultZoom);
							const zoomValueEl = setting.controlEl.createSpan({
								cls: 'native-powerpoint-doc-editor-setting-value',
								text: formatZoom(selectedZoom),
							});
							setting
								.addSlider((slider) =>
									slider
										.setLimits(MIN_DEFAULT_ZOOM, MAX_DEFAULT_ZOOM, DEFAULT_ZOOM_STEP)
										.setValue(selectedZoom)
										.onChange(async (value) => {
											const zoom = normalizeDefaultZoom(value);
											this.plugin.pluginSettings.defaultZoom = zoom;
											zoomValueEl.setText(formatZoom(zoom));
											await this.plugin.saveSettings();
										}),
								)
								.addButton((button) =>
									button
										.setButtonText(descriptors.defaultZoom.resetLabel ?? i18n.t('common:actions.reset'))
										.onClick(async () => {
											this.plugin.pluginSettings.defaultZoom = defaultZoom;
											await this.plugin.saveSettings();
											this.refreshSettingsUi();
										}),
								);
						},
					},
				],
			},
			{
				type: 'group',
				heading: sectionLabels.saving,
				items: [
					{
						name: descriptors.autosave.name,
						desc: descriptors.autosave.description,
						control: { type: 'toggle', key: 'autosave' },
					},
					{
						name: descriptors.createBackupsBeforeSave.name,
						desc: descriptors.createBackupsBeforeSave.description,
						control: { type: 'toggle', key: 'createBackupsBeforeSave' },
					},
					{
						name: descriptors.powerPointAutosaveEnabled.name,
						desc: descriptors.powerPointAutosaveEnabled.description,
						control: { type: 'toggle', key: 'powerPointAutosaveEnabled' },
					},
				],
			},
			{
				type: 'group',
				heading: sectionLabels.powerpoint,
				items: [
					{
						name: descriptors.powerPointShowInspector.name,
						desc: descriptors.powerPointShowInspector.description,
						control: { type: 'toggle', key: 'powerPointShowInspector' },
					},
					{
						name: descriptors.powerPointHideUnsupportedSvgContent.name,
						desc: descriptors.powerPointHideUnsupportedSvgContent.description,
						control: { type: 'toggle', key: 'powerPointHideUnsupportedSvgContent' },
					},
					{
						name: descriptors.powerPointOpenWithYoloMode.name,
						desc: descriptors.powerPointOpenWithYoloMode.description,
						control: { type: 'toggle', key: 'powerPointOpenWithYoloMode' },
					},
					{
						name: i18n.t('settings:templates.title'),
						desc: i18n.t('settings:templates.description'),
						render: (setting) => {
							const templateLinks = setting.controlEl.createDiv({ cls: 'native-powerpoint-template-links' });
							for (const link of FREE_TEMPLATE_LINKS) {
								templateLinks.createEl('a', {
									cls: 'native-powerpoint-template-link',
									text: link.name,
									attr: {
										href: link.url,
										rel: 'noopener noreferrer',
										target: '_blank',
									},
								});
							}
						},
					},
				],
			},
			{
				type: 'group',
				heading: sectionLabels.search,
				items: [
					{
						name: descriptors.enableDocxSearchIndex.name,
						desc: descriptors.enableDocxSearchIndex.description,
						control: { type: 'toggle', key: 'enableDocxSearchIndex' },
					},
					{
						name: descriptors.autoIndexDocxSearch.name,
						desc: descriptors.autoIndexDocxSearch.description,
						control: { type: 'toggle', key: 'autoIndexDocxSearch' },
					},
					{
						name: descriptors.rebuildDocxSearchIndex.name,
						desc: descriptors.rebuildDocxSearchIndex.description,
						render: (setting) => {
							setting.addButton((button) =>
								button
									.setButtonText(
										descriptors.rebuildDocxSearchIndex.actionLabel ?? i18n.t('common:actions.rebuild'),
									)
									.onClick(async () => {
										await this.plugin.rebuildDocxSearchIndex(true);
									}),
							);
						},
					},
				],
			},
			{
				type: 'group',
				heading: sectionLabels.ai,
				items: [
					{
						name: descriptors.enableAiInterfacing.name,
						desc: descriptors.enableAiInterfacing.description,
						control: { type: 'toggle', key: 'enableAiInterfacing' },
					},
				],
			},
			{
				type: 'group',
				heading: sectionLabels.diagnostics,
				items: [
					{
						name: descriptors.debugLogging.name,
						desc: descriptors.debugLogging.description,
						control: { type: 'toggle', key: 'debugLogging' },
					},
					{
						name: descriptors.copyDocxLog.name,
						desc: descriptors.copyDocxLog.description,
						render: (setting) => {
							setting.addButton((button) =>
								button
									.setButtonText(descriptors.copyDocxLog.actionLabel ?? i18n.t('common:actions.copy'))
									.onClick(async () => {
										await this.plugin.copyDebugLog('docx');
									}),
							);
						},
					},
					{
						name: descriptors.copyPptxLog.name,
						desc: descriptors.copyPptxLog.description,
						render: (setting) => {
							setting.addButton((button) =>
								button
									.setButtonText(descriptors.copyPptxLog.actionLabel ?? i18n.t('common:actions.copy'))
									.onClick(async () => {
										await this.plugin.copyDebugLog('pptx');
									}),
							);
						},
					},
					{
						name: descriptors.copyFullLog.name,
						desc: descriptors.copyFullLog.description,
						render: (setting) => {
							setting.addButton((button) =>
								button
									.setButtonText(descriptors.copyFullLog.actionLabel ?? i18n.t('common:actions.copy'))
									.onClick(async () => {
										await this.plugin.copyDebugLog('all');
									}),
							);
						},
					},
					{
						name: 'Report bug',
						render: (setting) => {
							const reportBugBox = setting.controlEl.createDiv({ cls: 'native-powerpoint-report-bug' });
							reportBugBox.createEl('a', {
								cls: 'native-powerpoint-report-bug-link',
								text: 'Report bug',
								attr: {
									href: 'https://github.com/MarsLuay/NativePowerPointDocEditor/issues',
									rel: 'noopener noreferrer',
									target: '_blank',
								},
							});
							reportBugBox.createEl('a', {
								cls: 'native-powerpoint-report-bug-link',
								text: 'Buy me a coffee',
								attr: {
									href: 'https://buymeacoffee.com/marwanluaye',
									rel: 'noopener noreferrer',
									target: '_blank',
								},
							});
						},
					},
				],
			},
		];
	}

	getControlValue(key: string): unknown {
		return this.plugin.pluginSettings[key as keyof NativePowerPointDocEditorSettings];
	}

	async setControlValue(key: string, value: unknown): Promise<void> {
		const settings = this.plugin.pluginSettings;
		const i18n = this.plugin.getI18n()!;

		switch (key as keyof NativePowerPointDocEditorSettings) {
			case 'authorName':
				settings.authorName = String(value).trim() || DEFAULT_SETTINGS.authorName;
				break;
			case 'disableDocxFiles':
				settings.disableDocxFiles = Boolean(value);
				await this.plugin.saveSettings();
				showI18nNotice(i18n, 'settings:fileHandoff.reloadDocxNotice');
				return;
			case 'disablePowerPointFiles':
				settings.disablePowerPointFiles = Boolean(value);
				await this.plugin.saveSettings();
				showI18nNotice(i18n, 'settings:fileHandoff.reloadPptxNotice');
				return;
			case 'editorTheme':
				settings.editorTheme = normalizeEditorThemePreference(value);
				await this.plugin.saveSettings();
				this.plugin.refreshDocxViews();
				this.plugin.refreshPowerPointViews();
				return;
			case 'showRuler':
				settings.showRuler = Boolean(value);
				await this.plugin.saveSettings();
				this.plugin.refreshDocxViews();
				return;
			case 'defaultZoom':
				settings.defaultZoom = normalizeDefaultZoom(value);
				break;
			case 'autosave':
				settings.autosave = Boolean(value);
				await this.plugin.saveSettings();
				this.plugin.refreshDocxViews();
				return;
			case 'createBackupsBeforeSave':
				settings.createBackupsBeforeSave = Boolean(value);
				await this.plugin.saveSettings();
				this.plugin.refreshDocxViews();
				return;
			case 'powerPointAutosaveEnabled':
				settings.powerPointAutosaveEnabled = Boolean(value);
				break;
			case 'powerPointShowInspector':
				settings.powerPointShowInspector = Boolean(value);
				await this.plugin.saveSettings();
				this.plugin.refreshPowerPointViews();
				return;
			case 'powerPointHideUnsupportedSvgContent':
				settings.powerPointHideUnsupportedSvgContent = Boolean(value);
				break;
			case 'powerPointOpenWithYoloMode':
				settings.powerPointOpenWithYoloMode = Boolean(value);
				break;
			case 'enableDocxSearchIndex':
				settings.enableDocxSearchIndex = Boolean(value);
				await this.plugin.saveSettings();
				if (settings.enableDocxSearchIndex) {
					await this.plugin.rebuildDocxSearchIndex(false);
				}
				return;
			case 'autoIndexDocxSearch':
				settings.autoIndexDocxSearch = Boolean(value);
				await this.plugin.saveSettings();
				if (settings.autoIndexDocxSearch && settings.enableDocxSearchIndex) {
					await this.plugin.rebuildDocxSearchIndex(false);
				}
				return;
			case 'enableAiInterfacing':
				settings.enableAiInterfacing = Boolean(value);
				await this.plugin.saveSettings();
				await this.plugin.syncAiInterfacing();
				return;
			case 'debugLogging':
				settings.debugLogging = Boolean(value);
				configureNativePowerPointDocEditorLogger(settings.debugLogging);
				infoLog('settings', `Debug logging ${settings.debugLogging ? 'enabled' : 'disabled'}`);
				break;
			default:
				return;
		}

		await this.plugin.saveSettings();
	}

	private refreshSettingsUi(): void {
		// Always re-render via display path. Do not call PluginSettingTab.update —
		// that API is newer than minAppVersion 1.8.7 and fails obsidianmd/no-unsupported-api.
		this.renderSettings();
	}

	private renderSettings(): void {
		const { containerEl } = this;
		const i18n = this.plugin.getI18n()!;
		const selectedZoom = normalizeDefaultZoom(this.plugin.pluginSettings.defaultZoom);

		this.plugin.pluginSettings.defaultZoom = selectedZoom;

		containerEl.empty();
		containerEl.addClass('native-powerpoint-doc-editor-settings-tab');
		const settingDescriptors = getNativePowerPointDocEditorSettingDescriptors(i18n);
		const sectionLabels = Object.fromEntries(
			getNativePowerPointDocEditorSettingsTabSections(i18n).map(({ id, label }) => [id, label]),
		) as Record<NativePowerPointDocEditorSettingSectionId, string>;
		const defaultAuthorName = String(settingDescriptors.authorName.defaultValue ?? DEFAULT_SETTINGS.authorName);
		const defaultTheme = normalizeEditorThemePreference(settingDescriptors.editorTheme.defaultValue);
		const defaultZoom = normalizeDefaultZoom(settingDescriptors.defaultZoom.defaultValue);

		new Setting(containerEl)
			.setName(sectionLabels.identity)
			.setHeading();

		new Setting(containerEl)
			.setName(settingDescriptors.authorName.name)
			.setDesc(settingDescriptors.authorName.description)
			.addText(text => text
				.setPlaceholder(settingDescriptors.authorName.placeholder ?? defaultAuthorName)
				.setValue(this.plugin.pluginSettings.authorName)
				.onChange(async (value) => {
					this.plugin.pluginSettings.authorName = value.trim() || defaultAuthorName;
					await this.plugin.saveSettings();
				}))
			.addButton(button => button
				.setButtonText(settingDescriptors.authorName.resetLabel ?? i18n.t('common:actions.reset'))
				.onClick(async () => {
					this.plugin.pluginSettings.authorName = defaultAuthorName;
					await this.plugin.saveSettings();
					this.renderSettings();
				}));

		new Setting(containerEl)
			.setName(i18n.t('settings:fileHandoff.section'))
			.setHeading();

		new Setting(containerEl)
			.setName(settingDescriptors.disableDocxFiles.name)
			.setDesc(settingDescriptors.disableDocxFiles.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.disableDocxFiles)
				.onChange(async (value) => {
					this.plugin.pluginSettings.disableDocxFiles = value;
					await this.plugin.saveSettings();
					showI18nNotice(i18n, 'settings:fileHandoff.reloadDocxNotice');
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.disablePowerPointFiles.name)
			.setDesc(settingDescriptors.disablePowerPointFiles.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.disablePowerPointFiles)
				.onChange(async (value) => {
					this.plugin.pluginSettings.disablePowerPointFiles = value;
					await this.plugin.saveSettings();
					showI18nNotice(i18n, 'settings:fileHandoff.reloadPptxNotice');
				}));

		new Setting(containerEl)
			.setName(sectionLabels.editorDefaults)
			.setHeading();

		new Setting(containerEl)
			.setName(settingDescriptors.editorTheme.name)
			.setDesc(settingDescriptors.editorTheme.description)
			.addDropdown(dropdown => {
				for (const option of getEditorThemeSettingOptions(i18n, defaultTheme)) {
					dropdown.addOption(option.value, option.label);
				}

				dropdown
					.setValue(this.plugin.pluginSettings.editorTheme)
					.onChange(async (value) => {
						this.plugin.pluginSettings.editorTheme = normalizeEditorThemePreference(value);
						await this.plugin.saveSettings();
						this.plugin.refreshDocxViews();
						this.plugin.refreshPowerPointViews();
					});
			});

		new Setting(containerEl)
			.setName(settingDescriptors.showRuler.name)
			.setDesc(settingDescriptors.showRuler.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.showRuler)
				.onChange(async (value) => {
					this.plugin.pluginSettings.showRuler = value;
					await this.plugin.saveSettings();
					this.plugin.refreshDocxViews();
				}));

		const zoomSetting = new Setting(containerEl)
			.setName(settingDescriptors.defaultZoom.name)
			.setDesc(settingDescriptors.defaultZoom.description);
		const zoomValueEl = zoomSetting.controlEl.createSpan({
			cls: 'native-powerpoint-doc-editor-setting-value',
			text: formatZoom(selectedZoom),
		});

		zoomSetting
			.addSlider(slider => slider
				.setLimits(MIN_DEFAULT_ZOOM, MAX_DEFAULT_ZOOM, DEFAULT_ZOOM_STEP)
				.setValue(selectedZoom)
				.onChange(async (value) => {
					const zoom = normalizeDefaultZoom(value);
					this.plugin.pluginSettings.defaultZoom = zoom;
					zoomValueEl.setText(formatZoom(zoom));
					await this.plugin.saveSettings();
			}))
			.addButton(button => button
				.setButtonText(settingDescriptors.defaultZoom.resetLabel ?? i18n.t('common:actions.reset'))
				.onClick(async () => {
					this.plugin.pluginSettings.defaultZoom = defaultZoom;
					await this.plugin.saveSettings();
					this.renderSettings();
				}));

		new Setting(containerEl)
			.setName(sectionLabels.saving)
			.setHeading();

		new Setting(containerEl)
			.setName(settingDescriptors.autosave.name)
			.setDesc(settingDescriptors.autosave.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.autosave)
				.onChange(async (value) => {
					this.plugin.pluginSettings.autosave = value;
					await this.plugin.saveSettings();
					this.plugin.refreshDocxViews();
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.createBackupsBeforeSave.name)
			.setDesc(settingDescriptors.createBackupsBeforeSave.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.createBackupsBeforeSave)
				.onChange(async (value) => {
					this.plugin.pluginSettings.createBackupsBeforeSave = value;
					await this.plugin.saveSettings();
					this.plugin.refreshDocxViews();
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.powerPointAutosaveEnabled.name)
			.setDesc(settingDescriptors.powerPointAutosaveEnabled.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.powerPointAutosaveEnabled)
				.onChange(async (value) => {
					this.plugin.pluginSettings.powerPointAutosaveEnabled = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName(sectionLabels.powerpoint)
			.setHeading();

		new Setting(containerEl)
			.setName(settingDescriptors.powerPointShowInspector.name)
			.setDesc(settingDescriptors.powerPointShowInspector.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.powerPointShowInspector)
				.onChange(async (value) => {
					this.plugin.pluginSettings.powerPointShowInspector = value;
					await this.plugin.saveSettings();
					this.plugin.refreshPowerPointViews();
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.powerPointHideUnsupportedSvgContent.name)
			.setDesc(settingDescriptors.powerPointHideUnsupportedSvgContent.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.powerPointHideUnsupportedSvgContent)
				.onChange(async (value) => {
					this.plugin.pluginSettings.powerPointHideUnsupportedSvgContent = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.powerPointOpenWithYoloMode.name)
			.setDesc(settingDescriptors.powerPointOpenWithYoloMode.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.powerPointOpenWithYoloMode)
				.onChange(async (value) => {
					this.plugin.pluginSettings.powerPointOpenWithYoloMode = value;
					await this.plugin.saveSettings();
				}));

		const templateBox = containerEl.createDiv({ cls: 'native-powerpoint-template-box' });
		templateBox.createDiv({ cls: 'native-powerpoint-template-title', text: i18n.t('settings:templates.title') });
		templateBox.createDiv({
			cls: 'native-powerpoint-template-desc',
			text: i18n.t('settings:templates.description'),
		});

		const templateLinks = templateBox.createDiv({ cls: 'native-powerpoint-template-links' });
		for (const link of FREE_TEMPLATE_LINKS) {
			templateLinks.createEl('a', {
				cls: 'native-powerpoint-template-link',
				text: link.name,
				attr: {
					href: link.url,
					rel: 'noopener noreferrer',
					target: '_blank',
				},
			});
		}

		new Setting(containerEl)
			.setName(sectionLabels.search)
			.setHeading();

		new Setting(containerEl)
			.setName(settingDescriptors.enableDocxSearchIndex.name)
			.setDesc(settingDescriptors.enableDocxSearchIndex.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.enableDocxSearchIndex)
				.onChange(async (value) => {
					this.plugin.pluginSettings.enableDocxSearchIndex = value;
					await this.plugin.saveSettings();
					if (value) {
						await this.plugin.rebuildDocxSearchIndex(false);
					}
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.autoIndexDocxSearch.name)
			.setDesc(settingDescriptors.autoIndexDocxSearch.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.autoIndexDocxSearch)
				.onChange(async (value) => {
					this.plugin.pluginSettings.autoIndexDocxSearch = value;
					await this.plugin.saveSettings();
					if (value && this.plugin.pluginSettings.enableDocxSearchIndex) {
						await this.plugin.rebuildDocxSearchIndex(false);
					}
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.rebuildDocxSearchIndex.name)
			.setDesc(settingDescriptors.rebuildDocxSearchIndex.description)
			.addButton(button => button
				.setButtonText(settingDescriptors.rebuildDocxSearchIndex.actionLabel ?? i18n.t('common:actions.rebuild'))
				.onClick(async () => {
					await this.plugin.rebuildDocxSearchIndex(true);
				}));

		new Setting(containerEl)
			.setName(sectionLabels.ai)
			.setHeading();

		new Setting(containerEl)
			.setName(settingDescriptors.enableAiInterfacing.name)
			.setDesc(settingDescriptors.enableAiInterfacing.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.enableAiInterfacing)
				.onChange(async (value) => {
					this.plugin.pluginSettings.enableAiInterfacing = value;
					await this.plugin.saveSettings();
					await this.plugin.syncAiInterfacing();
				}));

		new Setting(containerEl)
			.setName(sectionLabels.diagnostics)
			.setHeading();

		new Setting(containerEl)
			.setName(settingDescriptors.debugLogging.name)
			.setDesc(settingDescriptors.debugLogging.description)
			.addToggle(toggle => toggle
				.setValue(this.plugin.pluginSettings.debugLogging)
				.onChange(async (value) => {
					this.plugin.pluginSettings.debugLogging = value;
					configureNativePowerPointDocEditorLogger(value);
					infoLog('settings', `Debug logging ${value ? 'enabled' : 'disabled'}`);
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.copyDocxLog.name)
			.setDesc(settingDescriptors.copyDocxLog.description)
			.addButton(button => button
				.setButtonText(settingDescriptors.copyDocxLog.actionLabel ?? i18n.t('common:actions.copy'))
				.onClick(async () => {
					await this.plugin.copyDebugLog('docx');
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.copyPptxLog.name)
			.setDesc(settingDescriptors.copyPptxLog.description)
			.addButton(button => button
				.setButtonText(settingDescriptors.copyPptxLog.actionLabel ?? i18n.t('common:actions.copy'))
				.onClick(async () => {
					await this.plugin.copyDebugLog('pptx');
				}));

		new Setting(containerEl)
			.setName(settingDescriptors.copyFullLog.name)
			.setDesc(settingDescriptors.copyFullLog.description)
			.addButton(button => button
				.setButtonText(settingDescriptors.copyFullLog.actionLabel ?? i18n.t('common:actions.copy'))
				.onClick(async () => {
					await this.plugin.copyDebugLog('all');
				}));

		const reportBugBox = containerEl.createDiv({ cls: 'native-powerpoint-report-bug' });
		reportBugBox.createEl('a', {
			cls: 'native-powerpoint-report-bug-link',
			text: 'Report bug',
			attr: {
				href: 'https://github.com/MarsLuay/NativePowerPointDocEditor/issues',
				rel: 'noopener noreferrer',
				target: '_blank',
			},
		});
		reportBugBox.createEl('a', {
			cls: 'native-powerpoint-report-bug-link',
			text: 'Buy me a coffee',
			attr: {
				href: 'https://buymeacoffee.com/marwanluaye',
				rel: 'noopener noreferrer',
				target: '_blank',
			},
		});
	}
}
