export { loadDocxEditorMessages, type LoadedLocale } from './docxEditorLocaleAdapter';
export { getPluginI18n, initPluginI18n, resolvePluginLocale } from './pluginI18n';
export {
	createPluginI18nService,
	PluginI18nService,
	type I18nService,
	type MessageKey,
} from './I18nService';
export {
	LOCALE_NAMESPACES,
	listInstalledLocales,
	loadBundledPluginMessages,
	loadLocale,
	loadPluginMessagesFromAdapter,
	mergeNamespaceMessages,
	type LocaleFileAdapter,
	type LocaleNamespace,
	type PluginMessages,
} from './localeLoader';
export { getObsidianLocale } from './obsidianLocale';
export {
	getLocaleDirection,
	localeCandidates,
	resolveAutomaticLocale,
} from './localeResolver';
export { formatMessage } from './messageFormat';
export {
	getDocxEditorSettingDescriptors,
	getDocxEditorSettingSectionLabels,
	getDocxEditorSettingsMenuSections,
	type DocxEditorSettingDescriptor,
} from './docxSettingsCatalog';
export {
	getEditorThemeSettingOptions,
	getNativePowerPointDocEditorSettingDescriptors,
	getNativePowerPointDocEditorSettingSectionLabels,
	getNativePowerPointDocEditorSettingsTabSections,
	type NativePowerPointDocEditorSettingDescriptor,
	type SettingsOption,
} from './settingsCatalog';
export { showI18nNotice } from './notify';
export { createPowerPointTranslate, pptNotice, pptT } from './powerpointNotify';
export type { TranslateFn, TranslateNoticeFn, TranslateValues } from './translate';
export { createTranslateNotice } from './translate';
