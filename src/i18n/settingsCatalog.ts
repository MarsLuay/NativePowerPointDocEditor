import type { I18nService } from './I18nService';
import {
	DEFAULT_SETTINGS,
	type DocxEditorSettingId,
	type DocxEditorSettingSectionId,
	type EditorThemePreference,
	type NativePowerPointDocEditorSettingId,
	type NativePowerPointDocEditorSettingSectionId,
} from '../settings';

export interface SettingsOption<TValue extends string = string> {
	value: TValue;
	label: string;
}

export interface NativePowerPointDocEditorSettingDescriptor {
	sectionId: NativePowerPointDocEditorSettingSectionId;
	name: string;
	description: string;
	defaultValue?: string | number | boolean;
	actionLabel?: string;
	resetLabel?: string;
	placeholder?: string;
}

export type DocxEditorSettingDescriptor = NativePowerPointDocEditorSettingDescriptor & {
	sectionId: DocxEditorSettingSectionId;
};

export function getNativePowerPointDocEditorSettingSectionLabels(
	i18n: I18nService,
): Record<NativePowerPointDocEditorSettingSectionId, string> {
	return {
		identity: i18n.t('settings:section.identity'),
		fileHandoff: i18n.t('settings:fileHandoff.section'),
		editorDefaults: i18n.t('settings:section.editorDefaults'),
		saving: i18n.t('settings:section.saving'),
		powerpoint: i18n.t('settings:section.powerpoint'),
		search: i18n.t('settings:section.search'),
		ai: i18n.t('settings:section.ai'),
		diagnostics: i18n.t('settings:section.diagnostics'),
	};
}

export function getNativePowerPointDocEditorSettingDescriptors(
	i18n: I18nService,
): Record<NativePowerPointDocEditorSettingId, NativePowerPointDocEditorSettingDescriptor> {
	return {
		authorName: {
			sectionId: 'identity',
			name: i18n.t('settings:docx.authorName.name'),
			description: i18n.t('settings:docx.authorName.description'),
			defaultValue: DEFAULT_SETTINGS.authorName,
			placeholder: DEFAULT_SETTINGS.authorName,
			resetLabel: i18n.t('common:actions.reset'),
		},
		disableDocxFiles: {
			sectionId: 'fileHandoff',
			name: i18n.t('settings:docx.disableDocxFiles.name'),
			description: i18n.t('settings:docx.disableDocxFiles.description'),
			defaultValue: DEFAULT_SETTINGS.disableDocxFiles,
		},
		disablePowerPointFiles: {
			sectionId: 'fileHandoff',
			name: i18n.t('settings:fileHandoff.disablePptx.name'),
			description: i18n.t('settings:fileHandoff.disablePptx.description'),
			defaultValue: DEFAULT_SETTINGS.disablePowerPointFiles,
		},
		editorTheme: {
			sectionId: 'editorDefaults',
			name: i18n.t('settings:docx.editorTheme.name'),
			description: i18n.t('settings:docx.editorTheme.description'),
			defaultValue: DEFAULT_SETTINGS.editorTheme,
		},
		showRuler: {
			sectionId: 'editorDefaults',
			name: i18n.t('settings:docx.showRuler.name'),
			description: i18n.t('settings:docx.showRuler.description'),
			defaultValue: DEFAULT_SETTINGS.showRuler,
		},
		defaultZoom: {
			sectionId: 'editorDefaults',
			name: i18n.t('settings:docx.defaultZoom.name'),
			description: i18n.t('settings:docx.defaultZoom.description'),
			defaultValue: DEFAULT_SETTINGS.defaultZoom,
			resetLabel: i18n.t('common:actions.reset'),
		},
		autosave: {
			sectionId: 'saving',
			name: i18n.t('settings:docx.autosave.name'),
			description: i18n.t('settings:docx.autosave.description'),
			defaultValue: DEFAULT_SETTINGS.autosave,
		},
		createBackupsBeforeSave: {
			sectionId: 'saving',
			name: i18n.t('settings:docx.createBackupsBeforeSave.name'),
			description: i18n.t('settings:docx.createBackupsBeforeSave.description'),
			defaultValue: DEFAULT_SETTINGS.createBackupsBeforeSave,
		},
		powerPointAutosaveEnabled: {
			sectionId: 'saving',
			name: i18n.t('settings:powerpoint.autosave.name'),
			description: i18n.t('settings:powerpoint.autosave.description'),
			defaultValue: DEFAULT_SETTINGS.powerPointAutosaveEnabled,
		},
		powerPointShowInspector: {
			sectionId: 'powerpoint',
			name: i18n.t('settings:powerpoint.showInspector.name'),
			description: i18n.t('settings:powerpoint.showInspector.description'),
			defaultValue: DEFAULT_SETTINGS.powerPointShowInspector,
		},
		powerPointHideUnsupportedSvgContent: {
			sectionId: 'powerpoint',
			name: i18n.t('settings:powerpoint.hideUnsupportedSvg.name'),
			description: i18n.t('settings:powerpoint.hideUnsupportedSvg.description'),
			defaultValue: DEFAULT_SETTINGS.powerPointHideUnsupportedSvgContent,
		},
		powerPointOpenWithYoloMode: {
			sectionId: 'powerpoint',
			name: i18n.t('settings:powerpoint.yoloMode.name'),
			description: i18n.t('settings:powerpoint.yoloMode.description'),
			defaultValue: DEFAULT_SETTINGS.powerPointOpenWithYoloMode,
		},
		enableDocxSearchIndex: {
			sectionId: 'search',
			name: i18n.t('settings:docx.enableDocxSearchIndex.name'),
			description: i18n.t('settings:docx.enableDocxSearchIndex.description'),
			defaultValue: DEFAULT_SETTINGS.enableDocxSearchIndex,
		},
		autoIndexDocxSearch: {
			sectionId: 'search',
			name: i18n.t('settings:docx.autoIndexDocxSearch.name'),
			description: i18n.t('settings:docx.autoIndexDocxSearch.description'),
			defaultValue: DEFAULT_SETTINGS.autoIndexDocxSearch,
		},
		rebuildDocxSearchIndex: {
			sectionId: 'search',
			name: i18n.t('settings:docx.rebuildDocxSearchIndex.name'),
			description: i18n.t('settings:docx.rebuildDocxSearchIndex.description'),
			actionLabel: i18n.t('common:actions.rebuild'),
		},
		enableAiInterfacing: {
			sectionId: 'ai',
			name: i18n.t('settings:ai.enableInterfacing.name'),
			description: i18n.t('settings:ai.enableInterfacing.description'),
			defaultValue: DEFAULT_SETTINGS.enableAiInterfacing,
		},
		debugLogging: {
			sectionId: 'diagnostics',
			name: i18n.t('settings:docx.debugLogging.name'),
			description: i18n.t('settings:docx.debugLogging.description'),
			defaultValue: DEFAULT_SETTINGS.debugLogging,
		},
		copyDocxLog: {
			sectionId: 'diagnostics',
			name: i18n.t('settings:docx.copyDocxLog.name'),
			description: i18n.t('settings:docx.copyDocxLog.description'),
			actionLabel: i18n.t('common:actions.copy'),
		},
		copyPptxLog: {
			sectionId: 'diagnostics',
			name: i18n.t('settings:powerpoint.copyPptxLog.name'),
			description: i18n.t('settings:powerpoint.copyPptxLog.description'),
			actionLabel: i18n.t('common:actions.copy'),
		},
		copyFullLog: {
			sectionId: 'diagnostics',
			name: i18n.t('settings:debug.copyFullLog.name'),
			description: i18n.t('settings:debug.copyFullLog.description'),
			actionLabel: i18n.t('settings:debug.copyFullLog.actionLabel'),
		},
	};
}

export function getNativePowerPointDocEditorSettingsTabSections(i18n: I18nService): ReadonlyArray<{
	id: NativePowerPointDocEditorSettingSectionId;
	label: string;
	settings: readonly NativePowerPointDocEditorSettingId[];
}> {
	const sectionLabels = getNativePowerPointDocEditorSettingSectionLabels(i18n);
	return [
		{ id: 'identity', label: sectionLabels.identity, settings: ['authorName'] },
		{ id: 'fileHandoff', label: sectionLabels.fileHandoff, settings: ['disableDocxFiles', 'disablePowerPointFiles'] },
		{ id: 'editorDefaults', label: sectionLabels.editorDefaults, settings: ['editorTheme', 'showRuler', 'defaultZoom'] },
		{ id: 'saving', label: sectionLabels.saving, settings: ['autosave', 'createBackupsBeforeSave', 'powerPointAutosaveEnabled'] },
		{ id: 'powerpoint', label: sectionLabels.powerpoint, settings: ['powerPointShowInspector', 'powerPointHideUnsupportedSvgContent', 'powerPointOpenWithYoloMode'] },
		{ id: 'search', label: sectionLabels.search, settings: ['enableDocxSearchIndex', 'autoIndexDocxSearch', 'rebuildDocxSearchIndex'] },
		{ id: 'ai', label: sectionLabels.ai, settings: ['enableAiInterfacing'] },
		{ id: 'diagnostics', label: sectionLabels.diagnostics, settings: ['debugLogging', 'copyDocxLog', 'copyPptxLog', 'copyFullLog'] },
	];
}

export function getDocxEditorSettingSectionLabels(i18n: I18nService): Record<DocxEditorSettingSectionId, string> {
	const sectionLabels = getNativePowerPointDocEditorSettingSectionLabels(i18n);
	return {
		identity: sectionLabels.identity,
		editorDefaults: sectionLabels.editorDefaults,
		saving: sectionLabels.saving,
		search: sectionLabels.search,
		fileHandoff: i18n.t('settings:section.fileHandoff'),
		diagnostics: sectionLabels.diagnostics,
	};
}

export function getDocxEditorSettingDescriptors(i18n: I18nService): Record<DocxEditorSettingId, DocxEditorSettingDescriptor> {
	const descriptors = getNativePowerPointDocEditorSettingDescriptors(i18n);
	return {
		authorName: descriptors.authorName as DocxEditorSettingDescriptor,
		editorTheme: descriptors.editorTheme as DocxEditorSettingDescriptor,
		showRuler: descriptors.showRuler as DocxEditorSettingDescriptor,
		defaultZoom: descriptors.defaultZoom as DocxEditorSettingDescriptor,
		autosave: descriptors.autosave as DocxEditorSettingDescriptor,
		createBackupsBeforeSave: descriptors.createBackupsBeforeSave as DocxEditorSettingDescriptor,
		enableDocxSearchIndex: descriptors.enableDocxSearchIndex as DocxEditorSettingDescriptor,
		autoIndexDocxSearch: descriptors.autoIndexDocxSearch as DocxEditorSettingDescriptor,
		rebuildDocxSearchIndex: descriptors.rebuildDocxSearchIndex as DocxEditorSettingDescriptor,
		disableDocxFiles: descriptors.disableDocxFiles as DocxEditorSettingDescriptor,
		debugLogging: descriptors.debugLogging as DocxEditorSettingDescriptor,
		copyDocxLog: descriptors.copyDocxLog as DocxEditorSettingDescriptor,
	};
}

export function getDocxEditorSettingsMenuSections(i18n: I18nService): ReadonlyArray<{
	id: DocxEditorSettingSectionId;
	label: string;
	settings: readonly DocxEditorSettingId[];
}> {
	const sectionLabels = getDocxEditorSettingSectionLabels(i18n);
	return [
		{ id: 'identity', label: sectionLabels.identity, settings: ['authorName'] },
		{ id: 'editorDefaults', label: sectionLabels.editorDefaults, settings: ['editorTheme', 'showRuler', 'defaultZoom'] },
		{ id: 'saving', label: sectionLabels.saving, settings: ['autosave', 'createBackupsBeforeSave'] },
		{ id: 'search', label: sectionLabels.search, settings: ['enableDocxSearchIndex', 'autoIndexDocxSearch', 'rebuildDocxSearchIndex'] },
		{ id: 'fileHandoff', label: sectionLabels.fileHandoff, settings: ['disableDocxFiles'] },
		{ id: 'diagnostics', label: sectionLabels.diagnostics, settings: ['debugLogging', 'copyDocxLog'] },
	];
}

export function getEditorThemeSettingOptions(
	i18n: I18nService,
	defaultTheme: EditorThemePreference,
): SettingsOption<EditorThemePreference>[] {
	const options = [
		{ value: 'system' as const, label: i18n.t('settings:theme.system') },
		{ value: 'light' as const, label: i18n.t('settings:theme.light') },
		{ value: 'dark' as const, label: i18n.t('settings:theme.dark') },
	];
	const suffix = i18n.t('settings:theme.defaultSuffix');
	return options.map((option) => ({
		value: option.value,
		label: option.value === defaultTheme ? `${option.label} ${suffix}` : option.label,
	}));
}
