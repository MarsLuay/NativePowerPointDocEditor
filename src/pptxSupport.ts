import { Notice } from 'obsidian';
import {
	NativePowerPointView,
	NATIVE_POWERPOINT_VIEW_TYPE,
	POWERPOINT_EXTENSIONS,
	isPowerPointExtension,
} from './NativePowerPointView';
import { infoLog } from './logger';
import type DocxidianPlugin from './main';
import type { NativePowerPointSettings } from './settings';

export { NativePowerPointView, NATIVE_POWERPOINT_VIEW_TYPE, isPowerPointExtension };

export function registerPowerPointSupport(
	plugin: DocxidianPlugin,
	getPowerPointSettings: () => NativePowerPointSettings,
) {
	plugin.registerView(
		NATIVE_POWERPOINT_VIEW_TYPE,
		(leaf) => new NativePowerPointView(leaf, getPowerPointSettings),
	);
	plugin.registerExtensions(POWERPOINT_EXTENSIONS, NATIVE_POWERPOINT_VIEW_TYPE);

	plugin.addCommand({
		id: 'open-powerpoint-file',
		name: 'Open PowerPoint file',
		callback: () => {
			const file = plugin.app.workspace.getActiveFile();
			if (!file || !isPowerPointExtension(file.extension)) {
				new Notice('Select a PowerPoint file to open it.');
				return;
			}

			const leaf = plugin.app.workspace.getLeaf('tab');
			void leaf.openFile(file, { active: true });
		},
	});
	plugin.addCommand({
		id: 'save-current-powerpoint-file',
		name: 'Save current PowerPoint file',
		callback: async () => {
			const view = plugin.app.workspace.getActiveViewOfType(NativePowerPointView);
			if (!view) {
				new Notice('Open a PowerPoint file to save it.');
				return;
			}

			await view.saveCurrentPresentation();
		},
	});

	infoLog('plugin', 'PowerPoint support registered');
}

export function refreshPowerPointViews(plugin: DocxidianPlugin) {
	for (const leaf of plugin.app.workspace.getLeavesOfType(NATIVE_POWERPOINT_VIEW_TYPE)) {
		const view = leaf.view;
		if (view instanceof NativePowerPointView) {
			view.refreshSettings();
		}
	}
}
