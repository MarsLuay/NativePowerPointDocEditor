import {
	NativePowerPointView,
	NATIVE_POWERPOINT_VIEW_TYPE,
	POWERPOINT_EXTENSIONS,
	isPowerPointExtension,
} from './NativePowerPointView';
import { infoLog } from './logger';
import type NativePowerPointDocEditorPlugin from './main';
import { pptNotice, pptT } from './i18n/powerpointNotify';
import type { NativePowerPointSettings } from './settings';

export { NativePowerPointView, NATIVE_POWERPOINT_VIEW_TYPE, isPowerPointExtension };

export function registerPowerPointSupport(
	plugin: NativePowerPointDocEditorPlugin,
	getPowerPointSettings: () => NativePowerPointSettings,
) {
	plugin.registerView(
		NATIVE_POWERPOINT_VIEW_TYPE,
		(leaf) => new NativePowerPointView(leaf, getPowerPointSettings),
	);
	plugin.registerExtensions(POWERPOINT_EXTENSIONS, NATIVE_POWERPOINT_VIEW_TYPE);

	plugin.addCommand({
		id: 'open-powerpoint-file',
		name: pptT('powerpoint:commands.openPowerPointFile'),
		callback: () => {
			const file = plugin.app.workspace.getActiveFile();
			if (!file || !isPowerPointExtension(file.extension)) {
				pptNotice('powerpoint:notice.selectPowerPointToOpen');
				return;
			}

			const leaf = plugin.app.workspace.getLeaf('tab');
			void leaf.openFile(file, { active: true });
		},
	});
	plugin.addCommand({
		id: 'save-current-powerpoint-file',
		name: pptT('powerpoint:commands.saveCurrentPowerPointFile'),
		callback: async () => {
			const view = plugin.app.workspace.getActiveViewOfType(NativePowerPointView);
			if (!view) {
				pptNotice('powerpoint:notice.openPowerPointToSave');
				return;
			}

			await view.saveCurrentPresentation();
		},
	});

	infoLog('plugin', 'PowerPoint support registered');
}

export function refreshPowerPointViews(plugin: NativePowerPointDocEditorPlugin) {
	for (const leaf of plugin.app.workspace.getLeavesOfType(NATIVE_POWERPOINT_VIEW_TYPE)) {
		const view = leaf.view;
		if (view instanceof NativePowerPointView) {
			view.refreshSettings();
		}
	}
}
