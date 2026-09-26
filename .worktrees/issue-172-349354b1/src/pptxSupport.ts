import {
	NativePowerPointView,
	NATIVE_POWERPOINT_VIEW_TYPE,
	POWERPOINT_EXTENSIONS,
	isPowerPointExtension,
} from './NativePowerPointView';
import type { App, Menu } from 'obsidian';
import { TFolder, normalizePath } from 'obsidian';
import type { PptxViewAgentBridge } from './ai/aiRuntime';
import { errorLog, infoLog } from './logger';
import type NativePowerPointDocEditorPlugin from './main';
import { pptNotice, pptT } from './i18n/powerpointNotify';
import type { NativePowerPointSettings } from './settings';
import { createAndOpenNewOfficeFile } from './vault/createNewOfficeFile';

export { NativePowerPointView, NATIVE_POWERPOINT_VIEW_TYPE, isPowerPointExtension };

export function registerPowerPointSupport(
	plugin: NativePowerPointDocEditorPlugin,
	getPowerPointSettings: () => NativePowerPointSettings,
) {
	plugin.registerView(
		NATIVE_POWERPOINT_VIEW_TYPE,
		(leaf) => new NativePowerPointView(
			leaf,
			getPowerPointSettings,
			(wordCount) => plugin.updateDocumentWordCount(leaf, wordCount),
			() => plugin.clearDocumentWordCount(leaf),
		),
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

	registerPowerPointFolderMenu(plugin);

	infoLog('plugin', 'PowerPoint support registered');
}

function registerPowerPointFolderMenu(plugin: NativePowerPointDocEditorPlugin) {
	plugin.registerEvent(plugin.app.workspace.on('file-menu', (menu, file) => {
		if (!(file instanceof TFolder)) {
			return;
		}

		addCreatePptxMenuItem(plugin, menu, file);
	}));
}

function addCreatePptxMenuItem(plugin: NativePowerPointDocEditorPlugin, menu: Menu, folder: TFolder) {
	menu.addItem((item) => {
		item
			.setTitle(pptT('powerpoint:menu.newPptx'))
			.setIcon('presentation')
			.onClick(async () => {
				try {
					await createAndOpenNewOfficeFile(plugin.app, folder, 'pptx');
				} catch (error) {
					const message = error instanceof Error ? error.message : String(error);
					pptNotice('powerpoint:notice.createFailed', { message });
					errorLog('file', 'Failed to create blank PPTX', { folder: folder.path, message });
				}
			});
	});
}

export function refreshPowerPointViews(plugin: NativePowerPointDocEditorPlugin) {
	for (const leaf of plugin.app.workspace.getLeavesOfType(NATIVE_POWERPOINT_VIEW_TYPE)) {
		const view = leaf.view;
		if (view instanceof NativePowerPointView) {
			view.refreshSettings();
		}
	}
}

/**
 * Persist every dirty PPTX view before the development hot-reloader disables
 * the plugin. Returns false when any source file could not be updated so the
 * caller can keep the current plugin instance and its in-memory edits alive.
 */
export async function savePowerPointViewsBeforePluginReload(
	plugin: NativePowerPointDocEditorPlugin,
): Promise<boolean> {
	const views = plugin.app.workspace
		.getLeavesOfType(NATIVE_POWERPOINT_VIEW_TYPE)
		.map((leaf) => leaf.view)
		.filter((view): view is NativePowerPointView => view instanceof NativePowerPointView);

	const results = await Promise.all(views.map((view) => view.saveBeforePluginReload()));
	return results.every(Boolean);
}

export function findPptxViewForPath(app: App, path: string): PptxViewAgentBridge | null {
	const normalized = normalizePath(path);
	for (const leaf of app.workspace.getLeavesOfType(NATIVE_POWERPOINT_VIEW_TYPE)) {
		const view = leaf.view;
		if (view instanceof NativePowerPointView && view.getLoadedPresentationPath() === normalized) {
			return view;
		}
	}
	return null;
}
