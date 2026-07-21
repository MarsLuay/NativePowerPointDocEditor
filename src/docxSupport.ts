import type { App, Menu, TAbstractFile } from 'obsidian';
import { TFile, TFolder, normalizePath } from 'obsidian';
import type { DocxViewAgentBridge } from './ai/aiRuntime';
import { showI18nNotice } from './i18n/notify';
import { processDocxEmbeds, registerDocxFileEmbed } from './DocxEmbedLoader';
import { DocxSearchModal } from './DocxSearchModal';
import { DocxView } from './DocxView';
import { VIEW_TYPE_DOCX } from './docxViewConstants';
import { configureDocxEditorChunkPaths } from './docxEditorLoader';
import { DocxSearchIndex } from './docxSearchIndex';
import { scheduleIdleWork } from './idleSchedule';
import { errorLog, infoLog } from './logger';
import { getDocxEditorLocale, loadDocxEditorLocale, preloadDocxEditorLocale } from './locales';
import type NativePowerPointDocEditorPlugin from './main';
import { createAndOpenNewOfficeFile } from './vault/createNewOfficeFile';
import { convertMarkdownFileToDocx } from './vault/markdownToDocx';

export { createDocxReactMount, DocxFileEmbed, renderDocxEmbeds, hasReviewMarkup } from './docxEditorChunk';
export { DocxView, VIEW_TYPE_DOCX };

const DOCX_EXTENSIONS = ['docx'];

export async function registerDocxSupport(
	plugin: NativePowerPointDocEditorPlugin,
): Promise<DocxSearchIndex> {
	configureDocxEditorChunkPaths([]);
	const docxSearchIndex = new DocxSearchIndex(plugin.app, plugin.manifest.dir);
	plugin.setDocxSearchIndex(docxSearchIndex);

	preloadDocxEditorLocale(plugin.getResolvedDocxEditorLanguage());
	void loadDocxEditorLocale(plugin.getResolvedDocxEditorLanguage());

	plugin.registerView(
		VIEW_TYPE_DOCX,
		(leaf) => new DocxView(
			leaf,
			() => plugin.pluginSettings.authorName,
			() => plugin.pluginSettings.editorTheme,
			() => plugin.getResolvedEditorTheme(),
			() => getDocxEditorLocale(plugin.getResolvedDocxEditorLanguage()),
			() => plugin.getI18n(),
			() => plugin.pluginSettings.showRuler,
			() => plugin.pluginSettings.autosave,
			() => plugin.pluginSettings.createBackupsBeforeSave,
			() => plugin.pluginSettings.defaultZoom,
			(wordCount) => plugin.updateDocumentWordCount(leaf, wordCount),
			() => plugin.clearDocumentWordCount(leaf),
		),
	);
	plugin.registerExtensions(DOCX_EXTENSIONS, VIEW_TYPE_DOCX);

	registerDocxFileEmbed(plugin, () => getDocxEditorLocale(plugin.getResolvedDocxEditorLanguage()));
	registerDeferredDocxEmbedProcessor(plugin);

	registerDocxCommands(plugin, docxSearchIndex);
	registerDocxFileMenu(plugin);
	registerDocxSearchEvents(plugin, docxSearchIndex);
	queueInitialDocxSearchIndex(plugin, docxSearchIndex);

	infoLog('plugin', 'DOCX support registered');
	return docxSearchIndex;
}

function registerDocxFileMenu(plugin: NativePowerPointDocEditorPlugin) {
	plugin.registerEvent(plugin.app.workspace.on('file-menu', (menu, file) => {
		if (file instanceof TFolder) {
			addCreateDocxMenuItem(plugin, menu, file);
			return;
		}

		if (file instanceof TFile && file.extension.toLowerCase() === 'md') {
			addConvertMarkdownToDocxMenuItem(plugin, menu, file);
		}
	}));
}

function addCreateDocxMenuItem(plugin: NativePowerPointDocEditorPlugin, menu: Menu, folder: TFolder) {
	const i18n = plugin.getI18n();
	menu.addItem((item) => {
		item
			.setTitle(i18n?.t('docx:menu.newDocx') ?? 'New DOCX')
			.setIcon('file-text')
			.onClick(async () => {
				try {
					await createAndOpenNewOfficeFile(plugin.app, folder, 'docx');
				} catch (error) {
					const message = error instanceof Error ? error.message : String(error);
					showI18nNotice(i18n, 'docx:notice.createFailed', { message });
					errorLog('file', 'Failed to create blank DOCX', { folder: folder.path, message });
				}
			});
	});
}

function addConvertMarkdownToDocxMenuItem(
	plugin: NativePowerPointDocEditorPlugin,
	menu: Menu,
	file: TFile,
) {
	const i18n = plugin.getI18n();
	menu.addItem((item) => {
		item
			.setTitle(i18n?.t('docx:menu.convertMarkdownToDocx') ?? 'Convert to DOCX')
			.setIcon('file-output')
			.onClick(async () => {
				try {
					const outputFile = await convertMarkdownFileToDocx(plugin.app, file);
					showI18nNotice(i18n, 'docx:notice.convertedMarkdownToDocx', { path: outputFile.path });
				} catch (error) {
					const message = error instanceof Error ? error.message : String(error);
					showI18nNotice(i18n, 'docx:notice.convertMarkdownFailed', { message });
					errorLog('file', 'Failed to convert Markdown to DOCX', { sourcePath: file.path, message });
				}
			});
	});
}

function registerDeferredDocxEmbedProcessor(plugin: NativePowerPointDocEditorPlugin) {
	let registered = false;

	const register = () => {
		if (registered) {
			return;
		}

		registered = true;
		plugin.registerMarkdownPostProcessor((el, ctx) => {
			processDocxEmbeds(plugin.app, el, ctx, () => getDocxEditorLocale(plugin.getResolvedDocxEditorLanguage()));
		}, 1000);
		infoLog('plugin', 'Registered deferred DOCX embed markdown processor');
	};

	if (typeof plugin.app.workspace?.on === 'function') {
		plugin.registerEvent(plugin.app.workspace.on('active-leaf-change', register));
	}

	scheduleIdleWork(register, { timeout: 5000 });
}

function registerDocxCommands(plugin: NativePowerPointDocEditorPlugin, docxSearchIndex: DocxSearchIndex) {
	plugin.addCommand({
		id: 'save-current-docx',
		name: 'Save current docx',
		callback: async () => {
			const docxView = plugin.app.workspace.getActiveViewOfType(DocxView);
			if (!docxView) {
				showI18nNotice(plugin.getI18n(), 'docx:notice.openToSave');
				return;
			}

			await docxView.saveCurrentDocument();
		},
	});
	plugin.addCommand({
		id: 'save-current-docx-as',
		name: 'Save current docx as...',
		callback: async () => {
			const docxView = plugin.app.workspace.getActiveViewOfType(DocxView);
			if (!docxView) {
				showI18nNotice(plugin.getI18n(), 'docx:notice.openToSaveCopy');
				return;
			}

			await docxView.saveCurrentDocumentAs();
		},
	});
	plugin.addCommand({
		id: 'duplicate-current-docx',
		name: 'Duplicate current docx',
		callback: async () => {
			const docxView = plugin.app.workspace.getActiveViewOfType(DocxView);
			if (!docxView) {
				showI18nNotice(plugin.getI18n(), 'docx:notice.openToDuplicate');
				return;
			}

			await docxView.duplicateCurrentDocument();
		},
	});
	plugin.addCommand({
		id: 'find-in-current-docx',
		name: 'Find in current docx',
		callback: () => {
			const docxView = plugin.app.workspace.getActiveViewOfType(DocxView);
			if (!docxView) {
				showI18nNotice(plugin.getI18n(), 'docx:notice.openToSearch');
				return;
			}

			docxView.openFindDialog();
		},
	});
	plugin.addCommand({
		id: 'find-replace-in-current-docx',
		name: 'Find and replace in current docx',
		callback: () => {
			const docxView = plugin.app.workspace.getActiveViewOfType(DocxView);
			if (!docxView) {
				showI18nNotice(plugin.getI18n(), 'docx:notice.openToSearch');
				return;
			}

			docxView.openFindReplaceDialog();
		},
	});
	plugin.addCommand({
		id: 'search-docx-files',
		name: 'Search docx files in vault',
		callback: async () => {
			if (!plugin.pluginSettings.enableDocxSearchIndex) {
				showI18nNotice(plugin.getI18n(), 'docx:notice.enableSearchIndexFirst');
				return;
			}

			await docxSearchIndex.load();
			new DocxSearchModal(plugin.app, docxSearchIndex).open();
		},
	});
	plugin.addCommand({
		id: 'rebuild-docx-search-index',
		name: 'Rebuild docx search index',
		callback: async () => {
			await rebuildDocxSearchIndex(plugin, docxSearchIndex, true, true);
		},
	});
}

function registerDocxSearchEvents(plugin: NativePowerPointDocEditorPlugin, docxSearchIndex: DocxSearchIndex) {
	plugin.registerEvent(plugin.app.vault.on('create', file => handleDocxSearchFileChanged(plugin, docxSearchIndex, file)));
	plugin.registerEvent(plugin.app.vault.on('modify', file => handleDocxSearchFileChanged(plugin, docxSearchIndex, file)));
	plugin.registerEvent(plugin.app.vault.on('delete', file => handleDocxSearchFileDeleted(plugin, docxSearchIndex, file)));
	plugin.registerEvent(plugin.app.vault.on('rename', (file, oldPath) => {
		handleDocxSearchFileDeleted(plugin, docxSearchIndex, oldPath);
		void handleDocxSearchFileChanged(plugin, docxSearchIndex, file);
	}));
}

function queueInitialDocxSearchIndex(plugin: NativePowerPointDocEditorPlugin, docxSearchIndex: DocxSearchIndex) {
	if (!plugin.pluginSettings.enableDocxSearchIndex || !plugin.pluginSettings.autoIndexDocxSearch) {
		return;
	}

	const cancelIdle = scheduleIdleWork(() => {
		void rebuildDocxSearchIndex(plugin, docxSearchIndex, false, false);
	}, { timeout: 10000 });

	plugin.register(() => cancelIdle());
}

async function handleDocxSearchFileChanged(
	plugin: NativePowerPointDocEditorPlugin,
	docxSearchIndex: DocxSearchIndex,
	file: TAbstractFile,
) {
	if (!plugin.pluginSettings.enableDocxSearchIndex || !plugin.pluginSettings.autoIndexDocxSearch) {
		return;
	}

	if (!(file instanceof TFile) || !docxSearchIndex.isDocxFile(file)) {
		return;
	}

	await docxSearchIndex.load();
	void docxSearchIndex.indexFile(file);
}

function handleDocxSearchFileDeleted(
	plugin: NativePowerPointDocEditorPlugin,
	docxSearchIndex: DocxSearchIndex,
	fileOrPath: TAbstractFile | string,
) {
	if (!plugin.pluginSettings.enableDocxSearchIndex) {
		return;
	}

	const path = typeof fileOrPath === 'string' ? fileOrPath : fileOrPath.path;
	if (!path.toLowerCase().endsWith('.docx')) {
		return;
	}

	void docxSearchIndex.removePath(path);
}

export async function rebuildDocxSearchIndex(
	plugin: NativePowerPointDocEditorPlugin,
	docxSearchIndex: DocxSearchIndex,
	force = false,
	showNotice = true,
) {
	if (plugin.pluginSettings.disableDocxFiles) {
		if (showNotice) {
			showI18nNotice(plugin.getI18n(), 'docx:notice.supportDisabled');
		}
		return;
	}

	if (!plugin.pluginSettings.enableDocxSearchIndex) {
		return;
	}

	try {
		await docxSearchIndex.load();
		const stats = force
			? await docxSearchIndex.rebuildSync({ force: true })
			: await docxSearchIndex.rebuildIncremental({ force: false });
		if (showNotice) {
			showI18nNotice(plugin.getI18n(), 'docx:notice.searchIndexReady', {
				total: stats.total,
				errors: stats.errors,
			});
		}
	} catch (error) {
		errorLog('search', 'Could not rebuild DOCX search index', error);
		if (showNotice) {
			showI18nNotice(plugin.getI18n(), 'docx:notice.searchIndexRebuildFailed');
		}
	}
}

export function findDocxViewForPath(app: App, path: string): DocxViewAgentBridge | null {
	const normalized = normalizePath(path);
	for (const leaf of app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)) {
		const view = leaf.view;
		if (view instanceof DocxView && view.getLoadedDocumentPath() === normalized) {
			return view;
		}
	}
	return null;
}

/**
 * Persist every dirty DOCX view before the development hot-reloader disables
 * the plugin. Returns false when any source file could not be updated so the
 * caller can keep the current plugin instance and its in-memory edits alive.
 */
export async function saveDocxViewsBeforePluginReload(
	plugin: NativePowerPointDocEditorPlugin,
): Promise<boolean> {
	const views = plugin.app.workspace
		.getLeavesOfType(VIEW_TYPE_DOCX)
		.map((leaf) => leaf.view)
		.filter((view): view is DocxView => view instanceof DocxView);

	const results = await Promise.all(views.map((view) => view.saveBeforePluginReload()));
	return results.every(Boolean);
}

export function refreshDocxViews(plugin: NativePowerPointDocEditorPlugin) {
	for (const leaf of plugin.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)) {
		const view = leaf.view;
		if (view instanceof DocxView) {
			view.refreshSettings();
			void loadDocxEditorLocale(plugin.getResolvedDocxEditorLanguage())
				.then(() => {
					view.refreshSettings();
				})
				.catch((error) => {
					errorLog('i18n', 'Could not refresh DOCX editor locale', error);
				});
		}
	}
}
