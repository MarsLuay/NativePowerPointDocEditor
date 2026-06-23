import type { TAbstractFile } from 'obsidian';
import { Notice, TFile } from 'obsidian';
import { processDocxEmbeds, registerDocxFileEmbed } from './DocxEmbedLoader';
import { DocxSearchModal } from './DocxSearchModal';
import { DocxView } from './DocxView';
import { VIEW_TYPE_DOCX } from './docxViewConstants';
import { configureDocxEditorChunkPaths } from './docxEditorLoader';
import { DocxSearchIndex } from './docxSearchIndex';
import { scheduleIdleWork } from './idleSchedule';
import { errorLog, infoLog } from './logger';
import { getDocxEditorLocale, loadDocxEditorLocale, preloadDocxEditorLocale } from './locales';
import type DocxidianPlugin from './main';
import type { DocxEditorSettingsController } from './DocxView';

export { createDocxReactMount, DocxFileEmbed, renderDocxEmbeds, hasReviewMarkup } from './docxEditorChunk';
export { DocxView, VIEW_TYPE_DOCX };

const DOCX_EXTENSIONS = ['docx'];

export async function registerDocxSupport(
	plugin: DocxidianPlugin,
	createDocxSettingsController: () => DocxEditorSettingsController,
): Promise<DocxSearchIndex> {
	configureDocxEditorChunkPaths([]);
	const docxSearchIndex = new DocxSearchIndex(plugin.app, plugin.manifest.dir);
	plugin.setDocxSearchIndex(docxSearchIndex);

	preloadDocxEditorLocale(plugin.settings.editorLanguage);
	void loadDocxEditorLocale(plugin.settings.editorLanguage);

	plugin.registerView(
		VIEW_TYPE_DOCX,
		(leaf) => new DocxView(
			leaf,
			() => plugin.settings.authorName,
			() => getDocxEditorLocale(plugin.settings.editorLanguage),
			() => plugin.settings.showRuler,
			() => plugin.settings.autosave,
			() => plugin.settings.createBackupsBeforeSave,
			() => plugin.settings.defaultZoom,
			createDocxSettingsController(),
		),
	);
	plugin.registerExtensions(DOCX_EXTENSIONS, VIEW_TYPE_DOCX);

	registerDocxFileEmbed(plugin, () => getDocxEditorLocale(plugin.settings.editorLanguage));
	registerDeferredDocxEmbedProcessor(plugin);

	registerDocxCommands(plugin, docxSearchIndex);
	registerDocxSearchEvents(plugin, docxSearchIndex);
	queueInitialDocxSearchIndex(plugin, docxSearchIndex);

	infoLog('plugin', 'DOCX support registered');
	return docxSearchIndex;
}

function registerDeferredDocxEmbedProcessor(plugin: DocxidianPlugin) {
	let registered = false;

	const register = () => {
		if (registered) {
			return;
		}

		registered = true;
		plugin.registerMarkdownPostProcessor((el, ctx) => {
			processDocxEmbeds(plugin.app, el, ctx, () => getDocxEditorLocale(plugin.settings.editorLanguage));
		}, 1000);
		infoLog('plugin', 'Registered deferred DOCX embed markdown processor');
	};

	if (typeof plugin.app.workspace?.on === 'function') {
		plugin.registerEvent(plugin.app.workspace.on('active-leaf-change', register));
	}

	scheduleIdleWork(register, { timeout: 5000 });
}

function registerDocxCommands(plugin: DocxidianPlugin, docxSearchIndex: DocxSearchIndex) {
	plugin.addCommand({
		id: 'save-current-docx',
		name: 'Save current docx',
		callback: async () => {
			const docxView = plugin.app.workspace.getActiveViewOfType(DocxView);
			if (!docxView) {
				new Notice('Open a docx file to save it.');
				return;
			}

			await docxView.saveCurrentDocument();
		},
	});
	plugin.addCommand({
		id: 'save-current-docx-as',
		name: 'Save current DOCX as...',
		callback: async () => {
			const docxView = plugin.app.workspace.getActiveViewOfType(DocxView);
			if (!docxView) {
				new Notice('Open a docx file to save a copy.');
				return;
			}

			await docxView.saveCurrentDocumentAs();
		},
	});
	plugin.addCommand({
		id: 'duplicate-current-docx',
		name: 'Duplicate current DOCX',
		callback: async () => {
			const docxView = plugin.app.workspace.getActiveViewOfType(DocxView);
			if (!docxView) {
				new Notice('Open a docx file to duplicate it.');
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
				new Notice('Open a docx file to search it.');
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
				new Notice('Open a docx file to search it.');
				return;
			}

			docxView.openFindReplaceDialog();
		},
	});
	plugin.addCommand({
		id: 'search-docx-files',
		name: 'Search DOCX files in vault',
		callback: async () => {
			if (!plugin.settings.enableDocxSearchIndex) {
				new Notice('Turn on the DOCX search index in Native PowerPoint Doc Editor settings first.');
				return;
			}

			await docxSearchIndex.load();
			new DocxSearchModal(plugin.app, docxSearchIndex).open();
		},
	});
	plugin.addCommand({
		id: 'rebuild-docx-search-index',
		name: 'Rebuild DOCX search index',
		callback: async () => {
			await rebuildDocxSearchIndex(plugin, docxSearchIndex, true, true);
		},
	});
}

function registerDocxSearchEvents(plugin: DocxidianPlugin, docxSearchIndex: DocxSearchIndex) {
	plugin.registerEvent(plugin.app.vault.on('create', file => handleDocxSearchFileChanged(plugin, docxSearchIndex, file)));
	plugin.registerEvent(plugin.app.vault.on('modify', file => handleDocxSearchFileChanged(plugin, docxSearchIndex, file)));
	plugin.registerEvent(plugin.app.vault.on('delete', file => handleDocxSearchFileDeleted(plugin, docxSearchIndex, file)));
	plugin.registerEvent(plugin.app.vault.on('rename', (file, oldPath) => {
		handleDocxSearchFileDeleted(plugin, docxSearchIndex, oldPath);
		void handleDocxSearchFileChanged(plugin, docxSearchIndex, file);
	}));
}

function queueInitialDocxSearchIndex(plugin: DocxidianPlugin, docxSearchIndex: DocxSearchIndex) {
	if (!plugin.settings.enableDocxSearchIndex || !plugin.settings.autoIndexDocxSearch) {
		return;
	}

	const cancelIdle = scheduleIdleWork(() => {
		void rebuildDocxSearchIndex(plugin, docxSearchIndex, false, false);
	}, { timeout: 10000 });

	plugin.register(() => cancelIdle());
}

async function handleDocxSearchFileChanged(
	plugin: DocxidianPlugin,
	docxSearchIndex: DocxSearchIndex,
	file: TAbstractFile,
) {
	if (!plugin.settings.enableDocxSearchIndex || !plugin.settings.autoIndexDocxSearch) {
		return;
	}

	if (!(file instanceof TFile) || !docxSearchIndex.isDocxFile(file)) {
		return;
	}

	await docxSearchIndex.load();
	void docxSearchIndex.indexFile(file);
}

function handleDocxSearchFileDeleted(
	plugin: DocxidianPlugin,
	docxSearchIndex: DocxSearchIndex,
	fileOrPath: TAbstractFile | string,
) {
	if (!plugin.settings.enableDocxSearchIndex) {
		return;
	}

	const path = typeof fileOrPath === 'string' ? fileOrPath : fileOrPath.path;
	if (!path.toLowerCase().endsWith('.docx')) {
		return;
	}

	void docxSearchIndex.removePath(path);
}

export async function rebuildDocxSearchIndex(
	plugin: DocxidianPlugin,
	docxSearchIndex: DocxSearchIndex,
	force = false,
	showNotice = true,
) {
	if (plugin.settings.disableDocxFiles) {
		if (showNotice) {
			new Notice('DOCX support is turned off for this plugin. Reload after turning it back on.');
		}
		return;
	}

	if (!plugin.settings.enableDocxSearchIndex) {
		return;
	}

	try {
		await docxSearchIndex.load();
		const stats = force
			? await docxSearchIndex.rebuildSync({ force: true })
			: await docxSearchIndex.rebuildIncremental({ force: false });
		if (showNotice) {
			new Notice(`DOCX search index ready: ${stats.total} files, ${stats.errors} errors.`);
		}
	} catch (error) {
		errorLog('search', 'Could not rebuild DOCX search index', error);
		if (showNotice) {
			new Notice('Could not rebuild DOCX search index. Check the Native PowerPoint Doc Editor debug log.');
		}
	}
}

export function refreshDocxViews(plugin: DocxidianPlugin) {
	for (const leaf of plugin.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)) {
		const view = leaf.view;
		if (view instanceof DocxView) {
			view.refreshSettings();
			void loadDocxEditorLocale(plugin.settings.editorLanguage).then(() => {
				view.refreshSettings();
			});
		}
	}
}
