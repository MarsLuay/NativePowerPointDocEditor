import type { App, TFile, WorkspaceLeaf } from 'obsidian';

const MELD_ENCRYPTED_MARKDOWN_EXTENSIONS = new Set(['mdenc', 'encrypted']);
const MELD_ENCRYPTED_VIEW_TYPE = 'meld-encrypted-view';

interface MeldEncryptedMarkdownView {
	file?: { path: string } | null;
	getViewType?: () => string;
	getUnencryptedViewData?: () => string;
	isSavingEnabled?: boolean;
}

export function isMarkdownDocxSourceExtension(extension: string): boolean {
	const normalizedExtension = extension.toLowerCase();
	return normalizedExtension === 'md' || MELD_ENCRYPTED_MARKDOWN_EXTENSIONS.has(normalizedExtension);
}

/** Read Markdown through Meld's unlocked view without changing its encrypted source file. */
export async function readMarkdownSourceForDocx(app: App, sourceFile: TFile): Promise<string> {
	const extension = sourceFile.extension.toLowerCase();
	if (extension === 'md') {
		return app.vault.read(sourceFile);
	}
	if (!MELD_ENCRYPTED_MARKDOWN_EXTENSIONS.has(extension)) {
		throw new Error('Only Markdown and Meld Encrypt Markdown files can be converted to DOCX.');
	}

	const openLeaf = findOpenMeldEncryptedLeaf(app, sourceFile);
	if (openLeaf) {
		return readUnlockedMeldMarkdown(openLeaf, sourceFile);
	}

	// Opening through Obsidian lets Meld own its password prompt and decryption path.
	const temporaryLeaf = app.workspace.getLeaf('tab');
	try {
		await temporaryLeaf.openFile(sourceFile, { active: true });
		return readUnlockedMeldMarkdown(temporaryLeaf, sourceFile);
	} finally {
		detachTemporaryLeafIfOpen(app, temporaryLeaf);
	}
}

function findOpenMeldEncryptedLeaf(app: App, sourceFile: TFile): WorkspaceLeaf | null {
	let foundLeaf: WorkspaceLeaf | null = null;
	app.workspace.iterateAllLeaves((leaf) => {
		const view = leaf.view as unknown as MeldEncryptedMarkdownView | null;
		if (view?.getViewType?.() === MELD_ENCRYPTED_VIEW_TYPE && view.file?.path === sourceFile.path) {
			foundLeaf = leaf;
		}
	});
	return foundLeaf;
}

function detachTemporaryLeafIfOpen(app: App, temporaryLeaf: WorkspaceLeaf): void {
	let isOpen = false;
	app.workspace.iterateAllLeaves((leaf) => {
		if (leaf === temporaryLeaf) isOpen = true;
	});
	if (isOpen) temporaryLeaf.detach();
}

function readUnlockedMeldMarkdown(leaf: WorkspaceLeaf, sourceFile: TFile): string {
	const view = leaf.view as unknown as MeldEncryptedMarkdownView | null;
	if (!view || view.getViewType?.() !== MELD_ENCRYPTED_VIEW_TYPE || view.file?.path !== sourceFile.path) {
		throw new Error('Meld Encrypt did not open this file. Enable Meld Encrypt and try again.');
	}
	if (view.isSavingEnabled !== true || typeof view.getUnencryptedViewData !== 'function') {
		throw new Error('Unlock this Meld Encrypt note before converting it.');
	}
	return view.getUnencryptedViewData.call(view);
}
