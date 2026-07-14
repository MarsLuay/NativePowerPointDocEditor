import type { App, TFolder } from 'obsidian';
import { TFile, normalizePath } from 'obsidian';

import { getAvailableNumberedPath } from '../export/artifactPaths';
import { errorLog, infoLog } from '../logger';
import { getVaultFolderPrefix } from './paths';
import { buildBlankDocxArrayBuffer, buildBlankPptxArrayBuffer } from './blankOfficePackages';

export type NewOfficeFileKind = 'docx' | 'pptx';

const DEFAULT_BASENAME: Record<NewOfficeFileKind, string> = {
	docx: 'Untitled',
	pptx: 'Untitled',
};

export function buildNewOfficeFileCandidatePath(folder: TFolder, kind: NewOfficeFileKind): string {
	return normalizePath(`${getVaultFolderPrefix(folder.path)}${DEFAULT_BASENAME[kind]}.${kind}`);
}

export function resolveAvailableOfficeFilePath(
	app: App,
	candidatePath: string,
): string {
	const exists = (path: string) => app.vault.getAbstractFileByPath(path) != null;
	if (!exists(candidatePath)) {
		return candidatePath;
	}
	return getAvailableNumberedPath(candidatePath, exists);
}

/**
 * Create a blank DOCX/PPTX in `folder`, then open it in a new leaf.
 * Avoids overwriting existing Untitled.* files by numbering keep-both paths.
 */
export async function createAndOpenNewOfficeFile(
	app: App,
	folder: TFolder,
	kind: NewOfficeFileKind,
): Promise<TFile> {
	const candidatePath = buildNewOfficeFileCandidatePath(folder, kind);
	const path = resolveAvailableOfficeFilePath(app, candidatePath);
	const buffer = kind === 'docx'
		? await buildBlankDocxArrayBuffer()
		: await buildBlankPptxArrayBuffer();

	infoLog('file', `Creating blank ${kind.toUpperCase()}`, { path, folder: folder.path });
	const file = await app.vault.createBinary(path, buffer);
	try {
		await app.workspace.getLeaf('tab').openFile(file, { active: true });
	} catch (error) {
		errorLog('file', `Created ${kind.toUpperCase()} but failed to open`, {
			path: file.path,
			error: error instanceof Error ? error.message : String(error),
		});
		throw error;
	}
	return file;
}
