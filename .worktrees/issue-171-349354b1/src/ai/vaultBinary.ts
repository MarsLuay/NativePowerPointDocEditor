import type { Vault } from 'obsidian';
import { TFile } from 'obsidian';
import { AI_ERROR_CODES, createAiError } from './errors';

export async function readVaultBinaryFile(
	vault: Vault,
	vaultPath: string,
): Promise<{ bytes: Uint8Array; extension: string }> {
	const file = vault.getAbstractFileByPath(vaultPath);
	if (!(file instanceof TFile)) {
		throw createAiError(AI_ERROR_CODES.FILE_NOT_FOUND, `Vault file not found: ${vaultPath}.`, {
			path: vaultPath,
		});
	}
	const bytes = new Uint8Array(await vault.readBinary(file));
	return { bytes, extension: file.extension };
}
