import { type TFile, normalizePath } from 'obsidian';

/** Trailing-slash folder prefix, or '' for the vault root. */
export function getVaultFolderPrefix(folderPath: string | null | undefined): string {
  return folderPath && folderPath !== '/' ? `${folderPath}/` : '';
}

/** Normalized path for `fileName` placed next to `file` (same folder). */
export function buildSiblingPath(file: TFile, fileName: string): string {
  return normalizePath(`${getVaultFolderPrefix(file.parent?.path)}${fileName}`);
}
