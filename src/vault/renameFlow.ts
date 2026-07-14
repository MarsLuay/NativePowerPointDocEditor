import type { App, TFile } from 'obsidian';

import { buildSiblingPath } from './paths';

export interface RenameResult {
  renamed: boolean;
  path: string;
}

/**
 * Shared title-rename step for both editors: resolve `fileName` to a sibling
 * path in the file's own folder and rename via `fileManager`, no-opping when the
 * target path is unchanged. Name normalization/sanitization, empty-name guards,
 * and success/error feedback stay with each host's rename UX.
 */
export async function renameFileToSiblingName(app: App, file: TFile, fileName: string): Promise<RenameResult> {
  const path = buildSiblingPath(file, fileName);
  if (path === file.path) {
    return { renamed: false, path };
  }
  await app.fileManager.renameFile(file, path);
  return { renamed: true, path };
}
