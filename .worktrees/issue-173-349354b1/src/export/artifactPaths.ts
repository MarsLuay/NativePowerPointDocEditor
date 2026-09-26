import { TFile, type Vault, normalizePath } from 'obsidian';

/**
 * Shared naming and overwrite/keep-both plumbing for editor export artifacts.
 * The renderer-specific bits (how bytes are produced, and the menu vs modal used
 * to ask the user) stay in each view; this module owns path derivation, the
 * numbered "keep both" fallback, conflict resolution, and the vault write.
 */

export { getVaultFolderPrefix } from '../vault/paths';

export type ArtifactConflictChoice = 'replace' | 'keep-both' | 'cancel';

export interface ArtifactWriteTarget {
  path: string;
  existingFile: TFile | null;
  replace: boolean;
}

const INVALID_ARTIFACT_NAME_CHARS = /[\\/:*?"<>|]/g;

export function sanitizeArtifactBaseName(name: string, fallback: string): string {
  return name.replace(INVALID_ARTIFACT_NAME_CHARS, '_') || fallback;
}

/**
 * Finds the first free `<base> N<ext>` path (N starting at 2), falling back to a
 * timestamp suffix. `exists` is injected so the caller supplies the vault lookup.
 */
export function getAvailableNumberedPath(path: string, exists: (candidate: string) => boolean): string {
  const lastSlashIndex = path.lastIndexOf('/');
  const folderPrefix = lastSlashIndex >= 0 ? `${path.slice(0, lastSlashIndex)}/` : '';
  const fileName = lastSlashIndex >= 0 ? path.slice(lastSlashIndex + 1) : path;
  const extensionIndex = fileName.lastIndexOf('.');
  const baseName = extensionIndex > 0 ? fileName.slice(0, extensionIndex) : fileName;
  const extension = extensionIndex > 0 ? fileName.slice(extensionIndex) : '';

  for (let index = 2; index < 1000; index += 1) {
    const candidatePath = normalizePath(`${folderPrefix}${baseName} ${index}${extension}`);
    if (!exists(candidatePath)) {
      return candidatePath;
    }
  }

  return normalizePath(`${folderPrefix}${baseName} ${Date.now()}${extension}`);
}

/**
 * Turns a user's conflict choice into a concrete write target. `keep-both`
 * resolves to a numbered sibling; `cancel` returns null.
 */
export function resolveArtifactConflict(
  requestedPath: string,
  existingFile: TFile | null,
  choice: ArtifactConflictChoice,
  exists: (candidate: string) => boolean,
): ArtifactWriteTarget | null {
  if (choice === 'cancel') {
    return null;
  }
  if (choice === 'replace') {
    return { path: requestedPath, existingFile, replace: true };
  }
  return { path: getAvailableNumberedPath(requestedPath, exists), existingFile: null, replace: false };
}

export async function writeVaultBinaryArtifact(
  vault: Vault,
  target: ArtifactWriteTarget,
  data: ArrayBuffer,
): Promise<{ path: string }> {
  if (target.existingFile) {
    await vault.modifyBinary(target.existingFile, data);
    return { path: target.path };
  }

  // Obsidian writes the artifact before resolving createBinary(). Some vault
  // adapters resolve that promise with null, so callers must report the target
  // path they requested rather than depending on an immediately indexed TFile.
  await vault.createBinary(target.path, data);
  return { path: target.path };
}
