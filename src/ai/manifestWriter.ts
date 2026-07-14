import type { DataAdapter } from 'obsidian';

import { buildCapabilityManifest } from './capabilities';
import { PLUGIN_ID } from './types';

export const AI_MANIFEST_RELATIVE_PATH = `plugins/${PLUGIN_ID}/ai/capabilities.json`;

export function getAiManifestPath(
	pluginDir: string | undefined,
	configDir?: string,
): string | null {
	const base = pluginDir || (configDir ? `${configDir}/plugins/${PLUGIN_ID}` : null);
	if (!base) return null;
	return `${base}/ai/capabilities.json`;
}

export async function writeCapabilitiesManifest(
	adapter: DataAdapter,
	pluginDir: string | undefined,
	pluginVersion: string,
	enabled: boolean,
	configDir?: string,
): Promise<string | null> {
	const manifestPath = getAiManifestPath(pluginDir, configDir);
	if (!manifestPath) return null;
	const manifest = buildCapabilityManifest({ pluginVersion, enabled });
	const directory = manifestPath.slice(0, manifestPath.lastIndexOf('/'));

	try {
		if (!(await adapter.exists(directory))) {
			await adapter.mkdir(directory);
		}
		await adapter.write(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
		return manifestPath;
	} catch {
		return null;
	}
}

export async function removeCapabilitiesManifest(
	adapter: DataAdapter,
	pluginDir: string | undefined,
	configDir?: string,
): Promise<void> {
	const manifestPath = getAiManifestPath(pluginDir, configDir);
	if (!manifestPath) return;
	try {
		if (await adapter.exists(manifestPath)) {
			await adapter.remove(manifestPath);
		}
	} catch {
		// Best-effort cleanup when AI interfacing is disabled.
	}
}
