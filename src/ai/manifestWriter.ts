import type { DataAdapter } from 'obsidian';

import { buildCapabilityManifest } from './capabilities';
import { PLUGIN_ID } from './types';

export const AI_MANIFEST_RELATIVE_PATH = `plugins/${PLUGIN_ID}/ai/capabilities.json`;

export function getAiManifestPath(pluginDir: string | undefined, configDir = '.obsidian'): string {
	const base = pluginDir || `${configDir}/plugins/${PLUGIN_ID}`;
	return `${base}/ai/capabilities.json`;
}

export async function writeCapabilitiesManifest(
	adapter: DataAdapter,
	pluginDir: string | undefined,
	pluginVersion: string,
	enabled: boolean,
): Promise<string | null> {
	const manifestPath = getAiManifestPath(pluginDir);
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
): Promise<void> {
	const manifestPath = getAiManifestPath(pluginDir);
	try {
		if (await adapter.exists(manifestPath)) {
			await adapter.remove(manifestPath);
		}
	} catch {
		// Best-effort cleanup when AI interfacing is disabled.
	}
}
