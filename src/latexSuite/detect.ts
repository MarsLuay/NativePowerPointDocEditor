import type { App } from 'obsidian';

export const LATEX_SUITE_PLUGIN_ID = 'obsidian-latex-suite';

export interface LatexSuitePluginLike {
	editorExtensions?: unknown[];
	api?: {
		plugin?: {
			keyboardEventPlugin?: unknown;
		};
	};
}

type PluginHost = {
	manifests?: Record<string, unknown>;
	plugins?: Record<string, LatexSuitePluginLike | undefined>;
	enabledPlugins?: Set<string>;
};

function pluginHost(app: App): PluginHost | null {
	return (app as App & { plugins?: PluginHost }).plugins ?? null;
}

export function isLatexSuiteInstalled(app: App): boolean {
	return Boolean(pluginHost(app)?.manifests?.[LATEX_SUITE_PLUGIN_ID]);
}

export function getLatexSuitePlugin(app: App): LatexSuitePluginLike | null {
	return pluginHost(app)?.plugins?.[LATEX_SUITE_PLUGIN_ID] ?? null;
}
