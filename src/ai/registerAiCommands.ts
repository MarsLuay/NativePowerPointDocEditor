import type { Plugin } from 'obsidian';
import { showI18nNotice } from '../i18n/notify';
import type { PluginI18nService } from '../i18n/I18nService';
import type { NpdeAiApi } from './pluginApi';
import {
	copyJsonToClipboard,
	getActiveDocumentPath,
	parseApplyRequest,
	parseDescribeRequest,
	parseRedoRequest,
	parseSaveRequest,
	parseUndoRequest,
	parseValidateRequest,
	readClipboardJson,
	resolveDocumentPath,
} from './commandProtocol';

import { AI_COMMAND_IDS, AI_LEGACY_COMMAND_IDS } from './aiCommandIds';

interface RegisterAiCommandsOptions {
	plugin: Plugin;
	getI18n: () => PluginI18nService | null;
	getAi: () => NpdeAiApi | undefined;
}

function requireAi(
	getAi: () => NpdeAiApi | undefined,
	getI18n: () => PluginI18nService | null,
): NpdeAiApi | null {
	const ai = getAi();
	if (!ai) {
		showI18nNotice(getI18n(), 'settings:ai.disabledNotice');
		return null;
	}
	return ai;
}

function resolvePathFromClipboard(
	plugin: Plugin,
	getI18n: () => PluginI18nService | null,
	requestPath: string | undefined,
): string | null {
	const resolved = resolveDocumentPath(
		requestPath,
		getActiveDocumentPath(plugin.app.workspace.getActiveFile()),
	);
	if (!resolved) {
		showI18nNotice(getI18n(), 'settings:ai.clipboardPathRequired');
		return null;
	}
	return resolved;
}

export function registerAiCommands(options: RegisterAiCommandsOptions): void {
	const { plugin, getI18n, getAi } = options;

	const registerCapabilities = (id: string, name: string) => {
		plugin.addCommand({
			id,
			name,
			callback: async () => {
				const ai = requireAi(getAi, getI18n);
				if (!ai) return;
				const manifest = ai.listCapabilities();
				await copyJsonToClipboard(manifest);
				showI18nNotice(getI18n(), 'settings:ai.manifestCopied');
			},
		});
	};

	const registerDescribe = (id: string, name: string) => {
		plugin.addCommand({
			id,
			name,
			callback: async () => {
				const ai = requireAi(getAi, getI18n);
				if (!ai) return;
				try {
					const payload = parseDescribeRequest(await readClipboardJson());
					const path = resolvePathFromClipboard(plugin, getI18n, payload.path);
					if (!path) return;
					const result = await ai.describe(path);
					await copyJsonToClipboard(result);
					showI18nNotice(getI18n(), 'settings:ai.describeCopied');
				} catch {
					showI18nNotice(getI18n(), 'settings:ai.clipboardInvalid');
				}
			},
		});
	};

	const registerApply = (id: string, name: string) => {
		plugin.addCommand({
			id,
			name,
			callback: async () => {
				const ai = requireAi(getAi, getI18n);
				if (!ai) return;
				try {
					const payload = parseApplyRequest(await readClipboardJson());
					const path = resolvePathFromClipboard(plugin, getI18n, payload.path);
					if (!path || !payload.ops) return;
					const result = await ai.apply(path, payload.ops, { dryRun: payload.dryRun === true });
					await copyJsonToClipboard(result);
					showI18nNotice(getI18n(), 'settings:ai.applyCopied');
				} catch {
					showI18nNotice(getI18n(), 'settings:ai.clipboardInvalid');
				}
			},
		});
	};

	registerCapabilities(AI_COMMAND_IDS.capabilities, 'AI: Copy capabilities (JSON)');
	registerDescribe(AI_COMMAND_IDS.describe, 'AI: Describe document (clipboard JSON in/out)');
	registerApply(AI_COMMAND_IDS.apply, 'AI: Apply operations (clipboard JSON in/out)');

	plugin.addCommand({
		id: AI_COMMAND_IDS.validate,
		name: 'AI: Validate operations (clipboard JSON in/out)',
		callback: async () => {
			const ai = requireAi(getAi, getI18n);
			if (!ai) return;
			try {
				const payload = parseValidateRequest(await readClipboardJson());
				const result = ai.validateOps(payload.ops ?? []);
				await copyJsonToClipboard(result);
				showI18nNotice(getI18n(), 'settings:ai.validateCopied');
			} catch {
				showI18nNotice(getI18n(), 'settings:ai.clipboardInvalid');
			}
		},
	});

	plugin.addCommand({
		id: AI_COMMAND_IDS.save,
		name: 'AI: Save document (clipboard JSON in/out)',
		callback: async () => {
			const ai = requireAi(getAi, getI18n);
			if (!ai) return;
			try {
				const payload = parseSaveRequest(await readClipboardJson());
				const path = resolvePathFromClipboard(plugin, getI18n, payload.path);
				if (!path) return;
				const session = await ai.openSession(path);
				const result = await session.save();
				await session.close();
				await copyJsonToClipboard(result);
				showI18nNotice(getI18n(), 'settings:ai.saveCopied');
			} catch {
				showI18nNotice(getI18n(), 'settings:ai.clipboardInvalid');
			}
		},
	});

	plugin.addCommand({
		id: AI_COMMAND_IDS.undo,
		name: 'AI: Undo agent edit (clipboard JSON in/out)',
		callback: async () => {
			const ai = requireAi(getAi, getI18n);
			if (!ai) return;
			try {
				const payload = parseUndoRequest(await readClipboardJson());
				const path = resolvePathFromClipboard(plugin, getI18n, payload.path);
				if (!path) return;
				const result = await ai.undo(path);
				await copyJsonToClipboard(result);
				showI18nNotice(getI18n(), 'settings:ai.undoCopied');
			} catch {
				showI18nNotice(getI18n(), 'settings:ai.clipboardInvalid');
			}
		},
	});

	plugin.addCommand({
		id: AI_COMMAND_IDS.redo,
		name: 'AI: Redo agent edit (clipboard JSON in/out)',
		callback: async () => {
			const ai = requireAi(getAi, getI18n);
			if (!ai) return;
			try {
				const payload = parseRedoRequest(await readClipboardJson());
				const path = resolvePathFromClipboard(plugin, getI18n, payload.path);
				if (!path) return;
				const result = await ai.redo(path);
				await copyJsonToClipboard(result);
				showI18nNotice(getI18n(), 'settings:ai.redoCopied');
			} catch {
				showI18nNotice(getI18n(), 'settings:ai.clipboardInvalid');
			}
		},
	});

	registerCapabilities(AI_LEGACY_COMMAND_IDS.capabilities, 'AI: Copy capability manifest (JSON, legacy id)');
	registerDescribe(AI_LEGACY_COMMAND_IDS.describe, 'AI: Describe document from clipboard JSON (legacy id)');
	registerApply(AI_LEGACY_COMMAND_IDS.apply, 'AI: Apply operations from clipboard JSON (legacy id)');
}
