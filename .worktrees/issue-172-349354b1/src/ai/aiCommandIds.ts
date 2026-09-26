export const AI_COMMAND_IDS = {
	capabilities: 'npde-ai-capabilities',
	describe: 'npde-ai-describe',
	apply: 'npde-ai-apply',
	validate: 'npde-ai-validate',
	save: 'npde-ai-save',
	undo: 'npde-ai-undo',
	redo: 'npde-ai-redo',
} as const;

export const AI_LEGACY_COMMAND_IDS = {
	capabilities: 'npde-ai-list-capabilities',
	describe: 'npde-ai-describe-document',
	apply: 'npde-ai-apply-operations',
} as const;

export const AI_COMMAND_ID_LIST = [
	AI_COMMAND_IDS.capabilities,
	AI_COMMAND_IDS.describe,
	AI_COMMAND_IDS.apply,
	AI_COMMAND_IDS.validate,
	AI_COMMAND_IDS.save,
	AI_COMMAND_IDS.undo,
	AI_COMMAND_IDS.redo,
	AI_LEGACY_COMMAND_IDS.capabilities,
	AI_LEGACY_COMMAND_IDS.describe,
	AI_LEGACY_COMMAND_IDS.apply,
] as const;
