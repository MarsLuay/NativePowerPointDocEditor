import type { DataAdapter } from 'obsidian';

import { NPDE_AI_SKILL, NPDE_AI_SKILL_PATH } from './npdeSkill';

const NPDE_AI_SKILL_DIRECTORIES = ['.agents', '.agents/skills', '.agents/skills/npde'];

/** Add the bundled NPDE skill without overwriting an existing user-authored skill. */
export async function writeAiSkill(adapter: DataAdapter): Promise<string | null> {
	try {
		for (const directory of NPDE_AI_SKILL_DIRECTORIES) {
			if (!(await adapter.exists(directory))) {
				await adapter.mkdir(directory);
			}
		}

		if (!(await adapter.exists(NPDE_AI_SKILL_PATH))) {
			await adapter.write(NPDE_AI_SKILL_PATH, NPDE_AI_SKILL);
		}
		return NPDE_AI_SKILL_PATH;
	} catch {
		return null;
	}
}
