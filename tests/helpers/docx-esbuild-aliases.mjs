import {
	createDocxEditorAliases,
	resolveDocxEditorPackagesRoot,
} from '../../scripts/lib/docx-editor-aliases.mjs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

/** Shared esbuild `alias` map for vendored `@npde/docx-editor-*` packages. */
export const docxEditorAliases = await createDocxEditorAliases(
	resolveDocxEditorPackagesRoot(projectRoot),
	projectRoot,
);

export { projectRoot };
