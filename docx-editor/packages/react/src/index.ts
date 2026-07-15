/**
 * @npde/docx-editor-react
 *
 * Curated root entry for the documented React editor API. Advanced surfaces
 * stay public through explicit subpaths:
 * - `@npde/docx-editor-react/ui`
 * - `@npde/docx-editor-react/dialogs`
 * - `@npde/docx-editor-react/hooks`
 * - `@npde/docx-editor-react/plugin-api`
 *
 * Framework-agnostic document utilities live in `@npde/docx-editor-core`.
 *
 * @packageDocumentation
 * @public
 */

export const VERSION = '0.0.2';

// Main editor contract
export {
  DocxEditor,
  type DocxEditorProps,
  type DocxEditorRef,
  type EditorMode,
} from './components/DocxEditor';
export { renderAsync, type RenderAsyncOptions, type DocxEditorHandle } from './renderAsync';

// Document factory helpers — re-exported from `@npde/docx-editor-core` so
// the common "spawn a blank editor" affordance is available without forcing
// consumers to add `-core` to their dependency tree alongside `-react`.
export {
  createEmptyDocument,
  createDocumentWithText,
  type CreateEmptyDocumentOptions,
} from '@npde/docx-editor-core';

// i18n contract — runtime only. Locale string types (LocaleStrings,
// Translations, PartialLocaleStrings, TranslationKey) live in
// `@npde/docx-editor-i18n`; import them from there.
export { LocaleProvider, useTranslation, type LocaleProviderProps } from './i18n';
