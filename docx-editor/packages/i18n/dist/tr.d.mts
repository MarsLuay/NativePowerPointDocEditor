/**
 * @npde/docx-editor-i18n/tr
 *
 * Turkish (`tr`) — direct locale subpath for per-locale code-splitting.
 *
 * ```ts
 * // Static — bundler ships only this locale's strings
 * import tr from '@npde/docx-editor-i18n/tr';
 *
 * // Dynamic — splits into its own chunk, loaded on demand
 * const tr = (await import('@npde/docx-editor-i18n/tr')).default;
 * ```
 *
 * For multi-locale apps, prefer the per-locale subpaths over importing
 * `locales` from the package root — `locales` pulls every locale into
 * the bundle.
 *
 * @packageDocumentation
 * @public
 */
import { PartialLocaleStrings } from './index.mjs';

/**
 * Turkish (`tr`) locale strings. Community-maintained; null leaves fall back to English.
 *
 * Identical content to the named `tr` export from the package root;
 * this subpath just lets bundlers code-split it.
 *
 * @public
 */
declare const tr: PartialLocaleStrings;

export { tr as default, tr };
