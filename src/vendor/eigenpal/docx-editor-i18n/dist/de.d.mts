/**
 * @eigenpal/docx-editor-i18n/de
 *
 * German (`de`) — direct locale subpath for per-locale code-splitting.
 *
 * ```ts
 * // Static — bundler ships only this locale's strings
 * import de from '@eigenpal/docx-editor-i18n/de';
 *
 * // Dynamic — splits into its own chunk, loaded on demand
 * const de = (await import('@eigenpal/docx-editor-i18n/de')).default;
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
 * German (`de`) locale strings. Community-maintained; null leaves fall back to English.
 *
 * Identical content to the named `de` export from the package root;
 * this subpath just lets bundlers code-split it.
 *
 * @public
 */
declare const de: PartialLocaleStrings;

export { de, de as default };
