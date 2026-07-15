import { PartialLocaleStrings } from './index.mjs';

/**
 * French (`fr`) locale strings. Community-maintained; null leaves fall back to English.
 *
 * Identical content to the named `fr` export from the package root;
 * this subpath just lets bundlers code-split it.
 *
 * @public
 */
declare const fr: PartialLocaleStrings;

export { fr as default, fr };
