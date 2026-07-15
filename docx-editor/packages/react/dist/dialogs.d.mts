/**
 * @npde/docx-editor-react/dialogs
 *
 * Modal dialogs for hyperlinks, find/replace, paste-special, page setup,
 * and keyboard shortcuts. Pair with the hooks under
 * `@npde/docx-editor-react/hooks` when wiring custom UI.
 *
 * @example
 * ```tsx
 * import { HyperlinkDialog, FindReplaceDialog } from '@npde/docx-editor-react/dialogs';
 * ```
 *
 * @packageDocumentation
 * @public
 */
export { B as BookmarkOption, F as FindReplaceDialog, a as FindReplaceDialogProps, H as HyperlinkData, b as HyperlinkDialog, c as HyperlinkDialogProps, K as KeyboardShortcut, d as KeyboardShortcutsDialog, e as KeyboardShortcutsDialogProps, P as PasteOption, f as PasteSpecialDialog, g as PasteSpecialDialogProps, S as ShortcutCategory, U as UseKeyboardShortcutsDialogOptions, h as UseKeyboardShortcutsDialogReturn, i as UsePasteSpecialOptions, j as UsePasteSpecialReturn, x as createBookmarkLinkData, y as createHyperlinkData, z as emailToMailto, A as extractBookmarksForDialog, k as formatShortcutKeys, l as getAllCategories, m as getAllPasteOptions, n as getCategoryLabel, o as getCommonShortcuts, p as getDefaultPasteOption, q as getDefaultShortcuts, C as getDisplayText, r as getPasteOption, s as getShortcutsByCategory, D as getUrlType, E as isBookmarkHyperlinkData, G as isExternalHyperlinkData, t as isPasteSpecialShortcut, I as isValidUrl, J as normalizeUrl, L as phoneToTel, v as useKeyboardShortcutsDialog, w as usePasteSpecial } from './KeyboardShortcutsDialog-DDN-wlZe.mjs';
import React__default from 'react';
import { SectionProperties } from '@npde/docx-editor-core/types/document';
export { FindMatch, FindOptions, FindResult, HighlightOptions, createDefaultFindOptions, createSearchPattern, escapeRegexString, findAllMatches, getDefaultHighlightOptions, getMatchCountText, isEmptySearch, replaceAllInContent, replaceFirstInContent } from '@npde/docx-editor-core/utils/findReplace';
import '@npde/docx-editor-core/utils';
import '@npde/docx-editor-i18n';

/**
 * Page Setup Dialog
 *
 * Modal for editing page layout properties:
 * - Page size (Letter, A4, Legal, etc.)
 * - Orientation (portrait/landscape)
 * - Margins (top, bottom, left, right) in inches
 */

interface PageSetupDialogProps {
    isOpen: boolean;
    onClose: () => void;
    onApply: (props: Partial<SectionProperties>) => void;
    currentProps?: SectionProperties;
}
declare function PageSetupDialog({ isOpen, onClose, onApply, currentProps, }: PageSetupDialogProps): React__default.ReactElement | null;

export { PageSetupDialog, type PageSetupDialogProps };
