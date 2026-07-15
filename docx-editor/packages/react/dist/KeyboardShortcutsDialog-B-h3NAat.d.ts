import React__default, { CSSProperties } from 'react';
import { FindOptions, FindResult, FindMatch } from '@eigenpal/docx-editor-core/utils/findReplace';
import { ParsedClipboardContent } from '@eigenpal/docx-editor-core/utils';
import { TranslationKey } from '@eigenpal/docx-editor-i18n';

/**
 * Find and Replace Dialog Component
 *
 * Modal dialog for searching and replacing text in the document.
 * Supports find, find next/previous, replace, and replace all operations.
 *
 * Logic and utilities are in separate files:
 * - findReplaceUtils.ts — Pure search/replace functions and types
 * - useFindReplace.ts   — React hook for dialog state management
 */

/**
 * Props for the FindReplaceDialog component
 */
interface FindReplaceDialogProps {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Callback when dialog is closed */
    onClose: () => void;
    /** Callback when searching for text */
    onFind: (searchText: string, options: FindOptions) => FindResult | null;
    /** Callback when navigating to next match */
    onFindNext: () => FindMatch | null;
    /** Callback when navigating to previous match */
    onFindPrevious: () => FindMatch | null;
    /** Callback when replacing current match */
    onReplace: (replaceText: string) => boolean;
    /** Callback when replacing all matches */
    onReplaceAll: (searchText: string, replaceText: string, options: FindOptions) => number;
    /** Callback to highlight matches in document */
    onHighlightMatches?: (matches: FindMatch[]) => void;
    /** Callback to clear highlights */
    onClearHighlights?: () => void;
    /** Initial search text (e.g., from selected text) */
    initialSearchText?: string;
    /** Whether to start in replace mode */
    replaceMode?: boolean;
    /** Current match result (from external state) */
    currentResult?: FindResult | null;
    /** Additional CSS class */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
}
/**
 * FindReplaceDialog component - Modal for finding and replacing text
 */
declare function FindReplaceDialog({ isOpen, onClose, onFind, onFindNext, onFindPrevious, onReplace, onReplaceAll, onHighlightMatches, onClearHighlights, initialSearchText, replaceMode, currentResult, className, style, }: FindReplaceDialogProps): React__default.ReactElement | null;

/**
 * Hyperlink Dialog Component
 *
 * Modal dialog for inserting and editing hyperlinks in the document.
 * Supports both external URLs and internal bookmark links.
 *
 * Features:
 * - Input for URL (http, https, mailto, tel, etc.)
 * - Input for display text
 * - Edit existing hyperlinks
 * - Remove hyperlink option
 * - Internal bookmark selection
 * - Validation and error handling
 */

/**
 * Hyperlink data structure for dialog
 */
interface HyperlinkData {
    /** URL for external link */
    url?: string;
    /** Display text for the link */
    displayText?: string;
    /** Internal bookmark name */
    bookmark?: string;
    /** Tooltip text */
    tooltip?: string;
}
/**
 * Bookmark option for internal link selection
 */
interface BookmarkOption {
    /** Bookmark name/ID */
    name: string;
    /** Optional display label */
    label?: string;
}
/**
 * Props for the HyperlinkDialog component
 */
interface HyperlinkDialogProps {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Callback when dialog is closed */
    onClose: () => void;
    /** Callback when hyperlink is inserted/updated */
    onSubmit: (data: HyperlinkData) => void;
    /** Callback when hyperlink is removed */
    onRemove?: () => void;
    /** Initial data for editing existing hyperlink */
    initialData?: HyperlinkData;
    /** Currently selected text (used as default display text) */
    selectedText?: string;
    /** Whether we're editing an existing hyperlink */
    isEditing?: boolean;
    /** Available bookmarks for internal links */
    bookmarks?: BookmarkOption[];
    /** Additional CSS class */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
}
/**
 * Validate a URL string
 * Supports http, https, mailto, tel, ftp protocols
 */
declare function isValidUrl(url: string): boolean;
/**
 * Normalize a URL by adding protocol if needed
 */
declare function normalizeUrl(url: string): string;
/**
 * Detect URL type from string
 */
declare function getUrlType(url: string): 'web' | 'email' | 'phone' | 'ftp' | 'unknown';
/**
 * HyperlinkDialog component - Modal for inserting/editing hyperlinks
 */
declare function HyperlinkDialog({ isOpen, onClose, onSubmit, onRemove, initialData, selectedText, isEditing, bookmarks, className, style, }: HyperlinkDialogProps): React__default.ReactElement | null;
/**
 * Create HyperlinkData from a URL string
 */
declare function createHyperlinkData(url: string, displayText?: string): HyperlinkData;
/**
 * Create HyperlinkData for an internal bookmark
 */
declare function createBookmarkLinkData(bookmark: string, displayText?: string): HyperlinkData;
/**
 * Check if HyperlinkData is for an external URL
 */
declare function isExternalHyperlinkData(data: HyperlinkData): boolean;
/**
 * Check if HyperlinkData is for an internal bookmark
 */
declare function isBookmarkHyperlinkData(data: HyperlinkData): boolean;
/**
 * Get display text from HyperlinkData, falling back to URL/bookmark
 */
declare function getDisplayText(data: HyperlinkData): string;
/**
 * Convert email address to mailto: link
 */
declare function emailToMailto(email: string): string;
/**
 * Convert phone number to tel: link
 */
declare function phoneToTel(phone: string): string;
/**
 * Extract bookmarks from document for the dialog
 */
declare function extractBookmarksForDialog(bookmarks: {
    name: string;
    id: number;
}[]): BookmarkOption[];
/**
 * Hook state for the Hyperlink dialog
 */
interface UseHyperlinkDialogState {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Initial data for the dialog (for editing) */
    initialData?: HyperlinkData;
    /** Currently selected text */
    selectedText?: string;
    /** Whether we're editing an existing hyperlink */
    isEditing: boolean;
}
/**
 * Hook return type for the Hyperlink dialog
 */
interface UseHyperlinkDialogReturn {
    /** Current state */
    state: UseHyperlinkDialogState;
    /** Open dialog for inserting new hyperlink */
    openInsert: (selectedText?: string) => void;
    /** Open dialog for editing existing hyperlink */
    openEdit: (data: HyperlinkData) => void;
    /** Close the dialog */
    close: () => void;
    /** Toggle dialog open/closed */
    toggle: () => void;
}
/**
 * Hook for managing Hyperlink dialog state
 */
declare function useHyperlinkDialog(): UseHyperlinkDialogReturn;

/**
 * Paste Special Dialog Component
 *
 * Provides paste options for pasting content with or without formatting.
 * Features:
 * - Paste with formatting (default)
 * - Paste as plain text (unformatted)
 * - Keyboard shortcut: Ctrl+Shift+V opens dialog
 */

/**
 * Paste option type
 */
type PasteOption = 'formatted' | 'plainText';
/**
 * Paste special dialog props
 */
interface PasteSpecialDialogProps {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Callback when dialog is closed */
    onClose: () => void;
    /** Callback when paste is confirmed */
    onPaste: (content: ParsedClipboardContent, asPlainText: boolean) => void;
    /** Optional custom position */
    position?: {
        x: number;
        y: number;
    };
    /** Additional className */
    className?: string;
}
/**
 * Paste option definition with translation keys
 */
interface PasteOptionDef {
    id: PasteOption;
    labelKey: TranslationKey;
    descriptionKey: TranslationKey;
    shortcutKey: TranslationKey;
}
/**
 * Hook return value for paste special
 */
interface UsePasteSpecialReturn {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Open the paste special dialog */
    openDialog: () => void;
    /** Close the dialog */
    closeDialog: () => void;
    /** Handle keyboard shortcut (Ctrl+Shift+V) */
    handleKeyDown: (event: KeyboardEvent) => boolean;
    /** Paste as plain text directly */
    pasteAsPlainText: () => Promise<void>;
}
/**
 * Options for usePasteSpecial hook
 */
interface UsePasteSpecialOptions {
    /** Callback when paste is confirmed */
    onPaste?: (content: ParsedClipboardContent, asPlainText: boolean) => void;
    /** Whether paste operations are enabled */
    enabled?: boolean;
}
declare const PasteSpecialDialog: React__default.FC<PasteSpecialDialogProps>;
/**
 * Hook to manage paste special dialog
 */
declare function usePasteSpecial(options?: UsePasteSpecialOptions): UsePasteSpecialReturn;
/**
 * Get paste option definition by id
 */
declare function getPasteOption(id: PasteOption): PasteOptionDef | undefined;
/**
 * Get all paste option definitions
 */
declare function getAllPasteOptions(): PasteOptionDef[];
/**
 * Get default paste option
 */
declare function getDefaultPasteOption(): PasteOption;
/**
 * Check if paste special shortcut
 */
declare function isPasteSpecialShortcut(event: KeyboardEvent): boolean;

/**
 * Keyboard-shortcut catalog — categorized list of every shortcut surfaced
 * in the KeyboardShortcutsDialog, plus the lookup helpers that filter and
 * label them.
 */

/**
 * Get all default shortcuts
 */
declare function getDefaultShortcuts(): KeyboardShortcut[];
/**
 * Get shortcuts by category
 */
declare function getShortcutsByCategory(category: ShortcutCategory): KeyboardShortcut[];
/**
 * Get common/frequently used shortcuts
 */
declare function getCommonShortcuts(): KeyboardShortcut[];
/**
 * Get category label translation key
 */
declare function getCategoryLabel(category: ShortcutCategory): string;
/**
 * Get all categories
 */
declare function getAllCategories(): ShortcutCategory[];

/**
 * Single-row renderer for a keyboard shortcut inside the dialog, plus the
 * platform-aware key formatter (Ctrl→⌘ on Mac). The formatter is exposed
 * as `formatShortcutKeys` for callers that render kbd badges outside the
 * dialog.
 */

/**
 * Format key combination for current platform
 */
declare function formatKeys(keys: string): string;

/**
 * Keyboard Shortcuts Dialog Component
 *
 * Displays all available keyboard shortcuts organized by category.
 * Features:
 * - Categorized shortcut list
 * - Search/filter functionality
 * - Platform-aware modifier keys (Ctrl/Cmd)
 * - Keyboard shortcut to open (Ctrl+/)
 *
 * The shortcut catalog and category metadata live in
 * `KeyboardShortcutsDialog/data.ts`. The per-row renderer and the
 * platform-aware key formatter live in `KeyboardShortcutsDialog/ShortcutItem.tsx`.
 */

/**
 * Keyboard shortcut definition
 */
interface KeyboardShortcut {
    /** Unique identifier */
    id: string;
    /** Display name */
    name: string;
    /** Description of what the shortcut does */
    description: string;
    /** Primary key combination (e.g., 'Ctrl+C') */
    keys: string;
    /** Alternative key combination */
    altKeys?: string;
    /** Category for grouping */
    category: ShortcutCategory;
    /** Whether this is a common/frequently used shortcut */
    common?: boolean;
    /** Translation key for display name (used internally) */
    nameKey?: TranslationKey;
    /** Translation key for description (used internally) */
    descriptionKey?: TranslationKey;
}
/**
 * Shortcut category
 */
type ShortcutCategory = 'editing' | 'formatting' | 'navigation' | 'clipboard' | 'selection' | 'view' | 'file' | 'other';
/**
 * Dialog props
 */
interface KeyboardShortcutsDialogProps {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Close callback */
    onClose: () => void;
    /** Custom shortcuts (merged with defaults) */
    customShortcuts?: KeyboardShortcut[];
    /** Whether to show search */
    showSearch?: boolean;
    /** Additional className */
    className?: string;
}
/**
 * Hook options
 */
interface UseKeyboardShortcutsDialogOptions {
    /** Whether the dialog can be opened with Ctrl+? or F1 */
    enabled?: boolean;
    /** Custom open shortcut (default: Ctrl+/) */
    openShortcut?: string;
}
/**
 * Hook return value
 */
interface UseKeyboardShortcutsDialogReturn {
    /** Whether dialog is open */
    isOpen: boolean;
    /** Open the dialog */
    open: () => void;
    /** Close the dialog */
    close: () => void;
    /** Toggle the dialog */
    toggle: () => void;
    /** Keyboard event handler */
    handleKeyDown: (event: KeyboardEvent) => void;
}
declare const KeyboardShortcutsDialog: React__default.FC<KeyboardShortcutsDialogProps>;
/**
 * Hook to manage keyboard shortcuts dialog
 */
declare function useKeyboardShortcutsDialog(options?: UseKeyboardShortcutsDialogOptions): UseKeyboardShortcutsDialogReturn;

export { extractBookmarksForDialog as A, type BookmarkOption as B, getDisplayText as C, getUrlType as D, isBookmarkHyperlinkData as E, FindReplaceDialog as F, isExternalHyperlinkData as G, type HyperlinkData as H, isValidUrl as I, normalizeUrl as J, type KeyboardShortcut as K, phoneToTel as L, type PasteOption as P, type ShortcutCategory as S, type UseKeyboardShortcutsDialogOptions as U, type FindReplaceDialogProps as a, HyperlinkDialog as b, type HyperlinkDialogProps as c, KeyboardShortcutsDialog as d, type KeyboardShortcutsDialogProps as e, PasteSpecialDialog as f, type PasteSpecialDialogProps as g, type UseKeyboardShortcutsDialogReturn as h, type UsePasteSpecialOptions as i, type UsePasteSpecialReturn as j, formatKeys as k, getAllCategories as l, getAllPasteOptions as m, getCategoryLabel as n, getCommonShortcuts as o, getDefaultPasteOption as p, getDefaultShortcuts as q, getPasteOption as r, getShortcutsByCategory as s, isPasteSpecialShortcut as t, useHyperlinkDialog as u, useKeyboardShortcutsDialog as v, usePasteSpecial as w, createBookmarkLinkData as x, createHyperlinkData as y, emailToMailto as z };
