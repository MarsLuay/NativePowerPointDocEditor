/**
 * Find & Replace Utility Functions
 *
 * Pure utility functions for text search, pattern matching, and
 * document search. Lifted from packages/react/src/components/dialogs/
 * findReplaceUtils.ts so the React adapter and the Vue adapter share
 * one implementation.
 * @packageDocumentation
 * @public
 */
/**
 * A single match result in the document
 */
interface FindMatch {
    /** Index of the paragraph containing the match */
    paragraphIndex: number;
    /** Index of the run/content within the paragraph */
    contentIndex: number;
    /** Character offset within the content */
    startOffset: number;
    /** Character offset for end of match */
    endOffset: number;
    /** The matched text */
    text: string;
}
/**
 * Find options for controlling search behavior
 */
interface FindOptions {
    /** Whether to match case */
    matchCase: boolean;
    /** Whether to match whole words only */
    matchWholeWord: boolean;
    /** Whether to use regular expressions (future) */
    useRegex?: boolean;
}
/**
 * Find result with all matches
 */
interface FindResult {
    /** All matches found */
    matches: FindMatch[];
    /** Total match count */
    totalCount: number;
    /** Current match index (0-based) */
    currentIndex: number;
}
/**
 * Highlight options for document rendering
 */
interface HighlightOptions {
    /** Background color for current match */
    currentMatchColor: string;
    /** Background color for other matches */
    otherMatchColor: string;
}
/**
 * Create default find options
 */
declare function createDefaultFindOptions(): FindOptions;
/**
 * Escape string for use in regex pattern
 */
declare function escapeRegexString(str: string): string;
/**
 * Create a regex pattern from search text and options
 */
declare function createSearchPattern(searchText: string, options: FindOptions): RegExp | null;
/**
 * Find all matches of search text in content
 */
declare function findAllMatches(content: string, searchText: string, options: FindOptions): Array<{
    start: number;
    end: number;
}>;
/**
 * Replace text in content
 */
declare function replaceAllInContent(content: string, searchText: string, replaceText: string, options: FindOptions): string;
/**
 * Replace first match in content
 */
declare function replaceFirstInContent(content: string, searchText: string, replaceText: string, options: FindOptions, startIndex?: number): {
    content: string;
    replaced: boolean;
    matchStart: number;
    matchEnd: number;
};
/**
 * Get match count for status display
 */
declare function getMatchCountText(result: FindResult | null): string;
/**
 * Check if search text is empty or whitespace-only
 */
declare function isEmptySearch(searchText: string): boolean;
/**
 * Get default highlight options
 */
declare function getDefaultHighlightOptions(): HighlightOptions;
/**
 * Find all matches in a document
 */
declare function findInDocument(document: any, searchText: string, options: FindOptions): FindMatch[];
/**
 * Find matches in a single paragraph
 */
declare function findInParagraph(paragraph: any, searchText: string, options: FindOptions, paragraphIndex: number): FindMatch[];
/**
 * Scroll to a match in the document
 */
declare function scrollToMatch(containerElement: HTMLElement | null, match: FindMatch): void;

export { type FindMatch, type FindOptions, type FindResult, type HighlightOptions, createDefaultFindOptions, createSearchPattern, escapeRegexString, findAllMatches, findInDocument, findInParagraph, getDefaultHighlightOptions, getMatchCountText, isEmptySearch, replaceAllInContent, replaceFirstInContent, scrollToMatch };
