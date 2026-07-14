/**
 * Shared find/replace shell conventions for the DOCX (React dialog) and PPTX
 * (imperative panel) editors. This owns only the labels/state conventions —
 * mode vocabulary, result-status wording rules, and match-index navigation.
 * The actual searching, highlighting, and replacement stay per-renderer.
 */

export type FindReplaceMode = 'find' | 'replace';

export interface FindResultStatusLabels {
  /** Shown when the query is empty (pass '' to render nothing). */
  noSearch: string;
  /** Shown when the query has no matches. */
  noMatches: string;
  /** Shown otherwise; `current` is 1-based. */
  resultCount: (current: number, total: number) => string;
}

/**
 * The shared "no search / no matches / N of M" result convention. `currentIndex`
 * is 0-based; callers supply the localized label surface.
 */
export function formatFindResultStatus(
  query: string,
  currentIndex: number,
  matchCount: number,
  labels: FindResultStatusLabels,
): string {
  if (!query.trim()) {
    return labels.noSearch;
  }
  if (matchCount <= 0) {
    return labels.noMatches;
  }
  return labels.resultCount(currentIndex + 1, matchCount);
}

/**
 * Wrap-around next/previous navigation shared by both shells. `direction` is
 * typically -1 or 1. Returns 0 when there are no matches.
 */
export function wrapMatchIndex(currentIndex: number, direction: number, matchCount: number): number {
  if (matchCount <= 0) {
    return 0;
  }
  return (currentIndex + direction + matchCount) % matchCount;
}
