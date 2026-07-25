/**
 * Keep paragraph `_originalFormatting` in sync with live PM attrs.
 *
 * Serialize prefers `_originalFormatting` when present and only overlays a few
 * fields (alignment/numPr/styleId/…). Spacing/indent edits that only update
 * PM attrs were lost on save+reopen. Syncing on write fixes that without
 * inlining style-resolved PM display values for untouched paragraphs.
 */

import type { ParagraphFormatting } from '../types/document';

/** PM attr keys that map 1:1 onto ParagraphFormatting / direct pPr. */
export const ORIGINAL_FORMATTING_SYNC_KEYS = [
  'alignment',
  'bidi',
  'spaceBefore',
  'spaceAfter',
  'lineSpacing',
  'lineSpacingRule',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'hangingIndent',
  'borders',
  'shading',
  'tabs',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'contextualSpacing',
  'outlineLevel',
  'numPr',
  'styleId',
] as const;

type SyncKey = (typeof ORIGINAL_FORMATTING_SYNC_KEYS)[number];

const SYNC_KEY_SET = new Set<string>(ORIGINAL_FORMATTING_SYNC_KEYS);

/** Keys cleared from direct formatting when a paragraph style is applied. */
export const STYLE_CLEARED_ORIGINAL_KEYS = [
  'alignment',
  'spaceBefore',
  'spaceAfter',
  'lineSpacing',
  'lineSpacingRule',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'hangingIndent',
  'contextualSpacing',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'outlineLevel',
  'runProperties',
] as const;

export function mergeParagraphAttrsWithOriginalFormatting(
  existing: Record<string, unknown>,
  updates: Record<string, unknown>
): Record<string, unknown> {
  const next: Record<string, unknown> = { ...existing, ...updates };
  const orig = existing._originalFormatting as ParagraphFormatting | null | undefined;
  if (!orig) {
    return next;
  }

  const nextOrig: Record<string, unknown> = { ...orig };
  let changed = false;
  for (const [key, value] of Object.entries(updates)) {
    if (!SYNC_KEY_SET.has(key)) continue;
    changed = true;
    if (value == null) {
      delete nextOrig[key];
    } else {
      nextOrig[key] = value;
    }
  }

  if (changed) {
    next._originalFormatting = nextOrig;
  }
  return next;
}

/**
 * After applyStyle: keep provenance bag, set styleId, drop direct pPr/rPr that
 * the style now owns so stale heading spacing etc. cannot resurrect on save.
 */
export function originalFormattingAfterApplyStyle(
  existing: ParagraphFormatting | null | undefined,
  styleId: string
): ParagraphFormatting {
  const next: Record<string, unknown> = existing ? { ...existing } : {};
  next.styleId = styleId;
  for (const key of STYLE_CLEARED_ORIGINAL_KEYS) {
    delete next[key];
  }
  return next as ParagraphFormatting;
}

export function clearOriginalIndentFields(
  existing: Record<string, unknown>
): Record<string, unknown> {
  return mergeParagraphAttrsWithOriginalFormatting(existing, {
    indentLeft: null,
    indentFirstLine: null,
    hangingIndent: null,
  });
}

export type { SyncKey };
