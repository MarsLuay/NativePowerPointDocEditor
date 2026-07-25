/**
 * Shared mark utility functions
 *
 * setMark, removeMark, isMarkActive, getMarkAttr, marksToTextFormatting, textFormattingToMarks, clearFormatting
 */

import { toggleMark } from 'prosemirror-commands';
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import { TextSelection } from 'prosemirror-state';
import type { MarkType, Mark, Schema } from 'prosemirror-model';
import type { TextFormatting } from '../../../types/document';

type MarkAttrs = Record<string, unknown>;

// ============================================================================
// PARAGRAPH DEFAULT FORMATTING HELPERS
// ============================================================================

export function marksToTextFormatting(marks: readonly Mark[]): TextFormatting {
  const formatting: TextFormatting = {};

  for (const mark of marks) {
    switch (mark.type.name) {
      case 'bold':
        formatting.bold = true;
        break;
      case 'italic':
        formatting.italic = true;
        break;
      case 'underline':
        formatting.underline = { style: mark.attrs.style || 'single' };
        break;
      case 'strike':
        formatting.strike = true;
        break;
      case 'textColor':
        formatting.color = mark.attrs;
        break;
      case 'highlight':
        formatting.highlight = mark.attrs.color;
        break;
      case 'fontSize':
        // CS-only RTL runs carry the size in `sizeCs`; fall back so the toolbar
        // field isn't blank for them.
        formatting.fontSize = mark.attrs.size ?? mark.attrs.sizeCs;
        // Preserve a genuinely distinct complex-script size so a run with
        // different Latin/CS sizes survives a read -> textFormattingToMarks
        // round-trip (e.g. stored-mark persistence); without it fontSizeCs
        // stays undefined and the next write re-aligns sizeCs to fontSize.
        // Only set when sizeCs is present so Latin-only runs stay fontSize-only.
        if (mark.attrs.sizeCs != null) formatting.fontSizeCs = mark.attrs.sizeCs;
        break;
      case 'fontFamily':
        formatting.fontFamily = {
          ascii: mark.attrs.ascii,
          hAnsi: mark.attrs.hAnsi,
        };
        break;
      case 'superscript':
        formatting.vertAlign = 'superscript';
        break;
      case 'subscript':
        formatting.vertAlign = 'subscript';
        break;
      case 'rtl':
        // Per-run right-to-left flag (`<w:rtl/>`). Without this case, formatting
        // helpers that route through markUtils (live-edit commands, clipboard)
        // silently drop run direction for Arabic/Hebrew/etc. text. Fixes #806.
        formatting.rtl = true;
        break;
    }
  }

  return formatting;
}

/**
 * Mirror the cursor's stored marks into the paragraph's `defaultTextFormatting`
 * attr so an empty paragraph renders with the right caret height/font.
 *
 * IMPORTANT: callers must invoke this BEFORE `tr.setStoredMarks(...)`. The
 * `setNodeMarkup` step appended here clears `tr.storedMarks` (every step does —
 * see prosemirror-state Transaction.addStep), so stored marks must be set last.
 * Marks are passed in explicitly rather than read off `tr.storedMarks` for the
 * same reason.
 */
function saveStoredMarksToParagraph(
  state: EditorState,
  tr: Transaction,
  marks: readonly Mark[]
): Transaction {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return tr;
  if (paragraph.textContent.length > 0) return tr;

  const originalFormatting = paragraph.attrs._originalFormatting as
    | { runProperties?: TextFormatting }
    | null
    | undefined;

  if (marks.length === 0) {
    const nextAttrs: Record<string, unknown> = {
      ...paragraph.attrs,
      defaultTextFormatting: null,
    };
    // Clear direct pPr/rPr on save so emptied fonts do not resurrect from
    // _originalFormatting after tab switch / reload.
    if (originalFormatting) {
      const nextOriginal = { ...originalFormatting };
      delete nextOriginal.runProperties;
      nextAttrs._originalFormatting = nextOriginal;
    }
    return tr.setNodeMarkup($from.before(), undefined, nextAttrs);
  }

  const defaultTextFormatting = marksToTextFormatting(marks);
  const nextAttrs: Record<string, unknown> = {
    ...paragraph.attrs,
    defaultTextFormatting,
  };
  // fromProseDoc serializes via _originalFormatting when present and previously
  // never copied defaultTextFormatting → runProperties. Sync both so empty-line
  // font/size survive autosave + tab remount.
  if (originalFormatting) {
    nextAttrs._originalFormatting = {
      ...originalFormatting,
      runProperties: defaultTextFormatting,
    };
  }

  return tr.setNodeMarkup($from.before(), undefined, nextAttrs);
}

// ============================================================================
// CORE MARK COMMANDS
// ============================================================================

/**
 * Apply a new stored-mark set at a collapsed cursor and mirror it into the
 * paragraph's defaultTextFormatting. Order matters: setNodeMarkup runs first
 * because every transform step clears tr.storedMarks, so setStoredMarks must
 * be the last mutation.
 */
function dispatchStoredMarks(
  state: EditorState,
  dispatch: (tr: Transaction) => void,
  marks: readonly Mark[]
): void {
  let tr = state.tr;
  tr = saveStoredMarksToParagraph(state, tr, marks);
  tr.setStoredMarks(marks);
  dispatch(tr);
}

/**
 * A range mark cannot affect an empty paragraph because it has no inline
 * content. Mirror the requested change into every selected empty paragraph's
 * defaults so Ctrl+A formatting and later cursor selection agree.
 */
function updateSelectedEmptyParagraphDefaults(
  state: EditorState,
  tr: Transaction,
  update: (formatting: TextFormatting) => TextFormatting
): Transaction {
  const { from, to, empty } = state.selection;
  if (empty) return tr;

  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name !== 'paragraph' || node.content.size !== 0) return;

    const current = (node.attrs.defaultTextFormatting as TextFormatting | null | undefined) ?? {};
    const next = update(current);
    const hasFormatting = Object.keys(next).length > 0;
    const nextAttrs: Record<string, unknown> = {
      ...node.attrs,
      defaultTextFormatting: hasFormatting ? next : null,
    };
    const originalFormatting = node.attrs._originalFormatting as
      | { runProperties?: TextFormatting }
      | null
      | undefined;
    if (originalFormatting) {
      const nextOriginal = { ...originalFormatting };
      const nextRunProperties = update(originalFormatting.runProperties ?? {});
      if (Object.keys(nextRunProperties).length > 0) {
        nextOriginal.runProperties = nextRunProperties;
      } else {
        delete nextOriginal.runProperties;
      }
      nextAttrs._originalFormatting = nextOriginal;
    }

    tr = tr.setNodeMarkup(pos, undefined, nextAttrs);
  });

  return tr;
}

function withoutFormattingKeys(
  formatting: TextFormatting,
  keys: readonly string[]
): TextFormatting {
  const next = { ...formatting } as Record<string, unknown>;
  for (const key of keys) {
    delete next[key];
  }
  return next as TextFormatting;
}

export function setMark(markType: MarkType, attrs: MarkAttrs): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    const mark = markType.create(attrs);

    if (empty) {
      if (dispatch) {
        const current = state.storedMarks || state.selection.$from.marks();
        const sansType = markType.isInSet(current)
          ? current.filter((m) => m.type !== markType)
          : current;
        dispatchStoredMarks(state, dispatch, [...sansType, mark]);
      }
      return true;
    }

    if (dispatch) {
      const patch = marksToTextFormatting([mark]);
      const tr = updateSelectedEmptyParagraphDefaults(
        state,
        state.tr.addMark(from, to, mark),
        (formatting) => ({ ...formatting, ...patch })
      );
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export function removeMark(markType: MarkType): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;

    if (empty) {
      if (dispatch) {
        const next = (state.storedMarks || state.selection.$from.marks()).filter(
          (m) => m.type !== markType
        );
        dispatchStoredMarks(state, dispatch, next);
      }
      return true;
    }

    if (dispatch) {
      const keys = Object.keys(marksToTextFormatting([markType.create()]));
      const tr = updateSelectedEmptyParagraphDefaults(
        state,
        state.tr.removeMark(from, to, markType),
        (formatting) => withoutFormattingKeys(formatting, keys)
      );
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/**
 * Toggle a boolean-ish mark (bold/italic/underline/strike/super/subscript).
 *
 * On a non-empty selection this is byte-for-byte stock `toggleMark` behavior
 * (delegated, not reimplemented, so range semantics never drift).
 *
 * On a collapsed cursor, stock `toggleMark` only flips `state.storedMarks`
 * via `addStoredMark`/`removeStoredMark` — it never touches the paragraph's
 * `defaultTextFormatting`/`_originalFormatting`, so the toggle is lost on
 * blur/save for an empty paragraph. Route that case through `setMark`/
 * `removeMark` instead, which call `dispatchStoredMarks` ->
 * `saveStoredMarksToParagraph` to keep those attrs in sync (same path
 * font/size/color already use). The `toggleMark(...)(state, undefined)`
 * call below reuses its `markApplies` applicability guard without
 * duplicating it.
 */
export function toggleMarkWithParagraphDefaults(
  markType: MarkType,
  attrs: MarkAttrs | null = null
): Command {
  const rangeToggle = toggleMark(markType, attrs);

  return (state, dispatch) => {
    const selection = state.selection;
    const $cursor = selection instanceof TextSelection ? selection.$cursor : null;

    if (!(selection.empty && $cursor)) {
      return rangeToggle(state, dispatch);
    }

    if (!rangeToggle(state, undefined)) return false;

    if (dispatch) {
      const active = markType.isInSet(state.storedMarks || $cursor.marks());
      if (active) {
        removeMark(markType)(state, dispatch);
      } else {
        setMark(markType, attrs || {})(state, dispatch);
      }
    }
    return true;
  };
}

/**
 * Check if a mark is active in the current selection
 */
export function isMarkActive(
  state: EditorState,
  markType: MarkType,
  attrs?: Record<string, unknown>
): boolean {
  const { from, to, empty } = state.selection;

  if (empty) {
    const marks = state.storedMarks || state.selection.$from.marks();
    return marks.some((mark) => {
      if (mark.type !== markType) return false;
      if (!attrs) return true;
      return Object.entries(attrs).every(([key, value]) => mark.attrs[key] === value);
    });
  }

  let hasMark = false;
  state.doc.nodesBetween(from, to, (node) => {
    if (node.isText) {
      const mark = markType.isInSet(node.marks);
      if (mark) {
        if (!attrs) {
          hasMark = true;
          return false;
        }
        const attrsMatch = Object.entries(attrs).every(([key, value]) => mark.attrs[key] === value);
        if (attrsMatch) {
          hasMark = true;
          return false;
        }
      }
    }
    return true;
  });

  return hasMark;
}

/**
 * Get the current value of a mark attribute
 */
export function getMarkAttr(state: EditorState, markType: MarkType, attr: string): unknown {
  const { empty, $from, from, to } = state.selection;

  if (empty) {
    const marks = state.storedMarks || $from.marks();
    for (const mark of marks) {
      if (mark.type === markType) {
        return mark.attrs[attr];
      }
    }
    return null;
  }

  let value: unknown = null;
  state.doc.nodesBetween(from, to, (node) => {
    if (node.isText && value === null) {
      const mark = markType.isInSet(node.marks);
      if (mark) {
        value = mark.attrs[attr];
        return false;
      }
    }
    return true;
  });

  return value;
}

/**
 * Convert TextFormatting to marks array (used to restore formatting on empty paragraphs)
 */
export function textFormattingToMarks(formatting: TextFormatting, schema: Schema): Mark[] {
  const marks: Mark[] = [];

  if (formatting.bold) {
    marks.push(schema.marks.bold.create());
  }
  if (formatting.italic) {
    marks.push(schema.marks.italic.create());
  }
  if (formatting.underline) {
    marks.push(
      schema.marks.underline.create({
        style: formatting.underline.style || 'single',
        color: formatting.underline.color,
      })
    );
  }
  if (formatting.strike) {
    marks.push(schema.marks.strike.create());
  }
  if (formatting.doubleStrike) {
    marks.push(schema.marks.strike.create({ double: true }));
  }
  if (formatting.color) {
    marks.push(
      schema.marks.textColor.create({
        rgb: formatting.color.rgb,
        themeColor: formatting.color.themeColor,
        themeTint: formatting.color.themeTint,
        themeShade: formatting.color.themeShade,
      })
    );
  }
  if (formatting.highlight) {
    marks.push(schema.marks.highlight.create({ color: formatting.highlight }));
  }
  if (formatting.fontSize) {
    marks.push(
      schema.marks.fontSize.create({
        size: formatting.fontSize,
        sizeCs: formatting.fontSizeCs ?? formatting.fontSize,
      })
    );
  }
  if (formatting.fontFamily) {
    marks.push(
      schema.marks.fontFamily.create({
        ascii: formatting.fontFamily.ascii,
        hAnsi: formatting.fontFamily.hAnsi,
        asciiTheme: formatting.fontFamily.asciiTheme,
      })
    );
  }
  if (formatting.vertAlign === 'superscript') {
    marks.push(schema.marks.superscript.create());
  }
  if (formatting.vertAlign === 'subscript') {
    marks.push(schema.marks.subscript.create());
  }
  if (formatting.rtl) {
    marks.push(schema.marks.rtl.create());
  }

  return marks;
}

/**
 * Clear all text formatting (remove all marks)
 */
export const clearFormatting: Command = (state, dispatch) => {
  const { from, to, empty } = state.selection;

  if (empty) {
    if (dispatch) {
      // Clear the paragraph's run defaults too, so EmptyParagraphFormatExtension
      // doesn't re-derive stored marks from them right after the clear.
      const tr = saveStoredMarksToParagraph(state, state.tr, []);
      tr.setStoredMarks([]);
      dispatch(tr);
    }
    return true;
  }

  if (dispatch) {
    let tr = state.tr;

    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.isText && node.marks.length > 0) {
        const start = Math.max(from, pos);
        const end = Math.min(to, pos + node.nodeSize);
        for (const mark of node.marks) {
          tr = tr.removeMark(start, end, mark.type);
        }
      }
    });

    dispatch(tr.scrollIntoView());
  }

  return true;
};

/**
 * Create a command that sets a mark on the selection
 */
export function createSetMarkCommand(markType: MarkType, attrs?: Record<string, unknown>): Command {
  return setMark(markType, attrs || {});
}

/**
 * Create a command that removes a mark from the selection
 */
export function createRemoveMarkCommand(markType: MarkType): Command {
  return removeMark(markType);
}
