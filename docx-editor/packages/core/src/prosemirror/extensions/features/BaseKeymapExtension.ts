/**
 * Base Keymap Extension — wraps prosemirror-commands baseKeymap
 *
 * Priority: Low (150) — must be the last keymap so other extensions can override keys
 */

import {
  baseKeymap,
  splitBlock,
  createParagraphNear,
  liftEmptyBlock,
  deleteSelection,
  joinBackward,
  joinForward,
  selectAll,
  selectParentNode,
  chainCommands,
} from 'prosemirror-commands';
import type { Mark, Node as PMNode, Schema } from 'prosemirror-model';
import { createExtension } from '../create';
import { textFormattingToMarks, marksToTextFormatting } from '../marks/markUtils';
import { Priority } from '../types';
import type { ExtensionRuntime, ExtensionContext } from '../types';
import type { Command, Transaction } from 'prosemirror-state';
import type { TextFormatting } from '../../../types/document';
import type { StyleResolver } from '../../styles/styleResolver';
import { paragraphAttrsFromResolvedStyle } from '../../styles/resolvedStyleAttrs';
import { getDocumentStyleResolver } from '../../plugins/documentStyles';
import { mergeParagraphAttrsWithOriginalFormatting } from '../../syncOriginalFormatting';

/**
 * Backspace at the start of a paragraph clears first-line indent / hanging indent
 * before joining with the previous paragraph (matches Word behavior).
 */
const clearIndentOnBackspace: Command = (state, dispatch) => {
  const { $cursor } = state.selection as {
    $cursor?: {
      parentOffset: number;
      parent: { type: { name: string }; attrs: Record<string, unknown> };
      pos: number;
      before: () => number;
    };
  };
  if (!$cursor) return false;

  // Only at the very start of a paragraph
  if ($cursor.parentOffset !== 0) return false;
  if ($cursor.parent.type.name !== 'paragraph') return false;

  const attrs = $cursor.parent.attrs;
  const hasFirstLine = attrs.indentFirstLine != null && (attrs.indentFirstLine as number) > 0;
  const hasHanging = !!attrs.hangingIndent;
  const hasIndentLeft = attrs.indentLeft != null && (attrs.indentLeft as number) > 0;

  if (!hasFirstLine && !hasHanging && !hasIndentLeft) return false;

  if (dispatch) {
    const pos = $cursor.before();
    const tr = state.tr.setNodeMarkup(
      pos,
      undefined,
      mergeParagraphAttrsWithOriginalFormatting(attrs, {
        indentFirstLine: null,
        hangingIndent: null,
        indentLeft: null,
      })
    );
    dispatch(tr.scrollIntoView());
  }
  return true;
};

/**
 * Custom Enter handler: splits the block, inherits style-related attrs,
 * clears paragraph borders, and preserves font marks on the new paragraph.
 *
 * splitBlock creates a new paragraph with default attrs (all null),
 * so we must manually copy style-related attrs from the source paragraph.
 * Word does NOT propagate paragraph borders (w:pBdr) on Enter.
 */
const INHERITED_PARA_ATTRS = [
  'defaultTextFormatting',
  'styleId',
  'lineSpacing',
  'lineSpacingRule',
  'spaceAfter',
  'spaceBefore',
  'contextualSpacing',
] as const;

/** Mark types inherited across Enter onto an empty new paragraph. */
export const STYLE_MARK_NAMES = new Set([
  'fontFamily',
  'fontSize',
  'textColor',
  'bold',
  'italic',
  'underline',
  'strike',
  'superscript',
  'subscript',
]);

/**
 * Replace an empty new paragraph's style with `nextStyleId`, projecting that
 * style's resolved paragraph + run formatting onto the node. Stored marks are
 * set from the next style's run formatting so typed text matches (and so the
 * previous paragraph's marks — e.g. a heading's bold — don't carry over).
 */
function applyNextParagraphStyle(
  tr: Transaction,
  pos: number,
  newPara: PMNode,
  nextStyleId: string,
  resolver: StyleResolver,
  schema: Schema
): void {
  const resolved = resolver.resolveParagraphStyle(nextStyleId);
  tr.setNodeMarkup(pos, undefined, {
    ...newPara.attrs,
    styleId: nextStyleId,
    ...paragraphAttrsFromResolvedStyle(resolved),
    borders: null,
  });
  // setStoredMarks MUST come after setNodeMarkup — every step clears them.
  tr.setStoredMarks(
    resolved.runFormatting ? textFormattingToMarks(resolved.runFormatting, schema) : []
  );
}

/**
 * Apply post-split paragraph inheritance to `tr`. Assumes the split already
 * happened and the caller's intent is that `tr.selection.$from` resolves
 * into the NEW (second) paragraph. When the source style has a `w:next`,
 * switches the empty new paragraph to it (e.g. heading → body text);
 * otherwise copies style-related attrs from `sourcePara`, clears borders,
 * and — for an empty new paragraph — syncs defaultTextFormatting and
 * `setStoredMarks` so typed text inherits font / size / color from the source.
 *
 * Shared by the plain Enter handler and the suggesting-mode Enter handler.
 */
export function applyPostSplitInheritance(
  tr: Transaction,
  sourcePara: PMNode | null,
  styleMarks: readonly Mark[],
  schema: Schema,
  resolver?: StyleResolver | null
): void {
  const { $from } = tr.selection;
  const newPara = $from.parent;
  if (newPara.type.name !== 'paragraph') return;

  // Word's `w:next`: pressing Enter at the end of a paragraph (the new
  // paragraph is empty) switches it to the style's follow-on style — e.g.
  // a heading drops to body text. Only applies to an empty trailing
  // paragraph; splitting mid-paragraph keeps the style on both halves.
  if (resolver && sourcePara && newPara.textContent.length === 0) {
    const nextStyleId = resolver.getNextStyleId(sourcePara.attrs.styleId as string | null);
    if (nextStyleId) {
      applyNextParagraphStyle(tr, $from.before(), newPara, nextStyleId, resolver, schema);
      return;
    }
  }

  const newAttrs: Record<string, unknown> = { ...newPara.attrs };
  let attrsChanged = false;

  if (sourcePara) {
    for (const key of INHERITED_PARA_ATTRS) {
      const srcVal = (sourcePara.attrs as Record<string, unknown>)[key];
      if (srcVal != null && newAttrs[key] == null) {
        newAttrs[key] = srcVal;
        attrsChanged = true;
      }
    }
  }

  if (newAttrs.borders) {
    newAttrs.borders = null;
    attrsChanged = true;
  }

  if (attrsChanged) {
    tr.setNodeMarkup($from.before(), undefined, newAttrs);
  }

  if (newPara.textContent.length !== 0) return;

  let effectiveMarks: Mark[] = [...styleMarks];
  if (effectiveMarks.length === 0 && sourcePara) {
    const dtf = sourcePara.attrs.defaultTextFormatting as TextFormatting | undefined;
    if (dtf) {
      const allMarks = textFormattingToMarks(dtf, schema);
      effectiveMarks = allMarks.filter((m) => STYLE_MARK_NAMES.has(m.type.name));
    }
  }
  if (effectiveMarks.length === 0) return;

  const fromMarks = marksToTextFormatting(effectiveMarks);
  const dtf: TextFormatting = {
    ...((newAttrs.defaultTextFormatting as TextFormatting) ?? {}),
    ...fromMarks,
  };
  const orig = newAttrs._originalFormatting as TextFormatting | Record<string, unknown> | null | undefined;
  const nextAttrs: Record<string, unknown> = {
    ...newAttrs,
    defaultTextFormatting: dtf,
  };
  if (orig && typeof orig === 'object') {
    nextAttrs._originalFormatting = {
      ...orig,
      runProperties: dtf,
    };
  }
  tr.setNodeMarkup($from.before(), undefined, nextAttrs);

  // setStoredMarks MUST come after setNodeMarkup — ReplaceStep clears stored marks.
  tr.setStoredMarks(effectiveMarks);
}

const splitBlockClearBorders: Command = (state, dispatch, view) => {
  // Capture source paragraph info BEFORE split (splitBlock resets everything)
  const { $from: preSplitFrom } = state.selection;
  const sourcePara = preSplitFrom.parent.type.name === 'paragraph' ? preSplitFrom.parent : null;

  // Collect style marks from the cursor position before splitting.
  // Use storedMarks if set, otherwise resolve from the position.
  const preMarks = state.storedMarks || preSplitFrom.marks();
  const styleMarks = preMarks.filter((m) => STYLE_MARK_NAMES.has(m.type.name));

  // Intercept splitBlock's transaction so we can modify it before dispatch.
  // This ensures attrs + stored marks are set in a single transaction,
  // avoiding a flash where the empty paragraph has no formatting.
  let splitTr: Transaction | null = null;
  const capturingDispatch = dispatch
    ? (tr: Transaction) => {
        splitTr = tr;
      }
    : undefined;

  if (!splitBlock(state, capturingDispatch, view)) {
    return false;
  }

  if (dispatch && splitTr !== null) {
    const tr = splitTr as Transaction;
    applyPostSplitInheritance(
      tr,
      sourcePara,
      styleMarks,
      state.schema,
      getDocumentStyleResolver(state)
    );
    dispatch(tr.scrollIntoView());
  }

  return true;
};

export const BaseKeymapExtension = createExtension({
  name: 'baseKeymap',
  priority: Priority.Low,
  onSchemaReady(_ctx: ExtensionContext): ExtensionRuntime {
    return {
      keyboardShortcuts: {
        // Base keymap provides default editing commands
        ...baseKeymap,
        // Override some keys with better defaults
        // Prefer border-clearing split; fall back to stock Enter chain pieces so
        // NodeSelection / edge gaps are not a total Enter no-op.
        Enter: chainCommands(splitBlockClearBorders, createParagraphNear, liftEmptyBlock),
        Backspace: chainCommands(deleteSelection, clearIndentOnBackspace, joinBackward),
        Delete: chainCommands(deleteSelection, joinForward),
        'Mod-a': selectAll,
        Escape: selectParentNode,
      },
      commands: {
        // Host bridges (PagedEditor) call this when focus-steal + preventDefault
        // leaves keymap props returning falsy.
        splitParagraph: () =>
          chainCommands(splitBlockClearBorders, createParagraphNear, liftEmptyBlock),
      },
    };
  },
});

/** Enter command chain used by BaseKeymap and host focus bridges. */
export const enterCommand: Command = chainCommands(
  splitBlockClearBorders,
  createParagraphNear,
  liftEmptyBlock,
);
