import type { Node } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { HarperGrammarLint, HarperGrammarSuggestion } from './harperGrammarService';

export const DOCX_GRAMMAR_TEXT_WINDOW = 2_000;

export interface DocxGrammarTextWindow {
  blockPos: number;
  textOffset: number;
  text: string;
}

export interface DocxGrammarDiagnostic {
  id: string;
  from: number;
  to: number;
  message: string;
  suggestions: HarperGrammarSuggestion[];
  version: number;
}

interface TextblockSlice {
  pos: number;
  text: string;
}

function textblocks(doc: Node): TextblockSlice[] {
  const blocks: TextblockSlice[] = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    blocks.push({ pos, text: node.textContent });
    return false;
  });
  return blocks;
}

function boundedWindow(before: string, after: string, blockPos: number): DocxGrammarTextWindow {
  if (after.length <= DOCX_GRAMMAR_TEXT_WINDOW) {
    return { blockPos, textOffset: 0, text: after };
  }
  let diff = 0;
  const limit = Math.min(before.length, after.length);
  while (diff < limit && before[diff] === after[diff]) diff += 1;
  const start = Math.max(0, diff - Math.floor(DOCX_GRAMMAR_TEXT_WINDOW / 4));
  return {
    blockPos,
    textOffset: start,
    text: after.slice(start, start + DOCX_GRAMMAR_TEXT_WINDOW),
  };
}

/** Paragraphs whose text changed. Unchanged text is skipped even if positions moved. */
export function collectChangedTextWindows(before: Node, after: Node): DocxGrammarTextWindow[] {
  const previous = textblocks(before);
  const next = textblocks(after);
  const windows: DocxGrammarTextWindow[] = [];
  let previousIndex = 0;
  let nextIndex = 0;
  while (previousIndex < previous.length && nextIndex < next.length) {
    const oldBlock = previous[previousIndex];
    const newBlock = next[nextIndex];
    if (!oldBlock || !newBlock) break;
    if (oldBlock.text === newBlock.text) {
      previousIndex += 1;
      nextIndex += 1;
      continue;
    }
    windows.push(boundedWindow(oldBlock.text, newBlock.text, newBlock.pos));
    if (next[nextIndex + 1]?.text === oldBlock.text) {
      nextIndex += 1;
      continue;
    }
    previousIndex += 1;
    nextIndex += 1;
  }
  while (nextIndex < next.length) {
    const added = next[nextIndex];
    if (added) windows.push(boundedWindow('', added.text, added.pos));
    nextIndex += 1;
  }
  return windows;
}

export function positionAtTextOffset(block: Node, blockPos: number, offset: number): number | null {
  if (offset < 0 || offset > block.textContent.length) return null;
  let seen = 0;
  let found: number | null = null;
  block.descendants((node, pos) => {
    if (!node.isText || found !== null) return false;
    const text = node.text ?? '';
    const next = seen + text.length;
    if (offset <= next) {
      found = blockPos + 1 + pos + (offset - seen);
      return false;
    }
    seen = next;
    return false;
  });
  if (found === null && offset === 0) return blockPos + 1;
  return found;
}

export function mapWindowSpan(
  doc: Node,
  window: DocxGrammarTextWindow,
  span: { start: number; end: number },
): { from: number; to: number } | null {
  const block = doc.nodeAt(window.blockPos);
  if (!block?.isTextblock) return null;
  const start = window.textOffset + span.start;
  const end = window.textOffset + span.end;
  const from = positionAtTextOffset(block, window.blockPos, start);
  const to = positionAtTextOffset(block, window.blockPos, end);
  if (from === null || to === null || to < from) return null;
  return { from, to };
}

export function acceptGrammarResult(resultVersion: number, currentVersion: number): boolean {
  return resultVersion === currentVersion;
}

export function grammarLintAllowed(input: { enabled: boolean; composing: boolean }): boolean {
  return input.enabled && !input.composing;
}

export function applyDocxGrammarSuggestion(
  state: EditorState,
  from: number,
  to: number,
  suggestion: HarperGrammarSuggestion,
): Transaction {
  let marks = state.doc.resolve(from).marks();
  state.doc.nodesBetween(from, Math.max(from, to), (node) => {
    if (node.isText && node.marks.length > 0) {
      marks = node.marks;
      return false;
    }
    return true;
  });
  if (suggestion.kind === 'remove') {
    return state.tr.delete(from, to);
  }
  if (suggestion.kind === 'insert-after') {
    return state.tr.insert(to, state.schema.text(suggestion.replacement, marks));
  }
  const text = suggestion.replacement ? state.schema.text(suggestion.replacement, marks) : null;
  return state.tr.replaceWith(from, to, text ? [text] : []);
}

export function diagnosticsFromLints(
  doc: Node,
  window: DocxGrammarTextWindow,
  lints: readonly HarperGrammarLint[],
  version: number,
): DocxGrammarDiagnostic[] {
  const diagnostics: DocxGrammarDiagnostic[] = [];
  lints.forEach((lint, index) => {
    const mapped = mapWindowSpan(doc, window, lint.span);
    if (!mapped) return;
    diagnostics.push({
      id: `${version}:${window.blockPos}:${index}`,
      from: mapped.from,
      to: mapped.to,
      message: lint.message,
      suggestions: lint.suggestions,
      version,
    });
  });
  return diagnostics;
}
