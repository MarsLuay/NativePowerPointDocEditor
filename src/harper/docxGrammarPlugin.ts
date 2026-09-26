import type { HarperGrammarLint } from './harperGrammarService';
import type { Transaction } from 'prosemirror-state';
import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import {
  acceptGrammarResult,
  applyDocxGrammarSuggestion,
  collectChangedTextWindows,
  diagnosticsFromLints,
  grammarLintAllowed,
  type DocxGrammarDiagnostic,
  type DocxGrammarTextWindow,
} from './docxGrammarDiagnostics';

export const docxGrammarPluginKey = new PluginKey<DocxGrammarPluginState>('npde-docx-grammar');

interface DocxGrammarPluginState {
  version: number;
  windows: DocxGrammarTextWindow[];
  diagnostics: DocxGrammarDiagnostic[];
  decorations: DecorationSet;
}

export interface DocxGrammarActions {
  ignore?(diagnostic: DocxGrammarDiagnostic): void;
  addToDictionary?(word: string): void;
}

export interface DocxGrammarPluginOptions {
  getEnabled: () => boolean;
  requestLint: (text: string) => Promise<HarperGrammarLint[] | null>;
  onReview?: (diagnostics: readonly DocxGrammarDiagnostic[]) => void;
  actions?: DocxGrammarActions;
  log?: (data: Record<string, unknown>) => void;
}

function emptyState(): DocxGrammarPluginState {
  return { version: 0, windows: [], diagnostics: [], decorations: DecorationSet.empty };
}

function decorationsFor(doc: EditorState['doc'], diagnostics: readonly DocxGrammarDiagnostic[]): DecorationSet {
  if (diagnostics.length === 0) return DecorationSet.empty;
  return DecorationSet.create(doc, diagnostics.map((diagnostic) => Decoration.inline(diagnostic.from, diagnostic.to, {
    class: 'native-powerpoint-doc-editor-grammar-mark',
  })));
}

function publish(state: DocxGrammarPluginState, onReview: DocxGrammarPluginOptions['onReview']): void {
  onReview?.(state.diagnostics);
}

export function clearDocxGrammarTransaction(state: EditorState) {
  const current = docxGrammarPluginKey.getState(state);
  if (!current || current.diagnostics.length === 0) return null;
  return state.tr.setMeta(docxGrammarPluginKey, { type: 'clear' });
}

export function applyDocxGrammarReviewSuggestion(
  state: EditorState,
  diagnosticId: string,
  suggestionIndex: number,
) {
  const current = docxGrammarPluginKey.getState(state);
  const diagnostic = current?.diagnostics.find((item) => item.id === diagnosticId);
  const suggestion = diagnostic?.suggestions[suggestionIndex];
  if (!diagnostic || !suggestion) return null;
  return applyDocxGrammarSuggestion(state, diagnostic.from, diagnostic.to, suggestion)
    .setMeta(docxGrammarPluginKey, { type: 'applied', id: diagnosticId });
}

export function createDocxGrammarPlugin(options: DocxGrammarPluginOptions): Plugin<DocxGrammarPluginState> {
  let requestVersion = 0;

  const lintWindows = (view: EditorView, windows: DocxGrammarTextWindow[], version: number) => {
    const captured = requestVersion;
    void (async () => {
      const next: DocxGrammarDiagnostic[] = [];
      for (const window of windows) {
        if (captured !== requestVersion || !grammarLintAllowed({ enabled: options.getEnabled(), composing: view.composing })) {
          return;
        }
        const lints = await options.requestLint(window.text);
        if (lints === null || captured !== requestVersion || !acceptGrammarResult(version, docxGrammarPluginKey.getState(view.state)?.version ?? -1)) {
          options.log?.({ phase: 'stale', version, lintCount: lints?.length ?? 0 });
          return;
        }
        next.push(...diagnosticsFromLints(view.state.doc, window, lints, version));
      }
      if (captured !== requestVersion) return;
      options.log?.({ phase: 'diagnostics', version, diagnosticCount: next.length, windowCount: windows.length });
      view.dispatch(view.state.tr.setMeta(docxGrammarPluginKey, { type: 'results', version, diagnostics: next }));
    })();
  };

  return new Plugin<DocxGrammarPluginState>({
    key: docxGrammarPluginKey,
    state: {
      init: () => emptyState(),
      apply(tr, value, _oldState, newState) {
        const meta = tr.getMeta(docxGrammarPluginKey) as
          | { type: 'dirty'; version: number; windows: DocxGrammarTextWindow[] }
          | { type: 'results'; version: number; diagnostics: DocxGrammarDiagnostic[] }
          | { type: 'clear' }
          | { type: 'applied'; id: string }
          | undefined;
        if (meta?.type === 'clear') {
          const cleared = emptyState();
          cleared.version = value.version + 1;
          publish(cleared, options.onReview);
          return cleared;
        }
        if (meta?.type === 'dirty') {
          return { ...value, version: meta.version, windows: meta.windows };
        }
        if (meta?.type === 'results') {
          if (!acceptGrammarResult(meta.version, value.version)) return value;
          const next = {
            ...value,
            diagnostics: meta.diagnostics,
            decorations: decorationsFor(newState.doc, meta.diagnostics),
          };
          publish(next, options.onReview);
          return next;
        }
        if (meta?.type === 'applied') {
          const diagnostics = value.diagnostics.filter((item) => item.id !== meta.id);
          const next = { ...value, diagnostics, decorations: decorationsFor(newState.doc, diagnostics) };
          publish(next, options.onReview);
          return next;
        }
        if (tr.docChanged) {
          return {
            ...value,
            decorations: value.decorations.map(tr.mapping, tr.doc),
          };
        }
        return value;
      },
    },
    props: {
      decorations(state) {
        return docxGrammarPluginKey.getState(state)?.decorations ?? DecorationSet.empty;
      },
    },
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) return null;
      if (transactions.some((transaction) => transaction.getMeta(docxGrammarPluginKey))) return null;
      if (!grammarLintAllowed({ enabled: options.getEnabled(), composing: false })) {
        const current = docxGrammarPluginKey.getState(newState);
        if (!current || current.diagnostics.length === 0) return null;
        return newState.tr.setMeta(docxGrammarPluginKey, { type: 'clear' });
      }
      const windows = collectChangedTextWindows(oldState.doc, newState.doc);
      const version = (docxGrammarPluginKey.getState(newState)?.version ?? 0) + 1;
      return newState.tr.setMeta(docxGrammarPluginKey, { type: 'dirty', version, windows });
    },
    view(editorView) {
      let wasComposing = editorView.composing;
      const run = () => {
        if (!grammarLintAllowed({ enabled: options.getEnabled(), composing: editorView.composing })) return;
        const current = docxGrammarPluginKey.getState(editorView.state);
        if (!current || current.windows.length === 0) return;
        requestVersion = current.version;
        lintWindows(editorView, current.windows, current.version);
      };
      return {
        update(view, previous) {
          const finishedComposition = wasComposing && !view.composing;
          wasComposing = view.composing;
          if (!options.getEnabled()) {
            if ((docxGrammarPluginKey.getState(view.state)?.diagnostics.length ?? 0) > 0) {
              const clear = clearDocxGrammarTransaction(view.state);
              if (clear) view.dispatch(clear);
            }
            options.onReview?.([]);
            return;
          }
          if (view.composing) return;
          const current = docxGrammarPluginKey.getState(view.state);
          const prior = docxGrammarPluginKey.getState(previous);
          if (finishedComposition || (current && prior && current.version !== prior.version && current.windows.length > 0)) {
            run();
          }
        },
        destroy() {
          requestVersion += 1;
          options.onReview?.([]);
        },
      };
    },
  });
}

export function ignoreDocxGrammarDiagnostic(
  state: EditorState,
  diagnosticId: string,
  actions: DocxGrammarActions | undefined,
): Transaction | null {
  const current = docxGrammarPluginKey.getState(state);
  const diagnostic = current?.diagnostics.find((item) => item.id === diagnosticId);
  if (!diagnostic) return null;
  actions?.ignore?.(diagnostic);
  const diagnostics = current?.diagnostics.filter((item) => item.id !== diagnosticId) ?? [];
  return state.tr.setMeta(docxGrammarPluginKey, { type: 'results', version: current?.version ?? 0, diagnostics });
}

export function addDocxGrammarDictionaryWord(
  state: EditorState,
  diagnosticId: string,
  actions: DocxGrammarActions | undefined,
): string | null {
  const current = docxGrammarPluginKey.getState(state);
  const diagnostic = current?.diagnostics.find((item) => item.id === diagnosticId);
  if (!diagnostic) return null;
  const word = state.doc.textBetween(diagnostic.from, diagnostic.to, '', '');
  if (word) actions?.addToDictionary?.(word);
  return word || null;
}
