import type { HarperGrammarLint, HarperGrammarSuggestion } from './harperGrammarService';

export interface PptxGrammarDiagnostic {
  id: string;
  start: number;
  end: number;
  message: string;
  suggestions: HarperGrammarSuggestion[];
  version: number;
}

export interface PptxGrammarScheduleHandle {
  cancel(): void;
}

export function replaceTextareaSpan(
  text: string,
  start: number,
  end: number,
  suggestion: HarperGrammarSuggestion,
): string {
  const boundedStart = Math.max(0, Math.min(start, text.length));
  const boundedEnd = Math.max(boundedStart, Math.min(end, text.length));
  if (suggestion.kind === 'remove') {
    return `${text.slice(0, boundedStart)}${text.slice(boundedEnd)}`;
  }
  if (suggestion.kind === 'insert-after') {
    return `${text.slice(0, boundedEnd)}${suggestion.replacement}${text.slice(boundedEnd)}`;
  }
  return `${text.slice(0, boundedStart)}${suggestion.replacement}${text.slice(boundedEnd)}`;
}

export function pptxGrammarLintAllowed(input: { enabled: boolean; composing: boolean }): boolean {
  return input.enabled && !input.composing;
}

export function acceptPptxGrammarResult(resultVersion: number, currentVersion: number): boolean {
  return resultVersion === currentVersion;
}

export function applyPptxGrammarEdit(input: {
  text: string;
  start: number;
  end: number;
  suggestion: HarperGrammarSuggestion;
  history: string[];
  save: (text: string) => void;
}): { previous: string; saved: string } {
  input.history.push(input.text);
  const saved = replaceTextareaSpan(input.text, input.start, input.end, input.suggestion);
  input.save(saved);
  return { previous: input.text, saved };
}

export interface PptxTextGrammarSession {
  noteText(text: string): void;
  setComposing(composing: boolean): void;
  setEnabled(enabled: boolean): void;
  apply(id: string, suggestionIndex: number): string | null;
  clear(): void;
  diagnostics(): readonly PptxGrammarDiagnostic[];
}

export function createPptxTextGrammarSession(options: {
  getEnabled: () => boolean;
  requestLint: (text: string) => Promise<HarperGrammarLint[] | null>;
  schedule?: (callback: () => void, delayMs: number) => PptxGrammarScheduleHandle;
  debounceMs?: number;
  onReview?: (diagnostics: readonly PptxGrammarDiagnostic[]) => void;
  log?: (data: Record<string, unknown>) => void;
}): PptxTextGrammarSession {
  const debounceMs = options.debounceMs ?? 250;
  const schedule = options.schedule ?? ((callback, delayMs) => {
    const timer = window.setTimeout(callback, delayMs);
    return { cancel: () => window.clearTimeout(timer) };
  });
  let version = 0;
  let composing = false;
  let text = '';
  let timer: PptxGrammarScheduleHandle | null = null;
  let diagnostics: PptxGrammarDiagnostic[] = [];

  const publish = () => options.onReview?.(diagnostics);

  return {
    noteText(nextText) {
      text = nextText;
      const captured = ++version;
      timer?.cancel();
      if (!pptxGrammarLintAllowed({ enabled: options.getEnabled(), composing })) {
        diagnostics = [];
        publish();
        return;
      }
      timer = schedule(() => {
        timer = null;
        void (async () => {
          const lints = await options.requestLint(text);
          if (lints === null || !acceptPptxGrammarResult(captured, version)) {
            options.log?.({ phase: 'stale', version: captured, lintCount: lints?.length ?? 0 });
            return;
          }
          diagnostics = lints.map((lint, index) => ({
            id: `${captured}:${index}`,
            start: lint.span.start,
            end: lint.span.end,
            message: lint.message,
            suggestions: lint.suggestions,
            version: captured,
          }));
          options.log?.({
            phase: 'diagnostics',
            version: captured,
            textLength: text.length,
            diagnosticCount: diagnostics.length,
          });
          publish();
        })();
      }, debounceMs);
    },
    setComposing(next) {
      composing = next;
      if (next) {
        timer?.cancel();
        timer = null;
      }
    },
    setEnabled(enabled) {
      if (enabled) return;
      version += 1;
      timer?.cancel();
      timer = null;
      diagnostics = [];
      publish();
    },
    apply(id, suggestionIndex) {
      const diagnostic = diagnostics.find((item) => item.id === id);
      const suggestion = diagnostic?.suggestions[suggestionIndex];
      if (!diagnostic || !suggestion || !acceptPptxGrammarResult(diagnostic.version, version)) return null;
      text = replaceTextareaSpan(text, diagnostic.start, diagnostic.end, suggestion);
      diagnostics = [];
      version += 1;
      publish();
      return text;
    },
    clear() {
      version += 1;
      timer?.cancel();
      timer = null;
      diagnostics = [];
      publish();
    },
    diagnostics() {
      return diagnostics;
    },
  };
}
