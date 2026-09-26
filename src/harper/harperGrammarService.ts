import { HARPER_JS_VERSION, type HarperRuntimeContribution } from './harperRuntimeBudget';

export interface HarperGrammarSpan {
  start: number;
  end: number;
}

export interface HarperGrammarSuggestion {
  kind: 'replace' | 'remove' | 'insert-after';
  replacement: string;
}

export interface HarperGrammarLint {
  span: HarperGrammarSpan;
  message: string;
  suggestions: HarperGrammarSuggestion[];
}

export interface HarperWorkerLinter {
  setup(): Promise<void>;
  getDefaultLintConfig(): Promise<unknown>;
  lint(text: string): Promise<readonly HarperLintLike[]>;
  dispose(): Promise<void>;
}

interface HarperLintLike {
  span?: HarperGrammarSpan | (() => HarperGrammarSpan);
  message?: string | (() => string);
  suggestions?: readonly HarperSuggestionLike[] | (() => readonly HarperSuggestionLike[]);
}

interface HarperSuggestionLike {
  kind?: number | string | (() => number | string);
  get_replacement_text?: () => string;
  replacement?: string;
}

export interface HarperScheduleHandle {
  cancel(): void;
}

export interface HarperGrammarLogEntry {
  level: 'debug' | 'error';
  message: string;
  data: Record<string, unknown>;
}

export interface HarperGrammarServiceOptions {
  createLinter: () => HarperWorkerLinter;
  debounceMs?: number;
  schedule?: (callback: () => void, delayMs: number) => HarperScheduleHandle;
  now?: () => number;
  log?: (entry: HarperGrammarLogEntry) => void;
}

export interface HarperGrammarService {
  initialize(): Promise<void>;
  requestLint(text: string): Promise<HarperGrammarLint[] | null>;
  disable(): void;
  dispose(): Promise<void>;
}

const DEFAULT_DEBOUNCE_MS = 250;
const MAX_LOG_MESSAGE_CHARS = 180;

function boundedMessage(value: unknown): string {
  const text = value instanceof Error
    ? value.message
    : typeof value === 'string'
      ? value
      : 'unknown error';
  return text.length <= MAX_LOG_MESSAGE_CHARS ? text : `${text.slice(0, MAX_LOG_MESSAGE_CHARS - 1)}…`;
}

function readSpan(lint: HarperLintLike): HarperGrammarSpan {
  const span = typeof lint.span === 'function' ? lint.span() : lint.span;
  return {
    start: span?.start ?? 0,
    end: span?.end ?? 0,
  };
}

function readMessage(lint: HarperLintLike): string {
  if (typeof lint.message === 'function') return lint.message();
  return lint.message ?? '';
}

function suggestionKind(value: number | string | undefined): HarperGrammarSuggestion['kind'] {
  if (value === 1 || value === 'Remove' || value === 'remove') return 'remove';
  if (value === 2 || value === 'InsertAfter' || value === 'insert-after') return 'insert-after';
  return 'replace';
}

function readSuggestions(lint: HarperLintLike): HarperGrammarSuggestion[] {
  const suggestions = typeof lint.suggestions === 'function' ? lint.suggestions() : lint.suggestions ?? [];
  return suggestions.map((suggestion) => {
    const kindValue = typeof suggestion.kind === 'function' ? suggestion.kind() : suggestion.kind;
    return {
      kind: suggestionKind(kindValue),
      replacement: suggestion.get_replacement_text?.() ?? suggestion.replacement ?? '',
    };
  });
}

export function mapHarperLints(lints: readonly HarperLintLike[]): HarperGrammarLint[] {
  return lints.map((lint) => ({
    span: readSpan(lint),
    message: readMessage(lint),
    suggestions: readSuggestions(lint),
  }));
}

function defaultSchedule(callback: () => void, delayMs: number): HarperScheduleHandle {
  const timer = window.setTimeout(callback, delayMs);
  return { cancel: () => window.clearTimeout(timer) };
}

/**
 * One shared Harper worker for every editor view. Requests are debounced and
 * versioned so a slower lint cannot overwrite a newer one. Source text is
 * never written to logs.
 */
export function createHarperGrammarService(options: HarperGrammarServiceOptions): HarperGrammarService {
  const debounceMs = options.debounceMs ?? DEFAULT_DEBOUNCE_MS;
  const schedule = options.schedule ?? defaultSchedule;
  const now = options.now ?? (() => Date.now());
  const log = options.log ?? (() => undefined);
  let disabled = false;
  let disposed = false;
  let version = 0;
  let linter: HarperWorkerLinter | null = null;
  let starting: Promise<void> | null = null;
  let timer: HarperScheduleHandle | null = null;
  const waiters = new Map<number, (result: HarperGrammarLint[] | null) => void>();

  const settleOlder = (current: number) => {
    for (const [waitingVersion, resolve] of waiters) {
      if (waitingVersion === current) continue;
      resolve(null);
      waiters.delete(waitingVersion);
    }
  };

  const initialize = async (): Promise<void> => {
    if (disposed) {
      throw new Error('Harper grammar service is disposed.');
    }
    if (linter) return;
    if (starting) return starting;

    starting = (async () => {
      const startedAt = now();
      try {
        const created = options.createLinter();
        await created.setup();
        await created.getDefaultLintConfig();
        if (disposed) {
          await created.dispose();
          return;
        }
        linter = created;
        log({
          level: 'debug',
          message: 'Harper grammar worker ready',
          data: {
            phase: 'initialize',
            durationMs: now() - startedAt,
            workerReused: false,
          },
        });
      } catch (error) {
        log({
          level: 'error',
          message: 'Harper grammar worker failed',
          data: {
            phase: 'initialize',
            durationMs: now() - startedAt,
            error: boundedMessage(error),
          },
        });
        throw error;
      } finally {
        starting = null;
      }
    })();
    return starting;
  };

  return {
    initialize,
    requestLint(text: string): Promise<HarperGrammarLint[] | null> {
      if (disabled || disposed) return Promise.resolve(null);
      const requestVersion = ++version;
      timer?.cancel();
      settleOlder(requestVersion);
      return new Promise((resolve) => {
        waiters.set(requestVersion, resolve);
        timer = schedule(() => {
          timer = null;
          void (async () => {
            const startedAt = now();
            try {
              await initialize();
              if (requestVersion !== version || !linter) {
                resolve(null);
                waiters.delete(requestVersion);
                return;
              }
              const raw = await linter.lint(text);
              if (requestVersion !== version) {
                resolve(null);
                waiters.delete(requestVersion);
                return;
              }
              const lints = mapHarperLints(raw);
              log({
                level: 'debug',
                message: 'Harper grammar lint finished',
                data: {
                  phase: 'lint',
                  version: requestVersion,
                  durationMs: now() - startedAt,
                  lintCount: lints.length,
                },
              });
              resolve(lints);
            } catch (error) {
              log({
                level: 'error',
                message: 'Harper grammar worker failed',
                data: {
                  phase: 'lint',
                  version: requestVersion,
                  durationMs: now() - startedAt,
                  error: boundedMessage(error),
                },
              });
              resolve(null);
            } finally {
              waiters.delete(requestVersion);
            }
          })();
        }, debounceMs);
      });
    },
    disable(): void {
      disabled = true;
      version += 1;
      timer?.cancel();
      timer = null;
      settleOlder(version);
    },
    async dispose(): Promise<void> {
      if (disposed) return;
      disposed = true;
      disabled = true;
      version += 1;
      timer?.cancel();
      timer = null;
      settleOlder(version);
      if (starting) {
        try {
          await starting;
        } catch {
          // Initialization failure is already logged.
        }
      }
      const active = linter;
      linter = null;
      await active?.dispose();
    },
  };
}

export function describeHarperPackagingRefusal(contribution: HarperRuntimeContribution): string {
  return `Harper ${HARPER_JS_VERSION} ${contribution.artifact} is ${contribution.rawBytes} bytes (${contribution.gzipBytes} gzip). Not embedded or materialized: fileLimit=${contribution.withinFileLimit}, mainEmbed=${contribution.embedFitsMain}.`;
}
