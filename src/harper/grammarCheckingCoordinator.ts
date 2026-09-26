import type { HarperGrammarLint, HarperGrammarService } from './harperGrammarService';

export interface GrammarCheckingCoordinator {
  setEnabled(enabled: boolean): void;
  requestLint(text: string): Promise<HarperGrammarLint[] | null>;
  dispose(): Promise<void>;
}

/**
 * Plugin-level on/off switch for the single shared grammar worker.
 * Disabling drops pending work and clears diagnostics. Enabling continues
 * on the same worker instead of creating one per editor.
 */
export function createGrammarCheckingCoordinator(options: {
  service: HarperGrammarService;
  clearDiagnostics: () => void;
}): GrammarCheckingCoordinator {
  let enabled = true;

  return {
    setEnabled(nextEnabled) {
      enabled = nextEnabled;
      if (nextEnabled) {
        options.service.enable();
        return;
      }
      options.clearDiagnostics();
      options.service.disable();
    },
    requestLint(text) {
      if (!enabled) return Promise.resolve(null);
      return options.service.requestLint(text);
    },
    dispose() {
      return options.service.dispose();
    },
  };
}
