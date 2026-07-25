// Runtime/environment helpers for the PowerPoint view: error normalization and
// WebAssembly GC support detection. Extracted from NativePowerPointView.ts.

export function cleanError(error: unknown): string {
  if (error instanceof Error) {
    return error.message || 'Unknown error';
  }
  if (typeof error === 'string') {
    return error || 'Unknown error';
  }
  if (typeof error === 'number' || typeof error === 'boolean' || typeof error === 'bigint') {
    return String(error);
  }
  if (error && typeof error === 'object') {
    try {
      return JSON.stringify(error) || 'Unknown error';
    } catch {
      return 'Unknown error';
    }
  }
  return 'Unknown error';
}

/**
 * The PPTX renderer is a MoonBit module compiled to the WebAssembly GC target,
 * which Chromium only enables by default from version 119 (the engine bundled
 * with Obsidian's installer, not the in-app app version). Older installers fail
 * every instantiation tier in `pptx-svg`, surfacing a raw "Wasm init failed —
 * requires WebAssembly GC support" error. DOCX uses a pure-JS path and is
 * unaffected, so we translate this specific failure into actionable guidance.
 */
export function isWasmGcUnsupportedError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : typeof error === 'string' ? error : '';
  return /WebAssembly GC support|Wasm init failed/i.test(message);
}

/** Yields so status/progress DOM updates can paint before heavy work continues. */
export function flushUi(): Promise<void> {
  const requestAnimationFrame =
		typeof window !== 'undefined' ? window.requestAnimationFrame.bind(window) : undefined;
  if (typeof requestAnimationFrame === 'function') {
    return new Promise((resolve) => {
      requestAnimationFrame(() => {
        requestAnimationFrame(() => resolve());
      });
    });
  }

  if (typeof queueMicrotask === 'function') {
    return new Promise((resolve) => queueMicrotask(resolve));
  }

  return Promise.resolve();
}
