// Runtime/environment helpers for the PowerPoint view: error normalization and
// WebAssembly GC support detection. Extracted from NativePowerPointView.ts.

export function cleanError(error: unknown): string {
  return error instanceof Error ? error.message : String(error || 'Unknown error');
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
  const message = error instanceof Error ? error.message : String(error || '');
  return /WebAssembly GC support|Wasm init failed/i.test(message);
}

export function getChromiumVersion(): string | null {
  try {
    // `process` is a Node/Electron global (not a DOM/popout `window` property), and
    // this also runs headless in Node, so `globalThis` is the correct portable access.
    // eslint-disable-next-line obsidianmd/no-global-this
    const versions = (globalThis as { process?: { versions?: { chrome?: string } } }).process?.versions;
    return versions?.chrome ?? null;
  } catch {
    return null;
  }
}
