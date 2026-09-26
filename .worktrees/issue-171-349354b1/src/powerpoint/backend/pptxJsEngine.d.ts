// Ambient types for the generated pure-JS PPTX engine fallback
// (src/powerpoint/backend/pptxJsEngine.mjs). The module is loaded lazily and only when the
// runtime lacks WebAssembly GC. Its exports mirror the Wasm module's exports, so
// a loose signature is sufficient — the renderer consumes them via its own typed
// `exports` surface.
declare module '*/pptxJsEngine.mjs' {
  export type PptxJsEngineExports = Record<string, (...args: never[]) => unknown>;
  export function createPptxJsEngine(): PptxJsEngineExports;
}
