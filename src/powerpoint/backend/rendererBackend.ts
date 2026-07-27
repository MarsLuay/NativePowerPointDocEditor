import type { PptxRenderer } from 'pptx-svg';
import { FontFidelity } from '../../FontFidelity';
import { shouldForceJsBackend } from '../forceJsBackend';

/** Renderer augmented with the build-time `initJsBackend` patch (see esbuild.config.mjs). */
interface JsBackendCapableRenderer {
  initJsBackend(engine: unknown): void;
}

interface PptxWasmRendererModule {
  PptxRenderer: typeof PptxRenderer;
  wasmBytes: Uint8Array;
}

export type PptxRendererBackend = 'wasm-gc' | 'js';

export interface PresentationRenderer {
  renderer: PptxRenderer;
  rendererBackend: PptxRendererBackend;
  fontFidelity: FontFidelity;
  slideCount: number;
}

function isWasmGcUnsupportedError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : typeof error === 'string' ? error : '';
  return /WebAssembly GC support|Wasm init failed/i.test(message);
}

async function initJsBackend(renderer: PptxRenderer): Promise<void> {
  const { createPptxJsEngine } = await import('./pptxJsEngine.mjs');
  (renderer as unknown as JsBackendCapableRenderer).initJsBackend(createPptxJsEngine());
}

/**
 * Keep the WebAssembly-GC renderer out of main.js. Obsidian Sync Standard has
 * a 5 MB per-file limit, and this renderer is only needed after a PowerPoint
 * file is opened. esbuild rewrites this source-relative import to the emitted
 * plugin-root artifact `pptx-wasm-renderer.mjs`.
 */
async function loadWasmRenderer(): Promise<PptxWasmRendererModule> {
  return await import('./pptxWasmRenderer.mjs');
}

async function initRendererBackend(renderer: PptxRenderer): Promise<PptxRendererBackend> {
  if (shouldForceJsBackend()) {
    await initJsBackend(renderer);
    return 'js';
  }

  try {
    const { wasmBytes } = await loadWasmRenderer();
    await renderer.init(wasmBytes);
    return 'wasm-gc';
  } catch (error) {
    if (!isWasmGcUnsupportedError(error)) throw error;
    await initJsBackend(renderer);
    return 'js';
  }
}

/**
 * Creates the renderer behind PresentationEngine's facade.
 * The renderer is derived state; the facade's lossless package buffer remains authoritative.
 */
export async function createPresentationRenderer(buffer: ArrayBuffer): Promise<PresentationRenderer> {
  const fontFidelity = new FontFidelity();
  const { PptxRenderer } = await loadWasmRenderer();
  const renderer = new PptxRenderer({
    logLevel: 'error',
    fontFallbacks: fontFidelity.getRendererFallbacks(),
    measureText: (text, fontFace, fontSizePx) => fontFidelity.measureText(text, fontFace, fontSizePx)
  });

  const rendererBackend = await initRendererBackend(renderer);
  const { slideCount } = await renderer.loadPptx(buffer);
  return { renderer, rendererBackend, fontFidelity, slideCount };
}
