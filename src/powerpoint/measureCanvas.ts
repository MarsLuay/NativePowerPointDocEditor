/**
 * Detached canvas for text measurement.
 *
 * Prefer Window.createEl / doc.win.createEl — Document.createEl on an SVG/XML
 * ownerDocument can append to a Document that already has a root and throw
 * HierarchyRequestError ("Only one element on document allowed").
 */
type CanvasHostWindow = Window & {
  createEl?: (tag: 'canvas') => HTMLCanvasElement;
};

export function createDetachedMeasureCanvas(ownerDocument?: Document | null): HTMLCanvasElement | null {
  const scopedWindow = (ownerDocument?.defaultView ?? (typeof window !== 'undefined' ? window : null)) as
    | CanvasHostWindow
    | null;
  if (scopedWindow && typeof scopedWindow.createEl === 'function') {
    return scopedWindow.createEl('canvas');
  }

  // Plain Chromium documents (including the headless PDF export harness) do
  // not expose Obsidian's createEl helper. Use the window's HTML document so
  // SVG/XML owner documents never receive a second document root.
  if (scopedWindow?.document && typeof scopedWindow.document.createElement === 'function') {
    try {
      return scopedWindow.document.createElement('canvas');
    } catch {
      // Keep the Obsidian doc.win fallback below for unusual host documents.
    }
  }

  const doc = ownerDocument ?? (typeof activeDocument !== 'undefined' ? activeDocument : null);
  if (!doc) return null;

  const win = doc.win as CanvasHostWindow | undefined;
  if (!win || typeof win.createEl !== 'function') return null;

  try {
    return win.createEl('canvas');
  } catch {
    if (typeof doc.createElement !== 'function') return null;
    try {
      return doc.createElement('canvas');
    } catch {
      return null;
    }
  }
}
