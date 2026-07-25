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

  const doc = ownerDocument ?? (typeof activeDocument !== 'undefined' ? activeDocument : null);
  if (!doc) return null;

  const win = doc.win as CanvasHostWindow | undefined;
  if (!win || typeof win.createEl !== 'function') return null;

  try {
    return win.createEl('canvas');
  } catch {
    return null;
  }
}
