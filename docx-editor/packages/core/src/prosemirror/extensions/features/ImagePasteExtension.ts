/**
 * Image Paste / Drop Extension — handles image files from clipboard paste and
 * external OS file drops.
 *
 * When an image file is present on the clipboard or dropped from Finder/Explorer,
 * this intercepts the event, reads the image data, and inserts an image node at
 * the caret (paste) or under the pointer (drop). Internal in-doc image
 * reposition drags (`pm-image-dragging`) are left to ImageDragExtension.
 */

import { Plugin, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { createExtension } from '../create';
import type { ExtensionRuntime } from '../types';
import { getClipboardImageFiles, dataTransferLooksLikeExternalImageDrop } from '../../../utils/clipboard';

const MAX_INLINE_IMAGE_WIDTH = 612; // ~6.375 inches at 96 DPI

async function readFileAsDataUrl(file: File): Promise<string> {
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error ?? new Error('Failed to read image file'));
    reader.readAsDataURL(file);
  });
}

async function loadImageSize(src: string): Promise<{ width: number; height: number }> {
  return await new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || 1, height: img.naturalHeight || 1 });
    img.onerror = () => reject(new Error('Failed to load pasted image'));
    img.src = src;
  });
}

/** True while ImageDragExtension is repositioning an existing in-doc image. */
export function isInternalImageDragging(view: EditorView): boolean {
  return view.dom.classList.contains('pm-image-dragging');
}

/**
 * Resolve the document position under a pointer for an external image drop.
 * Returns null when coords cannot be mapped into the document.
 */
export function resolveExternalImageDropPos(
  view: EditorView,
  clientX: number,
  clientY: number
): number | null {
  const coords = view.posAtCoords({ left: clientX, top: clientY });
  if (!coords) return null;
  const pos = Math.max(0, Math.min(coords.pos, view.state.doc.content.size));
  return TextSelection.near(view.state.doc.resolve(pos)).from;
}

/**
 * Insert image files inline starting at `startPos` (or the current selection).
 * Position is captured before async decode so later selection moves cannot
 * relocate the insert.
 */
export async function insertImageFiles(
  view: EditorView,
  files: File[],
  startPos?: number
): Promise<void> {
  const imageType = view.state.schema.nodes.image;
  if (!imageType || files.length === 0) return;

  let insertPos = startPos ?? view.state.selection.from;

  for (const file of files) {
    let dataUrl: string;
    try {
      dataUrl = await readFileAsDataUrl(file);
    } catch {
      continue;
    }

    let naturalWidth = 1;
    let naturalHeight = 1;
    try {
      ({ width: naturalWidth, height: naturalHeight } = await loadImageSize(dataUrl));
    } catch {
      // Fall back to a safe minimal size if the image can't be decoded
      naturalWidth = 1;
      naturalHeight = 1;
    }

    let width = naturalWidth;
    let height = naturalHeight;

    if (width > MAX_INLINE_IMAGE_WIDTH) {
      const scale = MAX_INLINE_IMAGE_WIDTH / width;
      width = MAX_INLINE_IMAGE_WIDTH;
      height = Math.max(1, Math.round(height * scale));
    }

    // Clamp against the live doc in case a concurrent edit moved content.
    insertPos = Math.max(0, Math.min(insertPos, view.state.doc.content.size));

    const imageNode = imageType.create({
      src: dataUrl,
      alt: file.name,
      width,
      height,
      rId: `rId_img_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      wrapType: 'inline',
      displayMode: 'inline',
    });

    const tr = view.state.tr.insert(insertPos, imageNode);
    insertPos += imageNode.nodeSize;
    tr.setSelection(TextSelection.create(tr.doc, insertPos));
    view.dispatch(tr.scrollIntoView());
  }

  view.focus();
}

function handleExternalImageDragOver(view: EditorView, event: Event): boolean {
  if (isInternalImageDragging(view)) return false;
  if (!view.state.schema.nodes.image) return false;

  const dragEvent = event as DragEvent;
  if (!dataTransferLooksLikeExternalImageDrop(dragEvent.dataTransfer)) {
    return false;
  }

  dragEvent.preventDefault();
  if (dragEvent.dataTransfer) {
    dragEvent.dataTransfer.dropEffect = 'copy';
  }
  return true;
}

function handleExternalImageDrop(view: EditorView, event: Event): boolean {
  if (isInternalImageDragging(view)) return false;
  if (!view.state.schema.nodes.image) return false;

  const dragEvent = event as DragEvent;
  const imageFiles = getClipboardImageFiles(dragEvent.dataTransfer);
  if (imageFiles.length === 0) return false;

  const insertPos = resolveExternalImageDropPos(
    view,
    dragEvent.clientX,
    dragEvent.clientY
  );
  if (insertPos == null) return false;

  dragEvent.preventDefault();
  void insertImageFiles(view, imageFiles, insertPos).catch(() => undefined);
  return true;
}

export const ImagePasteExtension = createExtension({
  name: 'imagePaste',
  onSchemaReady(_ctx): ExtensionRuntime {
    const plugin = new Plugin({
      props: {
        handleDOMEvents: {
          paste(view, event) {
            const clipboardEvent = event as ClipboardEvent;
            const imageFiles = getClipboardImageFiles(clipboardEvent.clipboardData);

            if (imageFiles.length === 0) {
              return false;
            }

            if (!view.state.schema.nodes.image) {
              return false;
            }

            clipboardEvent.preventDefault();
            void insertImageFiles(view, imageFiles).catch(() => undefined);
            return true;
          },
          dragover: handleExternalImageDragOver,
          drop: handleExternalImageDrop,
        },
      },
    });

    return { plugins: [plugin] };
  },
});
