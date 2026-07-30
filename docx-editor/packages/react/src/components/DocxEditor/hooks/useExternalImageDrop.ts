/**
 * External OS/Finder image file drop onto the painted pages surface.
 *
 * PagedEditor keeps ProseMirror off-screen; drop/dragover on view.dom never
 * fires under the pointer. This hook attaches to the visible editor root and
 * inserts via getPositionFromMouse + insertImageFromFile({ pos }).
 */

import { useCallback } from 'react';
import type { DragEvent as ReactDragEvent, RefObject } from 'react';
import {
  dataTransferLooksLikeExternalImageDrop,
  getClipboardImageFiles,
} from '@npde/docx-editor-core/utils';
import { insertImageFromFile } from '@npde/docx-editor-core/prosemirror/commands';
import type { HiddenProseMirrorRef } from '../HiddenProseMirror';

export interface UseExternalImageDropOptions {
  hiddenPMRef: RefObject<HiddenProseMirrorRef | null>;
  getPositionFromMouse: (clientX: number, clientY: number) => number | null;
  readOnly?: boolean;
}

function insertImageFromFileAtPos(
  view: NonNullable<ReturnType<HiddenProseMirrorRef['getView']>>,
  file: File,
  pos: number
): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    insertImageFromFile(view, file, {
      pos,
      onInserted: (dimensions) => resolve(dimensions),
      onError: (error) => reject(error),
    });
  });
}

export function useExternalImageDrop({
  hiddenPMRef,
  getPositionFromMouse,
  readOnly = false,
}: UseExternalImageDropOptions) {
  const handleDragOver = useCallback(
    (event: ReactDragEvent) => {
      if (readOnly) return;
      if (!dataTransferLooksLikeExternalImageDrop(event.dataTransfer)) return;

      const view = hiddenPMRef.current?.getView();
      if (!view?.state.schema.nodes.image) return;
      if (view.dom.classList.contains('pm-image-dragging')) return;

      event.preventDefault();
      event.stopPropagation();
      event.dataTransfer.dropEffect = 'copy';
    },
    [hiddenPMRef, readOnly]
  );

  const handleDrop = useCallback(
    (event: ReactDragEvent) => {
      if (readOnly) return;

      const view = hiddenPMRef.current?.getView();
      if (!view?.state.schema.nodes.image) return;
      if (view.dom.classList.contains('pm-image-dragging')) return;

      const imageFiles = getClipboardImageFiles(event.dataTransfer);
      if (imageFiles.length === 0) return;

      const insertPos = getPositionFromMouse(event.clientX, event.clientY);
      if (insertPos == null) return;

      event.preventDefault();
      event.stopPropagation();

      void (async () => {
        let pos = insertPos;
        for (const file of imageFiles) {
          try {
            await insertImageFromFileAtPos(view, file, pos);
            // Inline image nodes are atoms (nodeSize 1). Advance for multi-drop.
            pos += 1;
          } catch {
            // Skip undecodable files; continue remaining drops.
          }
        }
        view.focus();
      })();
    },
    [getPositionFromMouse, hiddenPMRef, readOnly]
  );

  return { handleDragOver, handleDrop };
}
