/**
 * Derived paragraph DOM classes for Word-like image layout.
 *
 * Replaces CSS `:has()` selectors that force broad style invalidation:
 * - `docx-p-block-image` — sole child is a block image
 * - `docx-p-has-float` — paragraph contains a floated image
 *
 * Stamped in `paragraph` toDOM for initial paint, and kept fresh by a view
 * plugin when image wrap attrs change without recreating the `<p>` node.
 */

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { ImageAttrs } from '../../schema/nodes';

export const DOC_X_P_BLOCK_IMAGE_CLASS = 'docx-p-block-image';
export const DOC_X_P_HAS_FLOAT_CLASS = 'docx-p-has-float';

export const paragraphImageLayoutKey = new PluginKey('paragraphImageLayout');

function isFloatingImage(node: PMNode): boolean {
  if (node.type.name !== 'image') return false;
  const attrs = node.attrs as ImageAttrs;
  return attrs.displayMode === 'float' && !!attrs.cssFloat && attrs.cssFloat !== 'none';
}

function isBlockImage(node: PMNode): boolean {
  if (node.type.name !== 'image') return false;
  return (node.attrs as ImageAttrs).displayMode === 'block';
}

/**
 * @returns class tokens to add on the paragraph DOM element
 */
export function getParagraphImageLayoutClasses(node: PMNode): string[] {
  if (node.type.name !== 'paragraph') return [];

  const classes: string[] = [];

  if (node.childCount === 1 && node.firstChild && isBlockImage(node.firstChild)) {
    classes.push(DOC_X_P_BLOCK_IMAGE_CLASS);
  }

  let hasFloat = false;
  node.forEach((child) => {
    if (isFloatingImage(child)) {
      hasFloat = true;
    }
  });
  if (hasFloat) {
    classes.push(DOC_X_P_HAS_FLOAT_CLASS);
  }

  return classes;
}

export function applyParagraphImageLayoutClasses(dom: HTMLElement, node: PMNode): void {
  const classes = new Set(getParagraphImageLayoutClasses(node));
  dom.classList.toggle(DOC_X_P_BLOCK_IMAGE_CLASS, classes.has(DOC_X_P_BLOCK_IMAGE_CLASS));
  dom.classList.toggle(DOC_X_P_HAS_FLOAT_CLASS, classes.has(DOC_X_P_HAS_FLOAT_CLASS));
}

export function syncParagraphImageLayoutClasses(view: EditorView): void {
  view.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph') {
      return true;
    }
    const dom = view.nodeDOM(pos);
    if (dom instanceof HTMLElement) {
      applyParagraphImageLayoutClasses(dom, node);
    }
    return false;
  });
}

export function createParagraphImageLayoutPlugin(): Plugin {
  return new Plugin({
    key: paragraphImageLayoutKey,
    view(editorView) {
      syncParagraphImageLayoutClasses(editorView);
      return {
        update(view, prevState) {
          if (view.state.doc.eq(prevState.doc)) return;
          syncParagraphImageLayoutClasses(view);
        },
      };
    },
  });
}
