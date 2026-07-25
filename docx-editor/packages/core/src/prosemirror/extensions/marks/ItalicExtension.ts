/**
 * Italic Mark Extension
 */

import { createMarkExtension } from '../create';
import { toggleMarkWithParagraphDefaults } from './markUtils';
import type { ExtensionContext, ExtensionRuntime } from '../types';

export const ItalicExtension = createMarkExtension({
  name: 'italic',
  schemaMarkName: 'italic',
  markSpec: {
    parseDOM: [
      { tag: 'i' },
      { tag: 'em' },
      {
        style: 'font-style',
        getAttrs: (value) => (value === 'italic' ? null : false),
      },
    ],
    toDOM() {
      return ['em', 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    return {
      commands: {
        toggleItalic: () => toggleMarkWithParagraphDefaults(ctx.schema.marks.italic),
      },
      keyboardShortcuts: {
        'Mod-i': toggleMarkWithParagraphDefaults(ctx.schema.marks.italic),
      },
    };
  },
});
