/**
 * Superscript Mark Extension
 */

import { createMarkExtension } from '../create';
import { toggleMarkWithParagraphDefaults } from './markUtils';
import type { ExtensionContext, ExtensionRuntime } from '../types';

export const SuperscriptExtension = createMarkExtension({
  name: 'superscript',
  schemaMarkName: 'superscript',
  markSpec: {
    excludes: 'subscript',
    parseDOM: [{ tag: 'sup' }],
    toDOM() {
      return ['sup', 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    return {
      commands: {
        toggleSuperscript: () => toggleMarkWithParagraphDefaults(ctx.schema.marks.superscript),
      },
    };
  },
});
