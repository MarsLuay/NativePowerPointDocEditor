/**
 * Subscript Mark Extension
 */

import { createMarkExtension } from '../create';
import { toggleMarkWithParagraphDefaults } from './markUtils';
import type { ExtensionContext, ExtensionRuntime } from '../types';

export const SubscriptExtension = createMarkExtension({
  name: 'subscript',
  schemaMarkName: 'subscript',
  markSpec: {
    excludes: 'superscript',
    parseDOM: [{ tag: 'sub' }],
    toDOM() {
      return ['sub', 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    return {
      commands: {
        toggleSubscript: () => toggleMarkWithParagraphDefaults(ctx.schema.marks.subscript),
      },
    };
  },
});
