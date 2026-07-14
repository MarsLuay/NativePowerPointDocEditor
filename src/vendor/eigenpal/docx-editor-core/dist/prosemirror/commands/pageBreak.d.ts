/**
 * Page Break Commands
 * @packageDocumentation
 * @public
 */
import { Command } from 'prosemirror-state';

/**
 * Page Break Commands
 * @packageDocumentation
 * @public
 */

/**
 * Insert a page break at the current cursor position.
 * Always ensures a paragraph follows the page break and places the cursor there.
 */
declare const insertPageBreak: Command;

export { insertPageBreak };
