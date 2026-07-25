/**
 * Section Break Commands
 * @packageDocumentation
 * @public
 */
import { Command } from 'prosemirror-state';

/**
 * Section Break Commands
 * @packageDocumentation
 * @public
 */

/**
 * Insert a "next page" section break at the cursor — starts a new section on a
 * new page.
 */
declare const insertSectionBreakNextPage: Command;
/**
 * Insert a "continuous" section break at the cursor — starts a new section on
 * the same page.
 */
declare const insertSectionBreakContinuous: Command;

export { insertSectionBreakContinuous, insertSectionBreakNextPage };
