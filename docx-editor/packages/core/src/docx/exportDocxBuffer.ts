/**
 * Export a Document model to a DOCX ArrayBuffer without DocumentAgent.
 *
 * Same semantics as the former DocumentAgent.toBuffer path: prefer selective
 * save against originalBuffer when options are provided, else full repack.
 */

import type { Document } from '../types/document';
import { createDocx, repackDocx } from './rezip';
import { attemptSelectiveSave, type SelectiveSaveOptions } from './selectiveSave';

export interface ExportDocxBufferOptions {
  selective?: SelectiveSaveOptions;
}

/**
 * Pack `doc` to DOCX bytes. Mutates `doc.originalBuffer` on successful save
 * so subsequent selective saves patch against the latest baseline.
 */
export async function exportDocxBuffer(
  doc: Document,
  options?: ExportDocxBufferOptions
): Promise<ArrayBuffer> {
  if (doc.originalBuffer) {
    if (options?.selective) {
      const result = await attemptSelectiveSave(doc, doc.originalBuffer, options.selective);
      if (result) {
        doc.originalBuffer = result;
        return result;
      }
    }
    const repacked = await repackDocx(doc);
    doc.originalBuffer = repacked;
    return repacked;
  }
  return createDocx(doc);
}
