import { Document } from './types/document.mjs';

/**
 * Selective Save Module
 *
 * Orchestrates selective XML patching for the save flow.
 * Serializes full document.xml, validates patch safety, builds patched XML,
 * and calls applyUpdatesToZip() to produce the final DOCX.
 *
 * Returns null on any failure, signaling the caller to fall back to full repack.
 */

interface SelectiveSaveOptions {
    /** Changed paragraph IDs to selectively patch */
    changedParaIds: Set<string>;
    /** Whether structural changes occurred (paragraph add/delete) */
    structuralChange: boolean;
    /** Whether any changes affected paragraphs without paraId */
    hasUntrackedChanges: boolean;
}
/**
 * Attempt a selective save — patch only changed paragraphs in document.xml.
 * Also updates comments, headers/footers, and core properties so that
 * all document parts stay in sync even when only paragraphs are patched.
 *
 * Returns the saved ArrayBuffer, or null if selective save is not possible
 * (caller should fall back to full repack).
 */
declare function attemptSelectiveSave(doc: Document, originalBuffer: ArrayBuffer, options: SelectiveSaveOptions): Promise<ArrayBuffer | null>;

export { type SelectiveSaveOptions as S, attemptSelectiveSave as a };
