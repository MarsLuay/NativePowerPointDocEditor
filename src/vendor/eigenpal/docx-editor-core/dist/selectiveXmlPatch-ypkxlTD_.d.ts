interface PatchValidationResult {
    safe: boolean;
    reason?: string;
}
/**
 * Validate that a selective patch can be safely applied.
 *
 * Checks:
 * - All changed paraIds exist in original XML (exactly once)
 * - All changed paraIds exist in serialized XML (exactly once)
 * - Paragraph count matches between original and serialized
 */
declare function validatePatchSafety(originalXml: string, serializedXml: string, changedIds: Set<string>): PatchValidationResult;
/**
 * Build a patched document.xml by splicing new paragraph XML into
 * the original at the correct offsets. Only changed paragraphs
 * are replaced; everything else is preserved byte-for-byte.
 *
 * Returns null if any step fails.
 */
declare function buildPatchedDocumentXml(originalXml: string, serializedXml: string, changedIds: Set<string>): string | null;

export { buildPatchedDocumentXml as b, validatePatchSafety as v };
