import { Document } from './types/document.mjs';
import { D as DocumentBody, P as Paragraph } from './content-BZ9rYecc.mjs';

/**
 * Variable Detector Utility
 *
 * Scans a DOCX document for template variables in the format {variable_name}
 * (standard docxtemplater syntax).
 * Returns a unique, sorted list of variable names found in the document.
 */

/**
 * Result of variable detection
 */
interface VariableDetectionResult {
    /** Unique variable names sorted alphabetically */
    variables: string[];
    /** Total count of variable occurrences */
    totalOccurrences: number;
    /** Variables by location */
    byLocation: {
        body: string[];
        headers: string[];
        footers: string[];
        footnotes: string[];
        endnotes: string[];
        textBoxes: string[];
    };
    /** Variable occurrences with positions */
    occurrences: VariableOccurrence[];
}
/**
 * A single variable occurrence with location info
 */
interface VariableOccurrence {
    /** Variable name (without braces) */
    name: string;
    /** Location type */
    location: 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'textBox';
    /** Paragraph index within location */
    paragraphIndex?: number;
    /** Section index (for headers/footers) */
    sectionIndex?: number;
}
/**
 * Detect all template variables in a document
 *
 * @param doc - The parsed document
 * @returns Array of unique variable names sorted alphabetically
 */
declare function detectVariables(doc: Document): string[];
/**
 * Detect variables with detailed information
 *
 * @param doc - The parsed document
 * @returns Detailed detection result
 */
declare function detectVariablesDetailed(doc: Document): VariableDetectionResult;
/**
 * Detect variables in document body
 */
declare function detectVariablesInBody(body: DocumentBody): string[];
/**
 * Detect variables in a paragraph
 */
declare function detectVariablesInParagraph(paragraph: Paragraph): string[];
/**
 * Extract variable names from text
 *
 * @param text - The text to search
 * @returns Array of variable names (without braces)
 */
declare function extractVariablesFromText(text: string): string[];
/**
 * Check if text contains template variables
 */
declare function hasTemplateVariables(text: string): boolean;
/**
 * Check if a variable name is valid
 */
declare function isValidVariableName(name: string): boolean;
/**
 * Sanitize a variable name
 */
declare function sanitizeVariableName(name: string): string;
/**
 * Format a variable name with braces (standard docxtemplater syntax)
 */
declare function formatVariable(name: string): string;
/**
 * Parse a variable string to get the name
 */
declare function parseVariable(variable: string): string | null;
/**
 * Replace variables in text with values
 *
 * @param text - The text containing variables
 * @param values - Map of variable name to replacement value
 * @returns Text with variables replaced
 */
declare function replaceVariables(text: string, values: Record<string, string>): string;
/**
 * Replace all variables in text with a placeholder
 *
 * @param text - The text containing variables
 * @param placeholder - Placeholder to use (default: empty string)
 * @returns Text with variables replaced
 */
declare function removeVariables(text: string, placeholder?: string): string;
/**
 * Check if document has any template variables
 */
declare function documentHasVariables(doc: Document): boolean;

export { type VariableDetectionResult as V, type VariableOccurrence as a, detectVariablesDetailed as b, detectVariablesInBody as c, detectVariables as d, detectVariablesInParagraph as e, documentHasVariables as f, extractVariablesFromText as g, formatVariable as h, hasTemplateVariables as i, isValidVariableName as j, replaceVariables as k, parseVariable as p, removeVariables as r, sanitizeVariableName as s };
