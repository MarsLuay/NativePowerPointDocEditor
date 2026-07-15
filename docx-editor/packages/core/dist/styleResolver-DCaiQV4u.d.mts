import { P as ParagraphFormatting, T as TextFormatting } from './formatting-DFtuRFQY.mjs';
import { S as StyleDefinitions, a as Style, D as DocDefaults } from './styles-BNjUANte.mjs';

/**
 * Style Resolver for ProseMirror Editor
 *
 * Resolves OOXML style definitions to final paragraph and run properties.
 * Handles the cascade:
 * 1. Document defaults (docDefaults)
 * 2. Normal style (if no explicit styleId)
 * 3. Style chain (basedOn inheritance - already resolved by styleParser)
 * 4. Inline properties
 *
 * Based on ECMA-376 style cascade rules.
 */

/**
 * Resolved style properties ready for rendering
 */
interface ResolvedParagraphStyle {
    /** Paragraph formatting (alignment, spacing, indentation, etc.) */
    paragraphFormatting?: ParagraphFormatting;
    /** Default run formatting from the style */
    runFormatting?: TextFormatting;
}
/**
 * StyleResolver provides efficient access to resolved style properties
 */
declare class StyleResolver {
    private readonly stylesById;
    private readonly docDefaults;
    private readonly defaultParagraphStyle;
    private readonly defaultTableStyle;
    private readonly defaultCharacterStyle;
    constructor(styleDefinitions: StyleDefinitions | undefined);
    /**
     * Get a style by ID
     */
    getStyle(styleId: string): Style | undefined;
    /**
     * Get all available paragraph styles (for toolbar dropdown)
     */
    getParagraphStyles(): Style[];
    /**
     * Whether a paragraph style with the given id is defined in the
     * document's `styles.xml`. Used by the agent toolkit to refuse
     * `set_paragraph_style({ styleId: 'NoSuchStyle' })` instead of
     * silently writing an invalid `<w:pStyle>` reference.
     */
    hasParagraphStyle(styleId: string): boolean;
    /**
     * Resolve paragraph style properties, including docDefaults cascade
     *
     * @param styleId - The style ID to resolve (e.g., 'Heading1', 'Normal')
     * @returns Resolved paragraph and run formatting
     */
    resolveParagraphStyle(styleId: string | undefined | null): ResolvedParagraphStyle;
    /**
     * Resolve the style applied to the paragraph that follows one styled with
     * `styleId` when the user presses Enter (OOXML `w:next`, §17.7.4.10).
     *
     * Returns null when the style has no `w:next`, when `next` points back at
     * the same style (the common heading-stays-heading case is handled by the
     * caller), or when the style is unknown.
     */
    getNextStyleId(styleId: string | undefined | null): string | null;
    /**
     * Get all available table styles (for style gallery)
     */
    getTableStyles(): Style[];
    /**
     * Resolve run (character) style properties
     *
     * @param styleId - The character style ID to resolve
     * @returns Resolved text formatting
     */
    resolveRunStyle(styleId: string | undefined | null): TextFormatting | undefined;
    /**
     * Get a character style's own properties WITHOUT docDefaults.
     * Used when the caller already has docDefaults applied (e.g., from paragraph style resolution).
     * This prevents docDefault fonts from incorrectly overriding paragraph style fonts.
     */
    getRunStyleOwnProperties(styleId: string | undefined | null): TextFormatting | undefined;
    /**
     * Get document defaults
     */
    getDocDefaults(): DocDefaults | undefined;
    /**
     * Get default paragraph style (usually "Normal")
     */
    getDefaultParagraphStyle(): Style | undefined;
    /**
     * Get the default table style (the one marked `w:default="1"`).
     *
     * Per ECMA-376 §17.7.4.18, tables that don't specify a `w:tblStyle`
     * inherit from this style. The styleId varies by document language
     * ("Normal Table", "TableNormal", "Tabelanormal", etc.) — find it by
     * the parsed `default` flag, not by name.
     */
    getDefaultTableStyle(): Style | undefined;
    /**
     * Get the default character style (the one marked `w:default="1"`).
     *
     * Per ECMA-376 §17.7.4.18, runs without an explicit `w:rStyle` reference
     * inherit from this style. The styleId varies by document language
     * ("Default Paragraph Font", "FontePadrao", "Fontepargpadro" in
     * Portuguese fixtures, etc.) — find it by the parsed `default` flag.
     */
    getDefaultCharacterStyle(): Style | undefined;
    /**
     * Check if a style exists
     */
    hasStyle(styleId: string): boolean;
    private findDefaultStyle;
    private mergeStyleIntoResult;
    /**
     * Merge paragraph formatting (source overrides target)
     */
    private mergeParagraphFormatting;
    private mergeTextFormatting;
}
/**
 * Create a style resolver from document's style definitions
 */
declare function createStyleResolver(styleDefinitions: StyleDefinitions | undefined): StyleResolver;

export { type ResolvedParagraphStyle as R, StyleResolver as S, createStyleResolver as c };
