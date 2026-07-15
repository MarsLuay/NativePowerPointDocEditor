import { T as TextFormatting, P as ParagraphFormatting, a as TableFormatting, b as TableRowFormatting, c as TableCellFormatting } from './formatting-DFtuRFQY.mjs';

/**
 * Styles, Theme, Font Table, Relationships & Media Types
 *
 * Types for document-level definitions that don't form the content tree.
 */

/**
 * Style type
 */
type StyleType = 'paragraph' | 'character' | 'numbering' | 'table';
/**
 * Style definition from `styles.xml` — a named, reusable bundle of
 * paragraph and/or character formatting. Word's "Heading 1", "Normal",
 * "Title", and "List Bullet" are styles; user-defined styles look the
 * same. `basedOn` chains styles for inheritance; `link` pairs a paragraph
 * style with a matching character style.
 *
 * See ECMA-376 §17.7.4.
 */
interface Style {
    /** Style ID */
    styleId: string;
    /** Style type */
    type: StyleType;
    /** Display name */
    name?: string;
    /** Based on style ID */
    basedOn?: string;
    /** Next style after Enter (for paragraph styles) */
    next?: string;
    /** Linked style (paragraph/character pair) */
    link?: string;
    /** UI sort priority */
    uiPriority?: number;
    /** Hidden from UI */
    hidden?: boolean;
    /** Semi-hidden from UI */
    semiHidden?: boolean;
    /** Unhide when used */
    unhideWhenUsed?: boolean;
    /** Quick format in gallery */
    qFormat?: boolean;
    /** Is default style */
    default?: boolean;
    /** Personal style (custom) */
    personal?: boolean;
    /** Paragraph properties (for paragraph/table styles) */
    pPr?: ParagraphFormatting;
    /** Run properties */
    rPr?: TextFormatting;
    /** Table properties (for table styles) */
    tblPr?: TableFormatting;
    /** Table row properties */
    trPr?: TableRowFormatting;
    /** Table cell properties */
    tcPr?: TableCellFormatting;
    /** Conditional table style parts */
    tblStylePr?: Array<{
        type: 'band1Horz' | 'band1Vert' | 'band2Horz' | 'band2Vert' | 'firstCol' | 'firstRow' | 'lastCol' | 'lastRow' | 'neCell' | 'nwCell' | 'seCell' | 'swCell';
        pPr?: ParagraphFormatting;
        rPr?: TextFormatting;
        tblPr?: TableFormatting;
        trPr?: TableRowFormatting;
        tcPr?: TableCellFormatting;
    }>;
}
/**
 * Document defaults (w:docDefaults)
 */
interface DocDefaults {
    /** Default run properties */
    rPr?: TextFormatting;
    /** Default paragraph properties */
    pPr?: ParagraphFormatting;
}
/**
 * Style definitions from styles.xml
 */
interface StyleDefinitions {
    /** Document defaults */
    docDefaults?: DocDefaults;
    /** Latent styles */
    latentStyles?: {
        defLockedState?: boolean;
        defUIPriority?: number;
        defSemiHidden?: boolean;
        defUnhideWhenUsed?: boolean;
        defQFormat?: boolean;
        count?: number;
    };
    /** Style definitions */
    styles: Style[];
}
/**
 * Theme color scheme (a:clrScheme)
 */
interface ThemeColorScheme {
    /** Dark 1 color (usually black) */
    dk1?: string;
    /** Light 1 color (usually white) */
    lt1?: string;
    /** Dark 2 color */
    dk2?: string;
    /** Light 2 color */
    lt2?: string;
    /** Accent colors 1-6 */
    accent1?: string;
    accent2?: string;
    accent3?: string;
    accent4?: string;
    accent5?: string;
    accent6?: string;
    /** Hyperlink color */
    hlink?: string;
    /** Followed hyperlink color */
    folHlink?: string;
}
/**
 * Theme font (with script variants)
 */
interface ThemeFont {
    /** Latin font */
    latin?: string;
    /** East Asian font */
    ea?: string;
    /** Complex script font */
    cs?: string;
    /** Script-specific fonts */
    fonts?: Record<string, string>;
}
/**
 * Theme font scheme (a:fontScheme)
 */
interface ThemeFontScheme {
    /** Major font (headings) */
    majorFont?: ThemeFont;
    /** Minor font (body text) */
    minorFont?: ThemeFont;
}
/**
 * Theme (from theme1.xml)
 */
interface Theme {
    /** Theme name */
    name?: string;
    /** Color scheme */
    colorScheme?: ThemeColorScheme;
    /** Font scheme */
    fontScheme?: ThemeFontScheme;
    /** Format scheme (fills, lines, effects) - simplified */
    formatScheme?: {
        name?: string;
    };
}
/**
 * Font info from fontTable.xml
 */
interface FontInfo {
    /** Font name */
    name: string;
    /** Alternate names */
    altName?: string;
    /** Panose-1 classification */
    panose1?: string;
    /** Character set */
    charset?: string;
    /** Font family type */
    family?: 'decorative' | 'modern' | 'roman' | 'script' | 'swiss' | 'auto';
    /** Pitch (fixed or variable) */
    pitch?: 'default' | 'fixed' | 'variable';
    /** Signature */
    sig?: {
        usb0?: string;
        usb1?: string;
        usb2?: string;
        usb3?: string;
        csb0?: string;
        csb1?: string;
    };
    /** Embedded regular face (`w:embedRegular`) */
    embedRegular?: FontEmbed;
    /** Embedded bold face (`w:embedBold`) */
    embedBold?: FontEmbed;
    /** Embedded italic face (`w:embedItalic`) */
    embedItalic?: FontEmbed;
    /** Embedded bold-italic face (`w:embedBoldItalic`) */
    embedBoldItalic?: FontEmbed;
}
/**
 * A single embedded font face referenced from `fontTable.xml`. The font binary
 * lives at the relationship target (an obfuscated `.odttf` under `word/fonts/`)
 * and is descrambled with `fontKey`. See `utils/fontDeobfuscation.ts`.
 */
interface FontEmbed {
    /** Relationship ID (`r:id`) resolved via `word/_rels/fontTable.xml.rels`. */
    relId: string;
    /** Obfuscation key GUID (`w:fontKey`), e.g. `{XXXXXXXX-....}`. */
    fontKey?: string;
    /** Whether the embedded face is subsetted (`w:subsetted`). */
    subsetted?: boolean;
}
/**
 * Font table from fontTable.xml
 */
interface FontTable {
    fonts: FontInfo[];
}
/** OOXML relationship type URI (e.g. `.../relationships/image`). */
type RelationshipType = string;
/**
 * Relationship entry
 */
interface Relationship {
    /** Relationship ID (e.g., "rId1") */
    id: string;
    /** Relationship type URI */
    type: RelationshipType;
    /** Target path or URL */
    target: string;
    /** Target mode */
    targetMode?: 'External' | 'Internal';
}
/**
 * Relationship map (keyed by rId)
 */
type RelationshipMap = Map<string, Relationship>;
/**
 * Media file from word/media/
 */
interface MediaFile {
    /** File path in ZIP */
    path: string;
    /** Original filename */
    filename?: string;
    /** MIME type */
    mimeType: string;
    /** Binary data */
    data: ArrayBuffer;
    /** Base64 encoded data for rendering */
    base64?: string;
    /** Data URL for direct use in src attributes */
    dataUrl?: string;
}

export type { DocDefaults as D, FontTable as F, MediaFile as M, Relationship as R, StyleDefinitions as S, Theme as T, Style as a, RelationshipMap as b, FontEmbed as c, FontInfo as d, RelationshipType as e, StyleType as f, ThemeColorScheme as g, ThemeFont as h, ThemeFontScheme as i };
