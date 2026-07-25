/**
 * Paragraph Extension — paragraph node with alignment, spacing, indent, style commands
 *
 * Moves:
 * - NodeSpec from nodes.ts (paragraph, ParagraphAttrs, paragraphAttrsToDOMStyle, getListClass helpers)
 * - Commands from paragraph.ts (alignment, spacing, indent, style)
 */

import { Fragment, type NodeSpec, type Schema } from 'prosemirror-model';
import type { Command, EditorState } from 'prosemirror-state';
import type {
  ParagraphAlignment,
  LineSpacingRule,
  ParagraphFormatting,
  TextFormatting,
  NumberFormat,
  TabStop,
  TabStopAlignment,
  TabLeader,
} from '../../../types/document';
import type { BorderSpec, ColorValue } from '../../../types/colors';
import { paragraphToStyle } from '../../../utils/formatToStyle';
import { collectHeadings } from '../../../utils/headingCollector';
import { createNodeExtension } from '../create';
import type { ExtensionContext, ExtensionRuntime } from '../types';
import type { ParagraphAttrs } from '../../schema/nodes';
import {
  paragraphAttrsFromResolvedStyle,
  listAttrsFromResolvedStyle,
} from '../../styles/resolvedStyleAttrs';
import type { NumberingMap } from '../../../docx/numberingParser';
import {
  mergeParagraphAttrsWithOriginalFormatting,
  originalFormattingAfterApplyStyle,
} from '../../syncOriginalFormatting';
import {
  DOC_X_P_BLOCK_IMAGE_CLASS,
  createParagraphImageLayoutPlugin,
  getParagraphImageLayoutClasses,
} from './paragraphImageLayout';

// ============================================================================
// HELPERS (from nodes.ts)
// ============================================================================

function paragraphAttrsToDOMStyle(attrs: ParagraphAttrs): string {
  let indentLeft = attrs.indentLeft;
  if (attrs.numPr?.numId && indentLeft == null) {
    const level = attrs.numPr.ilvl ?? 0;
    indentLeft = (level + 1) * 720;
  }

  const formatting = {
    alignment: attrs.alignment,
    spaceBefore: attrs.spaceBefore,
    spaceAfter: attrs.spaceAfter,
    // HTML-origin auto spacing (w:beforeAutospacing/afterAutospacing) isn't a
    // tracked PM attr; it rides along on _originalFormatting. Forward it so
    // paragraphToStyle can render Word's ~14px auto spacing (issue #811).
    beforeAutospacing: attrs._originalFormatting?.beforeAutospacing,
    afterAutospacing: attrs._originalFormatting?.afterAutospacing,
    lineSpacing: attrs.lineSpacing,
    lineSpacingRule: attrs.lineSpacingRule,
    indentLeft: indentLeft,
    indentRight: attrs.indentRight,
    indentFirstLine: attrs.indentFirstLine,
    hangingIndent: attrs.hangingIndent,
    borders: attrs.borders,
    shading: attrs.shading,
  };

  const style = paragraphToStyle(formatting);
  if (style.marginTop) {
    style['--docx-space-before'] = style.marginTop;
  }
  if (style.marginBottom) {
    style['--docx-space-after'] = style.marginBottom;
  }
  return Object.entries(style)
    .map(([key, value]) => {
      const cssKey = key.replace(/([A-Z])/g, '-$1').toLowerCase();
      return `${cssKey}: ${value}`;
    })
    .join('; ');
}

function numFmtToClass(numFmt: NumberFormat | undefined): string {
  switch (numFmt) {
    case 'upperRoman':
      return 'docx-list-upper-roman';
    case 'lowerRoman':
      return 'docx-list-lower-roman';
    case 'upperLetter':
      return 'docx-list-upper-alpha';
    case 'lowerLetter':
      return 'docx-list-lower-alpha';
    case 'decimal':
    case 'decimalZero':
    default:
      return 'docx-list-decimal';
  }
}

function getListClass(
  numPr?: ParagraphAttrs['numPr'],
  listIsBullet?: boolean,
  listNumFmt?: NumberFormat
): string {
  if (!numPr?.numId) return '';

  const level = numPr.ilvl ?? 0;

  if (listIsBullet) {
    return `docx-list-bullet docx-list-level-${level}`;
  }

  const formatClass = numFmtToClass(listNumFmt);
  return `docx-list-numbered ${formatClass} docx-list-level-${level}`;
}

function parsePositiveInteger(value: string | undefined): number | undefined {
  if (!value || !/^\d+$/.test(value)) return undefined;
  const parsed = Number.parseInt(value, 10);
  return parsed > 0 ? parsed : undefined;
}

function parseListLevel(value: string | undefined): number | undefined {
  if (!value || !/^\d+$/.test(value)) return undefined;
  const parsed = Number.parseInt(value, 10);
  return parsed >= 0 && parsed <= 8 ? parsed : undefined;
}

function listFormatFromClassList(classList: DOMTokenList): NumberFormat | undefined {
  if (classList.contains('docx-list-upper-roman')) return 'upperRoman';
  if (classList.contains('docx-list-lower-roman')) return 'lowerRoman';
  if (classList.contains('docx-list-upper-alpha')) return 'upperLetter';
  if (classList.contains('docx-list-lower-alpha')) return 'lowerLetter';
  if (classList.contains('docx-list-decimal')) return 'decimal';
  return undefined;
}

function getExternalListLevel(listElement: HTMLElement): number {
  let level = 0;
  let ancestor = listElement.parentElement?.closest('ul, ol') as HTMLElement | null;
  while (ancestor) {
    level += 1;
    ancestor = ancestor.parentElement?.closest('ul, ol') as HTMLElement | null;
  }
  return level;
}

/**
 * Lists are paragraph attrs rather than ProseMirror list nodes. Preserve the
 * editor's list metadata on its HTML clipboard format and recognize native
 * HTML list items from other apps.
 */
function extractListAttrsFromElement(element: HTMLElement): Partial<ParagraphAttrs> {
  const classList = element.classList;
  const listItem = element.tagName.toLowerCase() === 'li'
    ? element
    : (element.closest('li') as HTMLElement | null);
  const htmlList = listItem?.parentElement?.closest('ul, ol') as HTMLElement | null;
  const listTag = htmlList?.tagName.toLowerCase();
  const internalBullet = classList.contains('docx-list-bullet');
  const internalNumbered = classList.contains('docx-list-numbered');
  const hasInternalListData = Boolean(element.dataset.listNumId) || internalBullet || internalNumbered;
  const hasExternalList = listTag === 'ul' || listTag === 'ol';
  if (!hasInternalListData && !hasExternalList) return {};

  const isBullet = element.dataset.listIsBullet === 'true' || internalBullet || listTag === 'ul';
  const levelClass = [...classList].find((name) => /^docx-list-level-\d+$/.test(name));
  const levelFromClass = levelClass ? parseListLevel(levelClass.replace('docx-list-level-', '')) : undefined;
  const level = parseListLevel(element.dataset.listLevel)
    ?? levelFromClass
    ?? (htmlList ? getExternalListLevel(htmlList) : 0);
  const numId = parsePositiveInteger(element.dataset.listNumId) ?? (isBullet ? 1 : 2);
  const listNumFmt = isBullet ? 'bullet' : (listFormatFromClassList(classList) ?? 'decimal');
  const listMarker = element.dataset.listMarker || (isBullet ? '•' : undefined);

  return {
    numPr: { numId, ilvl: level },
    listIsBullet: isBullet,
    listNumFmt,
    listMarker,
  };
}

function parseClipboardJsonObject<T>(value: string | undefined): T | undefined {
  if (!value) return undefined;
  try {
    const parsed: unknown = JSON.parse(value);
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed)
      ? parsed as T
      : undefined;
  } catch {
    return undefined;
  }
}

function parseDefaultTextFormatting(value: string | undefined): TextFormatting | undefined {
  return parseClipboardJsonObject<TextFormatting>(value);
}

function cssBorderStyleToOoxml(cssStyle: string): BorderSpec['style'] {
  switch (cssStyle.toLowerCase()) {
    case 'solid':
      return 'single';
    case 'double':
      return 'double';
    case 'dotted':
      return 'dotted';
    case 'dashed':
      return 'dashed';
    case 'groove':
      return 'threeDEngrave';
    case 'ridge':
      return 'threeDEmboss';
    case 'inset':
      return 'inset';
    case 'outset':
      return 'outset';
    default:
      return 'single';
  }
}

function cssBorderWidthToEighths(cssWidth: string): number {
  const trimmed = cssWidth.trim().toLowerCase();
  if (trimmed === 'thin') return 4;
  if (trimmed === 'medium') return 8;
  if (trimmed === 'thick') return 16;
  const value = Number.parseFloat(trimmed);
  if (Number.isNaN(value)) return 8;
  if (trimmed.endsWith('pt')) return Math.round(value * 8);
  return Math.round(value * 6); // CSS pixels or unitless widths
}

function parseCssBorderColor(value: string): ColorValue | undefined {
  const hex = value.match(/#([0-9a-f]{6})/i);
  if (hex) return { rgb: hex[1].toUpperCase() };
  const shortHex = value.match(/#([0-9a-f]{3})$/i);
  if (shortHex) {
    const [red, green, blue] = shortHex[1];
    return { rgb: `${red}${red}${green}${green}${blue}${blue}`.toUpperCase() };
  }
  const rgb = value.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
  if (!rgb) return undefined;
  return {
    rgb: [rgb[1], rgb[2], rgb[3]]
      .map((component) => Number.parseInt(component, 10).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase(),
  };
}

function extractParagraphBordersFromStyle(style: CSSStyleDeclaration): ParagraphAttrs['borders'] | undefined {
  const parseSide = (cssStyle: string, cssColor: string, cssWidth: string): BorderSpec | undefined => {
    if (!cssStyle || cssStyle === 'none' || cssStyle === 'hidden') return undefined;
    return {
      style: cssBorderStyleToOoxml(cssStyle),
      color: parseCssBorderColor(cssColor),
      size: cssBorderWidthToEighths(cssWidth),
    };
  };
  const top = parseSide(style.borderTopStyle, style.borderTopColor, style.borderTopWidth);
  const bottom = parseSide(style.borderBottomStyle, style.borderBottomColor, style.borderBottomWidth);
  const left = parseSide(style.borderLeftStyle, style.borderLeftColor, style.borderLeftWidth);
  const right = parseSide(style.borderRightStyle, style.borderRightColor, style.borderRightWidth);
  return top || bottom || left || right ? { top, bottom, left, right } : undefined;
}

function extractParagraphBorders(element: HTMLElement): ParagraphAttrs['borders'] | undefined {
  return parseClipboardJsonObject<NonNullable<ParagraphAttrs['borders']>>(element.dataset.paragraphBorders)
    ?? extractParagraphBordersFromStyle(element.style);
}

function extractEmptyParagraphTextFormatting(element: HTMLElement): TextFormatting | undefined {
  const copiedFormatting = parseDefaultTextFormatting(element.dataset.defaultTextFormatting);
  if (copiedFormatting) return copiedFormatting;
  if (element.textContent?.trim()) return undefined;

  const fontSize = element.style.fontSize.trim();
  const ptMatch = fontSize.match(/^([\d.]+)pt$/);
  const pxMatch = fontSize.match(/^([\d.]+)px$/);
  const halfPoints = ptMatch
    ? Math.round(Number.parseFloat(ptMatch[1]) * 2)
    : (pxMatch ? Math.round(Number.parseFloat(pxMatch[1]) * 1.5) : undefined);
  const fontFamily = element.style.fontFamily.trim();
  if (halfPoints == null && !fontFamily) return undefined;

  return {
    ...(halfPoints != null ? { fontSize: halfPoints, fontSizeCs: halfPoints } : {}),
    ...(fontFamily ? { fontFamily: { ascii: fontFamily, hAnsi: fontFamily } } : {}),
  };
}

// ============================================================================
// CSS-TO-TWIPS HELPERS (for paste from external apps like Google Docs)
// ============================================================================

/**
 * Parse a CSS length value to twips.
 * Supports pt, px, in, cm, mm units. Returns undefined for unparseable values.
 *
 * Conversion factors (1 inch = 1440 twips):
 * - 1pt = 20 twips (1440/72)
 * - 1px = 15 twips (1440/96)
 * - 1cm = 567 twips (1440/2.54, rounded)
 * - 1mm = 56.7 twips (1440/25.4)
 */
function cssLengthToTwips(value: string): number | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  const num = parseFloat(trimmed);
  if (isNaN(num) || num === 0) return undefined;

  if (trimmed.endsWith('pt')) return Math.round(num * 20);
  if (trimmed.endsWith('px')) return Math.round(num * 15);
  if (trimmed.endsWith('in')) return Math.round(num * 1440);
  if (trimmed.endsWith('mm')) return Math.round(num * (1440 / 25.4));
  if (trimmed.endsWith('cm')) return Math.round(num * (1440 / 2.54));
  // Bare number — treat as pixels (browser computed style default)
  if (/^[\d.]+$/.test(trimmed)) return Math.round(num * 15);
  return undefined;
}

/**
 * Map CSS text-align value to ParagraphAlignment.
 */
function cssTextAlignToAlignment(value: string): ParagraphAlignment | undefined {
  switch (value.trim().toLowerCase()) {
    case 'left':
    case 'start':
      return 'left';
    case 'center':
      return 'center';
    case 'right':
    case 'end':
      return 'right';
    case 'justify':
      return 'both';
    default:
      return undefined;
  }
}

/**
 * Parse CSS line-height to twips.
 * - Unitless multiplier (e.g. "1.5"): 240 twips * multiplier (single=240)
 * - Percentage (e.g. "150%"): 240 twips * (pct/100)
 * - Absolute length (e.g. "18pt"): converted to twips directly with 'exact' rule
 *
 * Returns { lineSpacing, lineSpacingRule } or undefined.
 */
function cssLineHeightToSpacing(
  value: string
): { lineSpacing: number; lineSpacingRule: LineSpacingRule } | undefined {
  if (!value || value === 'normal') return undefined;
  const trimmed = value.trim();

  // Percentage (e.g. "150%")
  if (trimmed.endsWith('%')) {
    const pct = parseFloat(trimmed);
    if (isNaN(pct) || pct === 0) return undefined;
    return { lineSpacing: Math.round(240 * (pct / 100)), lineSpacingRule: 'auto' };
  }

  // Absolute length (has a unit like pt, px, etc.)
  if (/[a-z]/i.test(trimmed)) {
    const twips = cssLengthToTwips(trimmed);
    if (twips == null) return undefined;
    return { lineSpacing: twips, lineSpacingRule: 'exact' };
  }

  // Unitless multiplier (e.g. "1.5", "2")
  const multiplier = parseFloat(trimmed);
  if (isNaN(multiplier) || multiplier === 0) return undefined;
  return { lineSpacing: Math.round(240 * multiplier), lineSpacingRule: 'auto' };
}

/**
 * Extract paragraph-level attributes from a pasted HTML <p> element's inline styles.
 * Used by parseDOM to preserve formatting from external apps (Google Docs, Word Online, etc.).
 */
function extractParagraphAttrsFromStyle(element: HTMLElement): Partial<ParagraphAttrs> {
  const style = element.style;
  const attrs: Partial<ParagraphAttrs> = {};

  // Alignment — text-align CSS property
  if (style.textAlign) {
    const alignment = cssTextAlignToAlignment(style.textAlign);
    if (alignment) attrs.alignment = alignment;
  }

  // Left indentation — margin-left or padding-left (Google Docs uses margin-left)
  const marginLeft = style.marginLeft || style.paddingLeft;
  if (marginLeft) {
    const twips = cssLengthToTwips(marginLeft);
    if (twips != null) attrs.indentLeft = twips;
  }

  // Right indentation — margin-right or padding-right
  const marginRight = style.marginRight || style.paddingRight;
  if (marginRight) {
    const twips = cssLengthToTwips(marginRight);
    if (twips != null) attrs.indentRight = twips;
  }

  // First-line indent — text-indent CSS property
  if (style.textIndent) {
    const twips = cssLengthToTwips(style.textIndent);
    if (twips != null) {
      if (twips < 0) {
        // Negative text-indent means hanging indent
        attrs.indentFirstLine = Math.abs(twips);
        attrs.hangingIndent = true;
      } else {
        attrs.indentFirstLine = twips;
      }
    }
  }

  // Line spacing — line-height CSS property
  if (style.lineHeight) {
    const spacing = cssLineHeightToSpacing(style.lineHeight);
    if (spacing) {
      attrs.lineSpacing = spacing.lineSpacing;
      attrs.lineSpacingRule = spacing.lineSpacingRule;
    }
  }

  // Space before/after — margin-top/margin-bottom
  if (style.marginTop) {
    const twips = cssLengthToTwips(style.marginTop);
    if (twips != null) attrs.spaceBefore = twips;
  }
  if (style.marginBottom) {
    const twips = cssLengthToTwips(style.marginBottom);
    if (twips != null) attrs.spaceAfter = twips;
  }

  const borders = extractParagraphBordersFromStyle(style);
  if (borders) {
    attrs.borders = borders;
  }

  return attrs;
}

// ============================================================================
// PARAGRAPH NODE SPEC
// ============================================================================

const paragraphNodeSpec: NodeSpec = {
  content: 'inline*',
  group: 'block',
  attrs: {
    paraId: { default: null },
    textId: { default: null },
    alignment: { default: null },
    spaceBefore: { default: null },
    spaceAfter: { default: null },
    lineSpacing: { default: null },
    lineSpacingRule: { default: null },
    spacingExplicit: { default: null },
    indentLeft: { default: null },
    indentRight: { default: null },
    indentFirstLine: { default: null },
    hangingIndent: { default: false },
    numPr: { default: null },
    numPrFromStyle: { default: null },
    listNumFmt: { default: null },
    listIsBullet: { default: null },
    listMarker: { default: null },
    listMarkerHidden: { default: null },
    listMarkerFontFamily: { default: null },
    listMarkerFontSize: { default: null },
    listMarkerSuffix: { default: null },
    listLevelNumFmts: { default: null },
    listAbstractNumId: { default: null },
    listStartOverride: { default: null },
    styleId: { default: null },
    borders: { default: null },
    shading: { default: null },
    tabs: { default: null },
    pageBreakBefore: { default: null },
    // `<w:lastRenderedPageBreak/>` — Word's cached layout marker.
    renderedPageBreakBefore: { default: null },
    keepNext: { default: null },
    keepLines: { default: null },
    contextualSpacing: { default: null },
    defaultTextFormatting: { default: null },
    sectionBreakType: { default: null },
    bidi: { default: null },
    outlineLevel: { default: null },
    bookmarks: { default: null },
    _originalFormatting: { default: null },
    _originalRunBoundaries: { default: null },
    _sectionProperties: { default: null },
    // Tracked structural revisions on the paragraph mark itself.
    // See ECMA-376 §17.13.5 — w:ins / w:del inside w:pPr/w:rPr.
    pPrIns: { default: null },
    pPrDel: { default: null },
    // Paragraph property changes (w:pPrChange), array of prior-snapshot
    // entries. Parser/serializer already model this — schema attr lets
    // the data survive the toProseDoc → fromProseDoc round-trip.
    pPrChange: { default: null },
  },
  parseDOM: [
    {
      tag: 'p',
      getAttrs(dom): ParagraphAttrs {
        const element = dom as HTMLElement;
        const listAttrs = extractListAttrsFromElement(element);
        const borders = extractParagraphBorders(element);

        // Start with data-attribute values (from our own editor's copy/paste)
        const attrs: ParagraphAttrs = {
          paraId: element.dataset.paraId || undefined,
          alignment: element.dataset.alignment as ParagraphAlignment | undefined,
          styleId: element.dataset.styleId || undefined,
          sectionBreakType:
            (element.dataset.sectionBreak as ParagraphAttrs['sectionBreakType']) || undefined,
          ...listAttrs,
          borders,
          defaultTextFormatting: extractEmptyParagraphTextFormatting(element),
        };

        // Extract paragraph formatting from inline CSS styles
        // (covers paste from Google Docs, Word Online, and other external apps)
        const styleAttrs = extractParagraphAttrsFromStyle(element);

        // Merge: data-attributes take precedence over CSS-extracted values
        return {
          ...styleAttrs,
          ...attrs,
          // For alignment, prefer data-attribute if present, otherwise use CSS
          alignment: attrs.alignment || styleAttrs.alignment || undefined,
        };
      },
    },
    {
      // HTML clipboard content from browsers, Word, and Google Docs commonly
      // represents bullets as `<ul><li>…</li></ul>`. Flatten each item into
      // our paragraph-with-list-attrs model rather than dropping the marker.
      tag: 'li',
      getAttrs(dom): ParagraphAttrs {
        const element = dom as HTMLElement;
        const styleAttrs = extractParagraphAttrsFromStyle(element);
        return {
          ...styleAttrs,
          ...extractListAttrsFromElement(element),
          borders: extractParagraphBorders(element),
          defaultTextFormatting: extractEmptyParagraphTextFormatting(element),
        };
      },
    },
    // Heading tags (h1-h6) — pasted from Google Docs, Word Online, etc.
    // Map to paragraphs with appropriate styleId and formatting extracted from CSS.
    ...(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const).map((tag) => ({
      tag,
      getAttrs(dom: HTMLElement): ParagraphAttrs {
        const level = parseInt(tag.charAt(1));
        const styleAttrs = extractParagraphAttrsFromStyle(dom);

        return {
          ...styleAttrs,
          styleId: `Heading${level}`,
          outlineLevel: level - 1,
        };
      },
    })),
  ],
  toDOM(node) {
    const attrs = node.attrs as ParagraphAttrs;
    let style = paragraphAttrsToDOMStyle(attrs);
    const listClass = getListClass(attrs.numPr, attrs.listIsBullet, attrs.listNumFmt);
    const imageLayoutClasses = getParagraphImageLayoutClasses(node);

    // Block-only image paragraphs: zero first-line indent in the style attribute
    // (CSS text-indent is only partially supported on the Obsidian baseline).
    if (imageLayoutClasses.includes(DOC_X_P_BLOCK_IMAGE_CLASS)) {
      style = style
        .split(';')
        .map((part) => part.trim())
        .filter((part) => part && !/^text-indent\s*:/i.test(part))
        .concat(['text-indent: 0'])
        .join('; ');
    }

    const domAttrs: Record<string, string> = {};

    if (style) {
      domAttrs.style = style;
    }

    if (listClass) {
      domAttrs.class = listClass;
    }

    if (attrs.paraId) {
      domAttrs['data-para-id'] = attrs.paraId;
    }

    if (attrs.alignment) {
      domAttrs['data-alignment'] = attrs.alignment;
    }

    if (attrs.styleId) {
      domAttrs['data-style-id'] = attrs.styleId;
    }

    if (attrs.listMarker) {
      domAttrs['data-list-marker'] = attrs.listMarker;
    }

    if (attrs.numPr?.numId) {
      domAttrs['data-list-num-id'] = String(attrs.numPr.numId);
      domAttrs['data-list-level'] = String(attrs.numPr.ilvl ?? 0);
      domAttrs['data-list-is-bullet'] = String(Boolean(attrs.listIsBullet));
    }

    if (node.content.size === 0 && attrs.defaultTextFormatting) {
      domAttrs['data-default-text-formatting'] = JSON.stringify(attrs.defaultTextFormatting);
    }

    if (attrs.borders) {
      domAttrs['data-paragraph-borders'] = JSON.stringify(attrs.borders);
    }

    if (attrs.bidi) {
      domAttrs.dir = 'rtl';
    }

    if (attrs.sectionBreakType) {
      domAttrs['data-section-break'] = attrs.sectionBreakType;
      domAttrs.class = (domAttrs.class ? domAttrs.class + ' ' : '') + 'docx-section-break';
    }

    if (attrs.pPrIns || attrs.pPrDel) {
      const rev = attrs.pPrIns ?? attrs.pPrDel!;
      const kindClass = attrs.pPrIns ? 'ep-revision-ins' : 'ep-revision-del';
      domAttrs.class =
        (domAttrs.class ? domAttrs.class + ' ' : '') + 'ep-revision-pmark ' + kindClass;
      domAttrs['data-revision-id'] = String(rev.revisionId);
      domAttrs['data-revision-author'] = rev.author;
      if (rev.date) domAttrs['data-revision-date'] = rev.date;
    } else if (Array.isArray(attrs.pPrChange) && attrs.pPrChange.length > 0) {
      // Property-change-only paragraph (no structural pPrIns/pPrDel).
      // Surface the first entry's id so sidebar click-to-jump can anchor.
      const first = attrs.pPrChange[0];
      domAttrs.class = (domAttrs.class ? domAttrs.class + ' ' : '') + 'ep-revision-prop-change';
      domAttrs['data-revision-id'] = String(first.info.id);
      domAttrs['data-revision-author'] = first.info.author;
      if (first.info.date) domAttrs['data-revision-date'] = first.info.date;
    }

    if (imageLayoutClasses.length > 0) {
      domAttrs.class =
        (domAttrs.class ? domAttrs.class + ' ' : '') + imageLayoutClasses.join(' ');
    }

    return ['p', domAttrs, 0];
  },
};

// ============================================================================
// PARAGRAPH COMMAND HELPERS
// ============================================================================

function setParagraphAttr(attr: string, value: unknown): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    if (!dispatch) return true;

    let tr = state.tr;
    const seen = new Set<number>();

    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === 'paragraph' && !seen.has(pos)) {
        seen.add(pos);
        tr = tr.setNodeMarkup(
          pos,
          undefined,
          mergeParagraphAttrsWithOriginalFormatting(node.attrs, { [attr]: value })
        );
      }
    });

    dispatch(tr.scrollIntoView());
    return true;
  };
}

function setParagraphAttrsCmd(attrs: Record<string, unknown>): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    if (!dispatch) return true;

    let tr = state.tr;
    const seen = new Set<number>();

    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === 'paragraph' && !seen.has(pos)) {
        seen.add(pos);
        tr = tr.setNodeMarkup(
          pos,
          undefined,
          mergeParagraphAttrsWithOriginalFormatting(node.attrs, attrs)
        );
      }
    });

    dispatch(tr.scrollIntoView());
    return true;
  };
}

function makeSetParagraphBottomBorder(border: BorderSpec): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;
    if (!dispatch) return true;

    let tr = state.tr;
    const seen = new Set<number>();
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name !== 'paragraph' || seen.has(pos)) return;
      seen.add(pos);
      const currentBorders = (node.attrs.borders as ParagraphAttrs['borders'] | null | undefined) ?? {};
      tr = tr.setNodeMarkup(
        pos,
        undefined,
        mergeParagraphAttrsWithOriginalFormatting(node.attrs, {
          borders: { ...currentBorders, bottom: { ...border } },
        })
      );
    });

    dispatch(tr.scrollIntoView());
    return true;
  };
}

// ============================================================================
// RESOLVED STYLE ATTRS (for applyStyle)
// ============================================================================

export interface ResolvedStyleAttrs {
  paragraphFormatting?: ParagraphFormatting;
  runFormatting?: TextFormatting;
  /**
   * Numbering definitions from the document package. When the applied style
   * carries a `w:numPr`, these resolve the numbering level into the list
   * marker attrs (template, per-level formats, counter key) so the painter
   * renders the style's numbering — e.g. "[Claim 1]" — instead of falling
   * back to a plain decimal marker.
   */
  numbering?: NumberingMap | null;
}

// ============================================================================
// COMMAND FACTORIES
// ============================================================================

function makeSetAlignment(alignment: ParagraphAlignment): Command {
  return (state, dispatch) => {
    return setParagraphAttr('alignment', alignment)(state, dispatch);
  };
}

function makeSetLineSpacing(value: number, rule: LineSpacingRule = 'auto'): Command {
  return (state, dispatch) => {
    return setParagraphAttrsCmd({
      lineSpacing: value,
      lineSpacingRule: rule,
    })(state, dispatch);
  };
}

function makeIncreaseIndent(amount: number = 720): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    if (!dispatch) return true;

    let tr = state.tr;
    const seen = new Set<number>();

    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === 'paragraph' && !seen.has(pos)) {
        seen.add(pos);
        const currentIndent = node.attrs.indentLeft || 0;
        tr = tr.setNodeMarkup(
          pos,
          undefined,
          mergeParagraphAttrsWithOriginalFormatting(node.attrs, {
            indentLeft: currentIndent + amount,
          })
        );
      }
    });

    dispatch(tr.scrollIntoView());
    return true;
  };
}

function makeDecreaseIndent(amount: number = 720): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    if (!dispatch) return true;

    let tr = state.tr;
    const seen = new Set<number>();

    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === 'paragraph' && !seen.has(pos)) {
        seen.add(pos);
        const currentIndent = node.attrs.indentLeft || 0;
        const newIndent = Math.max(0, currentIndent - amount);
        tr = tr.setNodeMarkup(
          pos,
          undefined,
          mergeParagraphAttrsWithOriginalFormatting(node.attrs, {
            indentLeft: newIndent > 0 ? newIndent : null,
          })
        );
      }
    });

    dispatch(tr.scrollIntoView());
    return true;
  };
}

function makeApplyStyle(schema: Schema) {
  return (styleId: string, resolvedAttrs?: ResolvedStyleAttrs): Command => {
    return (state, dispatch) => {
      const { $from, $to } = state.selection;

      if (!dispatch) return true;

      let tr = state.tr;
      const seen = new Set<number>();

      // Build marks from run formatting if provided
      const styleMarks: import('prosemirror-model').Mark[] = [];
      if (resolvedAttrs?.runFormatting) {
        const rpr = resolvedAttrs.runFormatting;

        if (rpr.bold) {
          styleMarks.push(schema.marks.bold.create());
        }
        if (rpr.italic) {
          styleMarks.push(schema.marks.italic.create());
        }
        if (rpr.fontSize || rpr.fontSizeCs) {
          styleMarks.push(
            schema.marks.fontSize.create({
              size: rpr.fontSize ?? null,
              sizeCs: rpr.fontSizeCs ?? rpr.fontSize ?? null,
            })
          );
        }
        if (rpr.fontFamily) {
          styleMarks.push(
            schema.marks.fontFamily.create({
              ascii: rpr.fontFamily.ascii,
              hAnsi: rpr.fontFamily.hAnsi,
              asciiTheme: rpr.fontFamily.asciiTheme,
            })
          );
        }
        if (rpr.color && !rpr.color.auto) {
          styleMarks.push(
            schema.marks.textColor.create({
              rgb: rpr.color.rgb,
              themeColor: rpr.color.themeColor,
              themeTint: rpr.color.themeTint,
              themeShade: rpr.color.themeShade,
            })
          );
        }
        if (rpr.underline && rpr.underline.style !== 'none') {
          styleMarks.push(
            schema.marks.underline.create({
              style: rpr.underline.style,
              color: rpr.underline.color,
            })
          );
        }
        if (rpr.strike || rpr.doubleStrike) {
          styleMarks.push(schema.marks.strike.create({ double: rpr.doubleStrike || false }));
        }
      }

      // Mark types that are controlled by style definitions
      const styleControlledMarks = [
        schema.marks.bold,
        schema.marks.italic,
        schema.marks.fontSize,
        schema.marks.fontFamily,
        schema.marks.textColor,
        schema.marks.underline,
        schema.marks.strike,
      ].filter(Boolean);

      state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
        if (node.type.name === 'paragraph' && !seen.has(pos)) {
          seen.add(pos);

          const newAttrs: Record<string, unknown> = {
            ...node.attrs,
            styleId,
          };

          if (resolvedAttrs) {
            // When applying a style, explicitly reset all style-controlled
            // paragraph attrs to the new style's values (or null to clear).
            // This prevents old style properties (e.g. heading line spacing)
            // from persisting when switching to a different style. The same
            // projection drives the Enter handler's next-style switch.
            Object.assign(newAttrs, paragraphAttrsFromResolvedStyle(resolvedAttrs));
            // A style with `w:numPr` attaches its numbering (numPr + marker
            // attrs). A style without numbering leaves existing list attrs
            // untouched — direct numbering survives a style switch in Word.
            const listAttrs = listAttrsFromResolvedStyle(resolvedAttrs, resolvedAttrs.numbering);
            if (listAttrs) {
              Object.assign(newAttrs, listAttrs);
            }
            // Drop stale direct pPr/rPr from load provenance so save does not
            // resurrect the previous style's spacing/indent/run props.
            newAttrs._originalFormatting = originalFormattingAfterApplyStyle(
              node.attrs._originalFormatting as ParagraphFormatting | null | undefined,
              styleId
            );
          } else if (node.attrs._originalFormatting) {
            newAttrs._originalFormatting = {
              ...(node.attrs._originalFormatting as ParagraphFormatting),
              styleId,
            };
          }

          tr = tr.setNodeMarkup(pos, undefined, newAttrs);

          // Only modify marks when we have resolved style attrs
          // (fallback path without resolvedAttrs just sets styleId)
          if (resolvedAttrs) {
            const paragraphStart = pos + 1;
            const paragraphEnd = pos + node.nodeSize - 1;

            if (paragraphEnd > paragraphStart) {
              // Clear old style-controlled marks first
              for (const markType of styleControlledMarks) {
                tr = tr.removeMark(paragraphStart, paragraphEnd, markType);
              }
              // Then add the new style's marks
              for (const mark of styleMarks) {
                tr = tr.addMark(paragraphStart, paragraphEnd, mark);
              }
            }
          }
        }
      });

      if (styleMarks.length > 0) {
        tr = tr.setStoredMarks(styleMarks);
      }

      dispatch(tr.scrollIntoView());
      return true;
    };
  };
}

// ============================================================================
// QUERY HELPERS (exported for toolbar)
// ============================================================================

export function getParagraphAlignment(state: EditorState): ParagraphAlignment | null {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return null;
  return paragraph.attrs.alignment || null;
}

export function getParagraphTabs(state: EditorState): TabStop[] | null {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return null;
  return paragraph.attrs.tabs || null;
}

export function getStyleId(state: EditorState): string | null {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return null;
  return paragraph.attrs.styleId || null;
}

export function getParagraphBidi(state: EditorState): boolean {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return false;
  return !!paragraph.attrs.bidi;
}

// ============================================================================
// EXTENSION
// ============================================================================

export const ParagraphExtension = createNodeExtension({
  name: 'paragraph',
  schemaNodeName: 'paragraph',
  nodeSpec: paragraphNodeSpec,
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const applyStyleFn = makeApplyStyle(ctx.schema);

    return {
      plugins: [createParagraphImageLayoutPlugin()],
      commands: {
        setAlignment: (alignment: ParagraphAlignment) => makeSetAlignment(alignment),
        alignLeft: () => makeSetAlignment('left'),
        alignCenter: () => makeSetAlignment('center'),
        alignRight: () => makeSetAlignment('right'),
        alignJustify: () => makeSetAlignment('both'),
        setLineSpacing: (value: number, rule?: LineSpacingRule) => makeSetLineSpacing(value, rule),
        singleSpacing: () => makeSetLineSpacing(240),
        oneAndHalfSpacing: () => makeSetLineSpacing(360),
        doubleSpacing: () => makeSetLineSpacing(480),
        setSpaceBefore: (twips: number) => setParagraphAttr('spaceBefore', twips),
        setSpaceAfter: (twips: number) => setParagraphAttr('spaceAfter', twips),
        setParagraphBottomBorder: (border?: BorderSpec) =>
          makeSetParagraphBottomBorder(border ?? {
            style: 'single',
            size: 6,
            space: 1,
            color: { rgb: '000000' },
          }),
        clearParagraphBorders: () => setParagraphAttr('borders', null),
        increaseIndent: (amount?: number) => makeIncreaseIndent(amount),
        decreaseIndent: (amount?: number) => makeDecreaseIndent(amount),
        setIndentLeft: (twips: number) => setParagraphAttr('indentLeft', twips > 0 ? twips : null),
        setIndentRight: (twips: number) =>
          setParagraphAttr('indentRight', twips > 0 ? twips : null),
        setIndentFirstLine: (twips: number, hanging?: boolean) =>
          setParagraphAttrsCmd({
            indentFirstLine: twips > 0 ? twips : null,
            hangingIndent: hanging ?? false,
          }),
        applyStyle: (styleId: string, resolvedAttrs?: ResolvedStyleAttrs) =>
          applyStyleFn(styleId, resolvedAttrs),
        clearStyle: () => setParagraphAttr('styleId', null),
        insertSectionBreak: (breakType: 'nextPage' | 'continuous' | 'oddPage' | 'evenPage') =>
          setParagraphAttr('sectionBreakType', breakType),
        removeSectionBreak: () => setParagraphAttr('sectionBreakType', null),
        generateTOC: () => {
          return (
            state: EditorState,
            dispatch?: (tr: import('prosemirror-state').Transaction) => void
          ) => {
            const headings = collectHeadings(state.doc);
            if (headings.length === 0) return false;
            if (!dispatch) return true;

            const { schema: s } = state;
            const tr = state.tr;

            // Generate unique bookmark names for each heading and set them on heading paragraphs
            const bookmarkEntries: Array<{ name: string; level: number; text: string }> = [];
            for (const h of headings) {
              const bookmarkName = `_Toc${Math.floor(100000000 + Math.random() * 900000000)}`;
              bookmarkEntries.push({ name: bookmarkName, level: h.level, text: h.text });

              // Map position through prior transaction steps, then resolve against current tr.doc
              const mappedPos = tr.mapping.map(h.pmPos);
              const $pos = tr.doc.resolve(mappedPos);
              const paragraphNode = $pos.nodeAfter;
              if (paragraphNode && paragraphNode.type.name === 'paragraph') {
                // Filter out any existing _Toc bookmarks to avoid duplicates on regeneration
                const existingBookmarks =
                  (paragraphNode.attrs.bookmarks as Array<{ id: number; name: string }>) || [];
                const filteredBookmarks = existingBookmarks.filter(
                  (b) => !b.name.startsWith('_Toc')
                );
                const newBookmarks = [
                  ...filteredBookmarks,
                  { id: Math.floor(Math.random() * 2147483647), name: bookmarkName },
                ];
                tr.setNodeMarkup(mappedPos, undefined, {
                  ...paragraphNode.attrs,
                  bookmarks: newBookmarks,
                });
              }
            }

            // Build TOC paragraphs
            const tocNodes: import('prosemirror-model').Node[] = [];

            // TOC title
            tocNodes.push(
              s.node('paragraph', { styleId: 'TOCHeading', alignment: 'center' }, [
                s.text('Table of Contents', [s.marks.bold.create()]),
              ])
            );

            // TOC entries with hyperlinks
            for (const entry of bookmarkEntries) {
              const indent = entry.level * 720; // 0.5 inch per level in twips
              const tocStyleId = `TOC${entry.level + 1}`; // TOC1, TOC2, etc.
              const linkMark = s.marks.hyperlink.create({ href: `#${entry.name}` });

              tocNodes.push(
                s.node(
                  'paragraph',
                  {
                    styleId: tocStyleId,
                    indentLeft: indent > 0 ? indent : null,
                  },
                  [s.text(entry.text, [linkMark])]
                )
              );
            }

            // Insert TOC at cursor position — use a Fragment for correct ordering
            const insertPos = tr.mapping.map(state.selection.from);
            tr.insert(insertPos, Fragment.from(tocNodes));
            dispatch(tr.scrollIntoView());
            return true;
          };
        },
        toggleBidi: () => {
          return (
            state: EditorState,
            dispatch?: (tr: import('prosemirror-state').Transaction) => void
          ) => {
            const { $from } = state.selection;
            const paragraph = $from.parent;
            if (paragraph.type.name !== 'paragraph') return false;
            const currentBidi = paragraph.attrs.bidi || false;
            return setParagraphAttr('bidi', currentBidi ? null : true)(state, dispatch);
          };
        },
        setRtl: () => setParagraphAttr('bidi', true),
        setLtr: () => setParagraphAttr('bidi', null),
        setTabs: (tabs: TabStop[]) => setParagraphAttr('tabs', tabs.length > 0 ? tabs : null),
        addTabStop: (
          position: number,
          alignment: TabStopAlignment = 'left',
          leader: TabLeader = 'none'
        ) => {
          return (
            state: EditorState,
            dispatch?: (tr: import('prosemirror-state').Transaction) => void
          ) => {
            const { $from } = state.selection;
            const paragraph = $from.parent;
            if (paragraph.type.name !== 'paragraph') return false;
            const currentTabs: TabStop[] = paragraph.attrs.tabs || [];
            const filtered = currentTabs.filter((t: TabStop) => t.position !== position);
            const newTabs = [...filtered, { position, alignment, leader }].sort(
              (a: TabStop, b: TabStop) => a.position - b.position
            );
            return setParagraphAttr('tabs', newTabs)(state, dispatch);
          };
        },
        removeTabStop: (position: number) => {
          return (
            state: EditorState,
            dispatch?: (tr: import('prosemirror-state').Transaction) => void
          ) => {
            const { $from } = state.selection;
            const paragraph = $from.parent;
            if (paragraph.type.name !== 'paragraph') return false;
            const currentTabs: TabStop[] = paragraph.attrs.tabs || [];
            const newTabs = currentTabs.filter((t: TabStop) => t.position !== position);
            return setParagraphAttr('tabs', newTabs.length > 0 ? newTabs : null)(state, dispatch);
          };
        },
      },
    };
  },
});
