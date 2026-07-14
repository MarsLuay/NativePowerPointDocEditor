/**
 * Document Content Model — barrel.
 *
 * All content-bearing types: runs, hyperlinks, bookmarks, fields,
 * images, shapes, tables, lists, paragraphs, headers/footers,
 * footnotes/endnotes, and sections.
 *
 * The types form a deeply interrelated tree (Paragraph ↔ Table ↔ ShapeTextBody)
 * and are split across `./content/*.ts` by domain. This file re-exports them
 * so existing imports from `@eigenpal/docx-editor-core/types/content` keep
 * working unchanged.
 * @packageDocumentation
 * @public
 */
export { B as BlockContent, v as BlockSdt, z as BookmarkEnd, A as BookmarkStart, x as BreakContent, G as Column, C as Comment, f as CommentRangeEnd, g as CommentRangeStart, J as ComplexField, h as Deletion, D as DocumentBody, K as DrawingContent, E as Endnote, L as EndnotePosition, N as EndnoteProperties, O as Field, Q as FieldCharContent, U as FieldType, V as FooterReference, F as Footnote, W as FootnotePosition, X as FootnoteProperties, y as HeaderFooter, Y as HeaderFooterType, Z as HeaderReference, H as Hyperlink, I as Image, _ as ImageCrop, $ as ImagePadding, a0 as ImagePosition, a1 as ImageSize, a2 as ImageTransform, a3 as ImageWrap, w as InlineSdt, i as Insertion, a4 as InstrTextContent, a5 as LineNumberRestart, a6 as MathEquation, M as MoveFrom, a7 as MoveFromRangeEnd, a8 as MoveFromRangeStart, j as MoveTo, a9 as MoveToRangeEnd, aa as MoveToRangeStart, ab as NoBreakHyphenContent, ac as NoteNumberRestart, ad as NoteRefMarkContent, ae as NoteReferenceContent, af as PageOrientation, P as Paragraph, k as ParagraphContent, b as ParagraphPropertyChange, ag as PropertyChangeInfo, l as Run, m as RunContent, R as RunPropertyChange, u as SdtDataBinding, t as SdtProperties, s as SdtType, ah as Section, S as SectionProperties, ai as SectionStart, aj as SeparatorContent, ak as Shape, al as ShapeContent, am as ShapeFill, an as ShapeOutline, ao as ShapeTextBody, ap as ShapeType, aq as SimpleField, ar as SoftHyphenContent, as as SymbolContent, at as TabContent, T as Table, n as TableCell, d as TableCellPropertyChange, c as TablePropertyChange, o as TableRow, e as TableRowPropertyChange, au as TableStructuralChangeInfo, av as TextBox, p as TextContent, q as TrackedChangeInfo, r as TrackedRunChange, aw as VerticalAlign } from '../content-B8ScSBzC.js';
export { D as DEFAULT_WATERMARK_PRESETS, P as PictureWatermark, T as TextWatermark, W as Watermark, p as pictureWatermarkDisplayEmu } from '../watermark-D90356ZM.js';
import '../formatting-JhqWT_XM.js';
import '../colors-C3vA7HUU.js';
import '../docx/wrapTypes.js';
import '../lists-Bn29SzeS.js';
