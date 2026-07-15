/**
 * Layout Painter
 *
 * Main entry point for rendering Layout data to DOM.
 * Provides reconciliation for efficient incremental updates.
 *
 * @experimental Stable enough for the first-party React adapter, but the
 * API may change in minor releases until a third-party adapter validates
 * it. Pin a version range if you depend on this directly.
 * @packageDocumentation
 * @public
 */
import '../layout-engine/types.js';
export { B as BlockLookup, a as BlockLookupEntry, F as FRAGMENT_CLASS_NAMES, I as IMAGE_CLASS_NAMES, b as IMAGE_LAYOUT_OPTIONS, c as ImageHitTestResult, d as ImageLayoutIconHint, e as ImageLayoutOptionDef, L as LAYOUT_IMAGE_CLASSES, f as LayoutPainter, P as PainterOptions, RenderContext, RenderPageOptions, RenderPagesUpdateKind, T as TABLE_CLASS_NAMES, g as TEXTBOX_CLASS_NAMES, h as applySdtFocus, i as buildBlockLookup, j as captureInlinePositionEmu, k as createPainter, l as deriveLayoutChoice, m as enclosingSdtGroupIds, n as findImageElement, o as hitTestImage, isFloatingImageRun, p as isImageLayoutOptionEnabled, isTextWrappingFloatingImageRun, renderAllPagesNow, r as renderFragment, q as renderImageFragment, s as renderLine, renderPage, renderPages, t as renderParagraphFragment, u as renderTableFragment, v as renderTextBoxFragment, w as sliceRunsForLine, x as toolbarValueToLayoutTarget } from './renderPage.js';
export { P as PageGeometry, p as pageGeometryFromPage, r as resolveAnchoredObjectPosition, a as resolveAnchoredObjectVerticalTop } from '../anchoredObjectPosition-CS-8BfhO.js';
export { F as FootnoteRenderItem, H as HeaderFooterContent } from '../footnotes-DMsicPGd.js';
import '../content-B8ScSBzC.js';
import '../formatting-JhqWT_XM.js';
import '../colors-C3vA7HUU.js';
import '../docx/wrapTypes.js';
import '../lists-Bn29SzeS.js';
import '../watermark-D90356ZM.js';
import 'prosemirror-model';
import '../nodes-Bu9WifHU.js';
import '../ImageExtension-BlI5afZD.js';
import '../styles-2J4U-Lgk.js';
