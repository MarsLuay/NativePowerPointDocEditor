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
import '../layout-engine/types.mjs';
export { B as BlockLookup, a as BlockLookupEntry, F as FRAGMENT_CLASS_NAMES, I as IMAGE_CLASS_NAMES, b as IMAGE_LAYOUT_OPTIONS, c as ImageHitTestResult, d as ImageLayoutIconHint, e as ImageLayoutOptionDef, L as LAYOUT_IMAGE_CLASSES, f as LayoutPainter, P as PainterOptions, RenderContext, RenderPageOptions, RenderPagesUpdateKind, T as TABLE_CLASS_NAMES, g as TEXTBOX_CLASS_NAMES, h as applySdtFocus, i as buildBlockLookup, j as captureInlinePositionEmu, k as createPainter, l as deriveLayoutChoice, m as enclosingSdtGroupIds, n as findImageElement, o as hitTestImage, isFloatingImageRun, p as isImageLayoutOptionEnabled, isTextWrappingFloatingImageRun, renderAllPagesNow, r as renderFragment, q as renderImageFragment, s as renderLine, renderPage, renderPages, t as renderParagraphFragment, u as renderTableFragment, v as renderTextBoxFragment, w as sliceRunsForLine, x as toolbarValueToLayoutTarget } from './renderPage.mjs';
export { P as PageGeometry, p as pageGeometryFromPage, r as resolveAnchoredObjectPosition, a as resolveAnchoredObjectVerticalTop } from '../anchoredObjectPosition-KJEyF-wr.mjs';
export { F as FootnoteRenderItem, H as HeaderFooterContent } from '../footnotes-BZ24OTAT.mjs';
import '../content-BZ9rYecc.mjs';
import '../formatting-DFtuRFQY.mjs';
import '../colors-C3vA7HUU.mjs';
import '../docx/wrapTypes.mjs';
import '../lists-CyGxd5Y2.mjs';
import '../watermark-D90356ZM.mjs';
import 'prosemirror-model';
import '../nodes-CHOFq5Ka.mjs';
import '../ImageExtension-BN327PNe.mjs';
import '../styles-Diw0MASy.mjs';
