import { EditorState, Transaction } from 'prosemirror-state';
import { Node } from 'prosemirror-model';
import { PageMargins, ColumnLayout, FlowBlock, Measure, Layout } from '../layout-engine/types.mjs';
import { Document } from '../types/document.mjs';
import { S as SectionProperties, y as HeaderFooter } from '../content-BZ9rYecc.mjs';
import { F as FloatPageGeometry } from '../measureBlocksPipeline-BMqbRFQB.mjs';
import { H as HeaderFooterContent, F as FootnoteRenderItem } from '../footnotes-BZ24OTAT.mjs';
import { W as Watermark } from '../watermark-D90356ZM.mjs';
import { T as Theme, S as StyleDefinitions } from '../styles-Diw0MASy.mjs';
import '../formatting-DFtuRFQY.mjs';
import '../colors-C3vA7HUU.mjs';
import '../docx/wrapTypes.mjs';
import '../lists-CyGxd5Y2.mjs';
import '../anchoredObjectPosition-KJEyF-wr.mjs';

/**
 * The pure layout COMPUTE pass shared by the React and Vue adapters — issue
 * #696 Tier 2, the clean half of the engine spine.
 *
 * This is the 6-step pass from React's `useLayoutPipeline` minus the DOM paint
 * + scroll/event side-effects (which stay adapter-side, where the framework
 * timing lives): PM doc → flow blocks → measure → header/footer resolve →
 * margin extension → `layoutDocument` (+ two-pass footnote stabilization) →
 * footnote render items. It is pure (no DOM, no refs, no rAF) and returns
 * everything the adapter needs to paint.
 *
 * The one injected seam is `measureBlocks` — each adapter passes its own
 * measurer (React's is caching), same pattern as `measureBlocksWithFloats`.
 * `getHfPmDoc` is the HF-unification seam (prefer the persistent PM doc over
 * re-parsing `HeaderFooter.content`).
 */

interface PageSizePx {
    w: number;
    h: number;
}
/** Adapter-supplied block measurer (React's is caching). */
type MeasureBlocksFn = (blocks: FlowBlock[], contentWidth: number | number[], pageGeometry?: FloatPageGeometry) => Measure[];
interface ComputeLayoutInputs {
    state: EditorState;
    document: Document | null;
    pageSize: PageSizePx;
    margins: PageMargins;
    columns: ColumnLayout | undefined;
    finalPageSize: PageSizePx;
    finalMargins: PageMargins;
    finalColumns: ColumnLayout | undefined;
    pageGap: number;
    contentWidth: number;
    theme: Theme | null | undefined;
    styles: StyleDefinitions | null | undefined;
    sectionProperties: SectionProperties | null | undefined;
    finalSectionProperties: SectionProperties | null | undefined;
    /** Resolved HF objects for the section (default + first-page). */
    headerContent: HeaderFooter | null | undefined;
    footerContent: HeaderFooter | null | undefined;
    firstPageHeaderContent: HeaderFooter | null | undefined;
    firstPageFooterContent: HeaderFooter | null | undefined;
    measureBlocks: MeasureBlocksFn;
    /** HF unification: the persistent PM doc for an HF, or null to re-parse content. */
    getHfPmDoc: (hf: HeaderFooter) => Node | null | undefined;
}
interface LayoutComputation {
    blocks: FlowBlock[];
    measures: Measure[];
    layout: Layout;
    headerContentForRender: HeaderFooterContent | undefined;
    footerContentForRender: HeaderFooterContent | undefined;
    firstPageHeaderForRender: HeaderFooterContent | undefined;
    firstPageFooterForRender: HeaderFooterContent | undefined;
    hasTitlePg: boolean;
    watermark: Watermark | undefined;
    headerDistancePx: number | undefined;
    footerDistancePx: number | undefined;
    pageBorders: SectionProperties['pageBorders'] | undefined;
    footnotesByPage: Map<number, FootnoteRenderItem[]> | undefined;
}
/**
 * Run the pure layout compute pass (the 6 steps in this file's header), lifted
 * verbatim from `useLayoutPipeline`. The adapter performs the DOM paint
 * (`renderPages`), scroll-restore, `painter:painted`, and state writeback with
 * the returned values.
 */
declare function computeLayout(inputs: ComputeLayoutInputs): LayoutComputation;

/**
 * rAF-coalescing layout scheduler shared by the React and Vue adapters
 * (issue #696 Tier 2). Rapid doc-changing transactions (a burst of
 * keystrokes) collapse to a single layout pass per animation frame: while a
 * frame is pending, later `schedule` calls just replace the target state, so
 * only the final state lays out.
 *
 * `scheduleFrame`/`cancelFrame` are injected so a headless/test host can pass
 * a synchronous stub; they default to requestAnimationFrame.
 */

interface LayoutScheduler {
    /** Request a layout for `state`, coalesced into the pending frame. */
    schedule(state: EditorState): void;
    /** Cancel any pending frame (call on teardown). */
    cancel(): void;
}
declare function createLayoutScheduler(run: (state: EditorState) => void, scheduleFrame?: (cb: () => void) => number, cancelFrame?: (handle: number) => void): LayoutScheduler;

/**
 * Strip ProseMirror's internal `UPDATED_SCROLL` flag from a transaction so the
 * hidden, off-screen editor's `updateState` does not force-scroll its ancestors
 * to the selection — the paginated painter owns scroll, not the PM view.
 *
 * Shared by both adapters' `dispatchTransaction` (issue #696 Tier 2). React has
 * always done this; Vue did not, so its hidden editor could yank an ancestor's
 * scroll on edit. Centralizing the magic bit + the drift canary here keeps the
 * two in lockstep.
 *
 * `Transaction.updated` is a private bitfield in `prosemirror-state` whose
 * `UPDATED_SCROLL` bit is not exported. It is `4` in current PM
 * (state/src/transaction.ts). If a future PM release renumbers the bits before
 * SCROLL, `assertScrollFlagShape` logs once so the constant can be updated.
 */

/**
 * Clear the scroll-into-view flag on `transaction` in place. Call inside
 * `dispatchTransaction` before `state.apply(transaction)`. `probeTr` should be
 * a fresh `view.state.tr` used only for the one-shot drift canary (so the real
 * transaction is never mutated by the probe).
 */
declare function stripScrollFlag(transaction: Transaction, probeTr: Transaction): void;

export { type ComputeLayoutInputs, type LayoutComputation, type LayoutScheduler, type MeasureBlocksFn, computeLayout, createLayoutScheduler, stripScrollFlag };
