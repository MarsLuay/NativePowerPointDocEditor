# Performance issue: typing latency on large docs with comments + suggestions

Living doc. Updated as we investigate. Last updated: 2026-06-18.

## Summary

A large document (~127K words / ~309 pages) edits smoothly with **no** review
markup. Once the same document carries **comments and tracked changes
(suggestions)**, typing and undo/redo become visibly laggy.

The plain large-doc spec (`performance-large-docs.spec.ts`) passes every typing
scenario. The new comments+suggestions variant fails the same scenarios purely
on keystroke→repaint latency.

## Status

- [x] Reproduced with a deterministic fixture + e2e test
- [x] Root cause identified
- [x] Fix landed (React + core)
- [x] Vue checked — not affected (uses DOM-rect anchoring, not the layout walk)
- [x] Tests green

## Reproduction

Fixture: [`e2e/fixtures/issue-68-large-comments-suggestions.docx`](e2e/fixtures/issue-68-large-comments-suggestions.docx)

- Same body content/size as `issue-68-large.docx` (~127K words, 309 pages).
- **212 comments** (one every ~10 body paragraphs) with `comments.xml` +
  `commentsExtended.xml`.
- **211 tracked changes** — alternating `w:ins` insertions and `w:del`
  deletions, spread across the whole document.

Generator: [`scripts/generate-large-doc-comments-suggestions.ts`](scripts/generate-large-doc-comments-suggestions.ts)
(modeled on `scripts/generate-large-doc-issue68.ts`).

```bash
bun scripts/generate-large-doc-comments-suggestions.ts
```

Test: [`e2e/tests/performance-large-docs-comments-suggestions.spec.ts`](e2e/tests/performance-large-docs-comments-suggestions.spec.ts)
(mirrors `performance-large-docs.spec.ts` + one comment-adjacent typing test).

```bash
npx playwright test e2e/tests/performance-large-docs-comments-suggestions.spec.ts --timeout=120000 --workers=1
```

## Measured latency

Budget is `<500ms` per keystroke (same in both specs). Both specs are the same
scenarios on the same body; the only difference is the comments+suggestions
fixture. Numbers from back-to-back local runs on 2026-06-18 (Darwin arm64),
`--workers=1`.

Plain baseline: [`e2e/tests/performance-large-docs.spec.ts`](e2e/tests/performance-large-docs.spec.ts) (308 pages, no review markup).
Review variant: [`e2e/tests/performance-large-docs-comments-suggestions.spec.ts`](e2e/tests/performance-large-docs-comments-suggestions.spec.ts) (309 pages, 212 comments, 211 suggestions).

| Scenario                              | Plain (baseline) | Comments + suggestions (before fix) | Comments + suggestions (after fix) |
| ------------------------------------- | ---------------- | ----------------------------------- | ---------------------------------- |
| Load                                  | 2466ms           | 7159ms                              | 3018ms                             |
| Typing at document start              | 99ms ✅          | **3682ms** ❌                       | **285ms** ✅                       |
| Typing in the middle                  | 17ms ✅          | 17ms ✅                             | 17ms ✅                            |
| Typing near document end              | 17ms ✅          | 17ms ✅                             | 17ms ✅                            |
| Typing next to a comment + suggestion | n/a              | **3510ms** ❌                       | **281ms** ✅                       |
| Scrolling after edit                  | 1505ms ✅        | 1506ms ✅                           | 1506ms ✅                          |
| Undo                                  | 64ms ✅          | **5215ms** ❌                       | **140ms** ✅                       |
| Redo                                  | 81ms ✅          | **2126ms** ❌                       | **212ms** ✅                       |

✅ = under the 500ms budget (pass) · ❌ = over budget (fail). All scenarios pass after the fix.

### What the comparison tells us

- The slow scenarios are all **edit→re-render** paths. Steady-state typing in
  the **middle / end** of the doc is identical (17ms) in both — so the cost is
  not per-keystroke layout of the whole document.
- The blowup is concentrated where a transaction touches the **start of the
  document** or runs **undo/redo** — i.e. the paths that rebuild
  comment/suggestion-dependent state. Start-of-doc typing goes 99ms → 3682ms
  (~37×); undo goes 64ms → 5215ms (~82×).
- The comment-adjacent typing test (review variant only) confirms editing right
  next to review markup is in the same ~3.5s range.
- Load itself is ~2.9× slower, but that's a one-time cost, not the lag the user
  feels while editing.

## Root cause (confirmed)

The comments + tracked-changes sidebar resolves a vertical pixel position for
every comment/suggestion anchor on each full layout pass, via
`computeAnchorPositions` in
[`sidebarAnchorPositions.ts`](packages/react/src/components/DocxEditor/internals/sidebarAnchorPositions.ts).
For each of the ~423 anchors it called
[`getCaretPosition`](packages/core/src/layout-bridge/selectionRects.ts),
which **linearly scans every page × fragment** from page 0 to find the one
containing that PM position. That is O(anchors × pages) ≈ 423 × 309 per pass,
and it runs more than once per keystroke.

Why start-vs-end asymmetry: a full layout pass (and this anchor recompute) fires
when an edit shifts downstream pages — i.e. at the **start of the document** and
on **undo/redo**. Edits near the **end** take an incremental path that skips it,
which is why mid/end typing stayed at 17ms in both fixtures.

Measured with temporary instrumentation on the review fixture:
`computeAnchorPositions` took **1045–1593ms** with 423 anchors vs **0–1ms** with
7 anchors visible.

## Fix

`computeAnchorPositions` visits anchors in ascending PM order (via
`pmDoc.descendants`), and pages/fragments are laid out in that same order — so
the page containing each successive anchor never moves backwards. The fix threads
a monotonic page hint through the scan:

- `getCaretPosition` gained an optional `startPageIndex` (additive, backward
  compatible) so the scan can resume mid-document instead of restarting at 0.
- `computeAnchorPositions` tracks the last matched page and passes it as the
  hint (both the caret path and the table fallback).

This turns the pass from O(anchors × pages) into O(anchors + pages):
`computeAnchorPositions` dropped from ~1500ms to ~40ms with 423 anchors.

Vue is unaffected: its sidebar reads anchor positions straight from painted DOM
rects (`querySelectorAll('[data-comment-id]')` in `CommentMarginMarkers.vue`),
not the layout-engine walk — so there is no React→Vue mirror to do here.

## Possible follow-ups (not blocking)

- `computeAnchorPositions` still runs on every full layout pass; could be skipped
  when neither the doc nor the layout changed shape.
- Confirm the load-path improvement (7159ms → 3018ms) is also from this fix —
  the same anchor pass runs during initial settling.

## Notes

- Changeset: `.changeset/large-doc-comment-anchor-perf.md`.
