# Upstream `pptx-svg` patches

Items #8 (retain + render `<a:highlight>`) and #9 (a real per-slide reparse) both
live in the `pptx-svg` MoonBit/WASM engine, not in this plugin. This folder holds
what was drafted/verified against that source so the work can be landed upstream
(or in a fork) and the WASM rebuilt.

Source patched: `github.com/t-ujiie-g/pptx-svg` @ `0.6.0` (latest published).
Local validation used the MoonBit toolchain in `~/.moon`.

---

## #9 — per-slide reparse: DONE in-repo, no WASM change needed

**Finding:** the "real per-slide `loadSlideXml`/reparse export" already exists
upstream. `pptx-svg` ships `restore_slide_ooxml(slideIdx, xml)` (added for
undo/redo): it resolves the slide's rels/layout/master/theme from the
already-parsed globals, parses **only** that slide, and caches it — O(1 slide),
no whole-deck re-parse and no skeleton rebuild. Our pinned **0.5.10** simply
doesn't export it yet; **0.6.0** does.

`initialize_pptx()` never parses slide bodies anyway (slides are lazy
placeholders parsed on first render), so the only cost `reinitializeWasm()` was
paying per edit is re-reading every layout/master/theme/rels file. That's what
`restore_slide_ooxml` skips.

**What was changed here:** `scripts/lib/patch-pptx-renderer.mjs` rewires the
injected `loadSlideXml()` to prefer `restore_slide_ooxml` and fall back to the
full `reinitializeWasm()` when the export is absent (0.5.10, pure-JS backend) or
returns `ERROR:` (so a malformed edit can't leave the model half-updated).

```js
loadSlideXml(slideIdx, xml) {
    this.persistFile(`ppt/slides/slide${slideIdx + 1}.xml`, xml);
    const reparseSlide = this.exports.restore_slide_ooxml;
    if (typeof reparseSlide === 'function') {
        const result = reparseSlide.call(this.exports, slideIdx, xml);
        if (typeof result !== 'string' || !result.startsWith('ERROR')) return;
    }
    this.reinitializeWasm();
}
```

This is safe to ship now: on 0.5.10 it's a no-op (the fallback path is exactly
the previous behavior). Branch logic is covered by
`tests/load-slide-xml-reparse.test.mjs`; the full suite + build pass.

**To activate the fast path:** bump `pptx-svg` to `^0.6.0`
(`npm i pptx-svg@0.6.0`), rebuild, and re-run the suite. (Treat as its own change
— 0.5.10 → 0.6.0 is a renderer version jump that the fidelity tests should gate.)

---

## #8 — retain `<a:highlight>`: DRAFTED + type-checked upstream

`pptx-svg`'s model has **zero** highlight handling (confirmed: no `highlight`
reference anywhere in `src/`), which is the root cause of the data loss this
plugin papers over with the engine's `slideRunCache` reconciliation. Once the
engine round-trips `<a:highlight>` losslessly, that reconciliation (and the
"second source of truth") can be deleted.

**Patch:** `pptx-svg-0001-retain-highlight.patch`. It:

- adds `highlight : Color` to `TextRun`;
- parses `<a:rPr>/<a:highlight>` (`parse_highlight`, sharing `<a:solidFill>`'s
  inner color container) in `parse_run_props`;
- serializes it back in `serialize_run_props` (schema order: after the fill
  group, before `<a:latin>`);
- round-trips it through the SVG path too (`data-ooxml-highlight-color` emitted
  by `render_run_tspan`, read back by `parse_text_from_svg`);
- updates every `TextRun` literal (incl. test fixtures) and adds a
  parse → serialize round-trip test.

**Verification:** `moon check` passes (full project, incl. tests). `moon test`
**cannot run here** — the test wasm needs the JS host `pptx_ffi` imports
(`measure_text`, `log`, …), which is the "can't compile/test the Wasm here"
constraint. The added MoonBit round-trip test will execute in the upstream
harness that provides the FFI.

**Apply:**

```bash
cd /path/to/pptx-svg
git apply /path/to/patches/pptx-svg-0001-retain-highlight.patch
moon check && moon test   # test needs the FFI host harness
# rebuild WASM + publish/repackage, then bump this plugin's dependency
```

After this lands and the rebuilt WASM is in place, the engine's
`slideRunCache` highlight re-graft (and the `<a:highlight>` half of
`reconcileRunPropsIntoBuffer`) becomes redundant and can be removed (#1/#2).

---

## #8 — *render* `<a:highlight>`: remaining upstream work (design)

The patch above makes highlight **lossless** but does not yet **paint** it; the
plugin's overlay-rect subsystem still draws the visible highlight. Painting it in
the engine (which is what lets the overlay be deleted) is a larger renderer
change that needs WASM build + visual iteration, so it is intentionally **not**
in the type-checked patch.

Injection point: `src/renderer/renderer_text.mbt`. Runs are emitted as
`<tspan>`s inside a single `<text>` via natural flow (no per-run absolute x), so
a highlight can't be a per-tspan attribute. Approach:

1. In the wrap/layout pass (`wrap_paragraph` / the per-line loop in
   `render_text`), the renderer already tracks the line's x-cursor and uses
   `@ffi.ffi_measure_text` to advance it. For each run segment whose
   `run.highlight` is set, capture `(x, y_line_top, measured_width,
   line_height)` as it lays the run out.
2. Emit those as `<rect>` elements in a background layer **before** the `<text>`
   element (same group/transform), filled with `run.highlight.to_css()`. A
   batched per-(line,run) rect mirrors what the JS overlay computes today.
3. Verify visually against the existing render golden tests + a highlighted
   fixture, then delete the plugin-side overlay subsystem.

This is the only part of #8/#9 that genuinely requires compiling and visually
testing the WASM.
