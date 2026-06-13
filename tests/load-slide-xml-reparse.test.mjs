// Unit coverage for the loadSlideXml entry point that the build patch injects
// into pptx-svg's PptxRenderer (scripts/lib/patch-pptx-renderer.mjs).
//
// The fast path prefers the surgical per-slide reparse `restore_slide_ooxml`
// (pptx-svg >= 0.6.0) over the O(deck) `reinitializeWasm()`. Our pinned 0.5.10
// wasm doesn't export it, so the live renderer can only exercise the fallback —
// these tests drive the patched code directly against a minimal stub so both
// branches are guarded regardless of the installed pptx-svg version.

import { test } from "node:test";
import assert from "node:assert/strict";
import { patchPptxRendererSource } from "../scripts/lib/patch-pptx-renderer.mjs";

// Minimal source carrying the exact anchors patchPptxRendererSource keys off of:
// the DEFAULT_WASM_URL line, the `exports` getter, and `persistFile`.
const STUB_SOURCE = `
const DEFAULT_WASM_URL = new URL('./main.wasm', import.meta.url).href;
class PptxRenderer {
    constructor(exports) {
        this.wasm = { exports };
        this.calls = [];
    }
    buildImportObject() { return { pptx_ffi: {} }; }
    get exports() {
        if (!this.wasm) throw new Error('not initialized');
        return this.wasm.exports;
    }
    persistFile(path, content) {
        this.calls.push(['persistFile', path, content]);
    }
    reinitializeWasm() {
        this.calls.push(['reinitializeWasm']);
    }
}
globalThis.__PptxRenderer = PptxRenderer;
`;

function makeRendererClass() {
	const patched = patchPptxRendererSource(STUB_SOURCE);
	// Evaluate the patched module body and hand back the class it defines.
	// eslint-disable-next-line no-new-func
	new Function(patched)();
	const cls = globalThis.__PptxRenderer;
	delete globalThis.__PptxRenderer;
	return cls;
}

test("loadSlideXml: uses restore_slide_ooxml when present and skips full reinit", () => {
	const PptxRenderer = makeRendererClass();
	const reparseArgs = [];
	const renderer = new PptxRenderer({
		restore_slide_ooxml: (slideIdx, xml) => {
			reparseArgs.push([slideIdx, xml]);
			return "OK";
		},
	});

	renderer.loadSlideXml(2, "<p:sld/>");

	assert.deepEqual(reparseArgs, [[2, "<p:sld/>"]], "reparses only the edited slide");
	assert.deepEqual(
		renderer.calls,
		[["persistFile", "ppt/slides/slide3.xml", "<p:sld/>"]],
		"persists the slide XML and does NOT call reinitializeWasm"
	);
});

test("loadSlideXml: falls back to reinitializeWasm when restore_slide_ooxml is absent", () => {
	const PptxRenderer = makeRendererClass();
	const renderer = new PptxRenderer({}); // pinned 0.5.10 shape: no reparse export

	renderer.loadSlideXml(0, "<p:sld/>");

	assert.deepEqual(renderer.calls, [
		["persistFile", "ppt/slides/slide1.xml", "<p:sld/>"],
		["reinitializeWasm"],
	]);
});

test("loadSlideXml: falls back to reinitializeWasm when the per-slide reparse rejects", () => {
	const PptxRenderer = makeRendererClass();
	const renderer = new PptxRenderer({
		restore_slide_ooxml: () => "ERROR:bad xml",
	});

	renderer.loadSlideXml(4, "<broken/>");

	assert.deepEqual(
		renderer.calls,
		[
			["persistFile", "ppt/slides/slide5.xml", "<broken/>"],
			["reinitializeWasm"],
		],
		"a rejected reparse must not leave the model half-updated"
	);
});
