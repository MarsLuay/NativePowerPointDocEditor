// Shared build-time patch for pptx-svg's PptxRenderer. Strips the default wasm
// URL (the bundle inlines bytes instead) and adds initJsBackend so the pure-JS
// engine fallback can stand in when WebAssembly GC is unavailable (older desktop
// installers, many mobile WebViews). Used by esbuild.config.mjs and the test
// bundler in tests/helpers/load-plugin-modules.mjs.

const DEFAULT_WASM_URL =
	"const DEFAULT_WASM_URL = new URL('./main.wasm', import.meta.url).href;";

// Expose a public single-slide entry point. The FFI feeds slide XML to Wasm from
// `this.files` on demand (get_file), and the renderer already has private
// persistFile()/reinitializeWasm() helpers its own structural ops use to apply
// file-level edits without a teardown. loadSlideXml() wires those together so the
// engine can do "patch one slide, re-render one slide" instead of exporting and
// reloading the whole deck. (Upstreamable as a public PptxRenderer method.)
//
// Reparse cost: `reinitializeWasm()` resets the whole model and re-reads every
// slide's XML on the next render — O(deck) per edit. pptx-svg >= 0.6.0 ships
// `restore_slide_ooxml(slideIdx, xml)`, which reparses *only* the edited slide
// (reusing the already-parsed layout/master/theme globals) and leaves the rest of
// the model untouched. Prefer it when present; fall back to the full reinit on
// older builds (our pinned 0.5.10) and the pure-JS backend, so this patch is a
// no-op there. `restore_slide_ooxml` returns "OK"/"ERROR:..."; a rejected reparse
// also falls back so a malformed edit can't leave the model half-updated.
const LOAD_SLIDE_XML_ANCHOR =
	"    persistFile(path, content) {";

const LOAD_SLIDE_XML_PATCH =
	"    loadSlideXml(slideIdx, xml) {\n" +
	"        this.persistFile(`ppt/slides/slide${slideIdx + 1}.xml`, xml);\n" +
	"        const reparseSlide = this.exports.restore_slide_ooxml;\n" +
	"        if (typeof reparseSlide === 'function') {\n" +
	"            const result = reparseSlide.call(this.exports, slideIdx, xml);\n" +
	"            if (typeof result !== 'string' || !result.startsWith('ERROR'))\n" +
	"                return;\n" +
	"        }\n" +
	"        this.reinitializeWasm();\n" +
	"    }\n" +
	"    persistFile(path, content) {";

const GETTER_ANCHOR = "    get exports() {\n        if (!this.wasm)";

const INIT_JS_BACKEND_PATCH =
	"    initJsBackend(engine) {\n" +
	"        this.__jsFfi = this.buildImportObject().pptx_ffi;\n" +
	"        this.wasm = { exports: engine };\n" +
	"        globalThis.pptx_ffi = this.__jsFfi;\n" +
	"    }\n" +
	"    get exports() {\n" +
	"        if (this.__jsFfi)\n" +
	"            globalThis.pptx_ffi = this.__jsFfi;\n" +
	"        if (!this.wasm)";

/**
 * @param {string} source Raw pptx-renderer.js source from pptx-svg.
 * @returns {string} Patched source ready for bundling.
 */
export function patchPptxRendererSource(source) {
	let contents = source.replace(DEFAULT_WASM_URL, "const DEFAULT_WASM_URL = undefined;");

	if (!contents.includes(GETTER_ANCHOR)) {
		throw new Error(
			"[patch-pptx-renderer] could not find the `exports` getter to patch — " +
				"pptx-svg internals changed; update scripts/lib/patch-pptx-renderer.mjs.",
		);
	}

	contents = contents.replace(GETTER_ANCHOR, INIT_JS_BACKEND_PATCH);

	if (!contents.includes(LOAD_SLIDE_XML_ANCHOR)) {
		throw new Error(
			"[patch-pptx-renderer] could not find persistFile() to anchor loadSlideXml — " +
				"pptx-svg internals changed; update scripts/lib/patch-pptx-renderer.mjs.",
		);
	}

	return contents.replace(LOAD_SLIDE_XML_ANCHOR, LOAD_SLIDE_XML_PATCH);
}
