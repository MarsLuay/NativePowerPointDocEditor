/**
 * Optional PowerPoint / HEIC runtimes. Built as sibling .mjs files for local
 * vault deploy and tests, and gzip-embedded into main.js so Obsidian community
 * releases only need main.js / manifest.json / styles.css.
 */
export const PPTX_RUNTIME_ARTIFACT_SPECS = [
	{
		source: 'src/powerpoint/backend/pptxJsEngine.mjs',
		artifact: 'pptx-js-engine.mjs',
		bundle: false,
	},
	{
		source: 'src/powerpoint/backend/pptxWasmRenderer.mjs',
		artifact: 'pptx-wasm-renderer.mjs',
		bundle: true,
	},
	{
		source: 'src/powerpoint/heicDecode.mjs',
		artifact: 'heic-decode.mjs',
		bundle: true,
	},
];

/** Obsidian community plugin installer only downloads these release assets. */
export const OBSIDIAN_SUPPORTED_RELEASE_ASSETS = [
	'main.js',
	'manifest.json',
	'styles.css',
];
