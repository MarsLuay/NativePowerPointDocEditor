// Mobile PPTX smoke test (CI proxy for Obsidian Mobile WebViews).
//
// Obsidian Mobile (iOS WKWebView / Android System WebView) has no Node.js or
// Electron APIs. The build's check-mobile-compat.mjs guards against static
// require() of those modules, but that does not prove PPTX rendering works.
//
// Many mobile WebViews still lack WebAssembly GC. The plugin's JS engine
// fallback is the path those devices use. This script forces that fallback and
// loads a real deck through PresentationEngine — the same code path the mobile
// app runs when WasmGC is unavailable.
//
// Usage: node scripts/smoke-mobile-pptx.mjs [path-to.pptx]

import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadPresentationEngineModule } from '../tests/helpers/load-plugin-modules.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, '..');

const fixtureCandidates = [
	process.argv[2],
	path.join(projectRoot, 'tests/fixtures/decks/features.pptx'),
].filter(Boolean);

let pptxPath;
for (const candidate of fixtureCandidates) {
	if (existsSync(candidate)) {
		pptxPath = candidate;
		break;
	}
}

if (!pptxPath) {
	console.error(
		'No PPTX fixture found. Pass a path or generate fixtures with ' +
			'`node tests/fixtures/generate-fixtures.mjs`.',
	);
	process.exit(1);
}

function assert(cond, msg) {
	if (!cond) {
		console.error(`FAIL: ${msg}`);
		process.exit(1);
	}
}

// Force the JS backend — same flag PresentationEngine reads in the browser.
globalThis.__NATIVE_PPTX_FORCE_JS__ = true;

const fileBuf = await readFile(pptxPath);
const buffer = fileBuf.buffer.slice(fileBuf.byteOffset, fileBuf.byteOffset + fileBuf.byteLength);

const { PresentationEngine } = await loadPresentationEngineModule();
const engine = await PresentationEngine.load(buffer);

assert(engine.slideCount > 0, `expected slideCount > 0, got ${engine.slideCount}`);

const { svg } = engine.renderSlide(0);
assert(typeof svg === 'string' && svg.includes('<svg'), 'renderSlide(0) did not return SVG');
assert(!svg.startsWith('ERROR:'), `renderSlide returned ${svg.slice(0, 80)}`);

console.log(
	`PASS: mobile PPTX path (JS fallback via PresentationEngine) rendered slide 0 of ` +
		`${path.basename(pptxPath)} (slides=${engine.slideCount}, svg=${svg.length} bytes).`,
);
