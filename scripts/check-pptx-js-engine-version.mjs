// Fail when src/vendor/pptx-js-engine.mjs was not regenerated for the installed
// pptx-svg version. Catches silent drift on older mobile runtimes that rely on
// the pure-JS fallback instead of the wasm-gc binary.

import {
  formatPptxJsEngineVersionMismatch,
  readInstalledPptxSvgVersion,
  readVendoredPptxJsEngineVersion,
  resolveProjectRoot,
} from './lib/pptx-svg-version.mjs';

const projectRoot = resolveProjectRoot(import.meta.url);

try {
  const installed = readInstalledPptxSvgVersion(projectRoot);
  const vendored = readVendoredPptxJsEngineVersion(projectRoot);

  if (!vendored || vendored !== installed) {
    console.error('[check:pptx-js-engine] Version mismatch.\n');
    console.error(formatPptxJsEngineVersionMismatch({ installed, vendored }));
    process.exit(1);
  }

  console.log(`[check:pptx-js-engine] Vendored JS engine matches pptx-svg v${installed}.`);
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  console.error(`[check:pptx-js-engine] ${message}`);
  process.exit(1);
}
