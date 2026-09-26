// Fail when src/powerpoint/backend/pptxJsEngine.mjs was not regenerated for the installed
// pptx-svg version.
import {
  formatPptxJsEngineVersionMismatch,
  readInstalledPptxSvgVersion,
  readPptxJsEngineVersion,
  resolveProjectRoot,
} from './lib/pptx-svg-version.mjs';

try {
  const projectRoot = resolveProjectRoot(import.meta.url);
  const installed = readInstalledPptxSvgVersion(projectRoot);
  const local = readPptxJsEngineVersion(projectRoot);

  if (!local || local !== installed) {
    console.error('[check:pptx-js-engine] Version mismatch.\n');
    console.error(formatPptxJsEngineVersionMismatch({ installed, local }));
    process.exit(1);
  }

  console.log(`[check:pptx-js-engine] Local JS engine matches pptx-svg v${installed}.`);
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  console.error(`[check:pptx-js-engine] ${message}`);
  process.exit(1);
}
