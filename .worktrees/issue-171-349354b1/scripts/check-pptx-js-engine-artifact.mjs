import { stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const artifactPath = path.join(projectRoot, 'pptx-js-engine.mjs');
const syncStandardLimitBytes = 5_000_000;
const artifact = await import(pathToFileURL(artifactPath).href);
const size = (await stat(artifactPath)).size;

if (typeof artifact.createPptxJsEngine !== 'function') {
  throw new Error(`${artifactPath} does not export createPptxJsEngine().`);
}
if (size > syncStandardLimitBytes) {
  throw new Error(
    `${artifactPath} is ${size} bytes, above Obsidian Sync Standard's 5 MB per-file limit.`,
  );
}

console.log(`[check:pptx-js-engine-artifact] ${path.basename(artifactPath)} is ${size} bytes and exports the JS fallback.`);
