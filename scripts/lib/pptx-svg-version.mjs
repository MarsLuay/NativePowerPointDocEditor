import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const VENDOR_RELATIVE_PATH = 'src/vendor/pptx-js-engine.mjs';
const VENDOR_VERSION_PATTERN = /\* Source: pptx-svg v([0-9]+\.[0-9]+\.[0-9]+)/;

export function resolveProjectRoot(fromImportMetaUrl) {
  let dir = path.dirname(fileURLToPath(fromImportMetaUrl));

  while (dir !== path.dirname(dir)) {
    if (
      existsSync(path.join(dir, 'package.json')) &&
      existsSync(path.join(dir, 'manifest.json'))
    ) {
      return dir;
    }
    dir = path.dirname(dir);
  }

  throw new Error('Could not find plugin project root (package.json + manifest.json).');
}

export function readInstalledPptxSvgVersion(projectRoot) {
  const packagePath = path.join(projectRoot, 'node_modules/pptx-svg/package.json');
  if (!existsSync(packagePath)) {
    throw new Error('node_modules/pptx-svg not found. Run `npm install` first.');
  }

  const version = JSON.parse(readFileSync(packagePath, 'utf8')).version;
  if (!version) {
    throw new Error('node_modules/pptx-svg/package.json is missing a version field.');
  }

  return version;
}

export function readVendoredPptxJsEngineVersion(projectRoot) {
  const vendorPath = path.join(projectRoot, VENDOR_RELATIVE_PATH);
  if (!existsSync(vendorPath)) {
    throw new Error(`${VENDOR_RELATIVE_PATH} is missing. Run \`npm run regen:pptx-js\`.`);
  }

  const header = readFileSync(vendorPath, 'utf8').slice(0, 4096);
  const match = header.match(VENDOR_VERSION_PATTERN);
  return match?.[1] ?? null;
}

export function formatPptxJsEngineVersionMismatch({ installed, vendored }) {
  if (!vendored) {
    return [
      'The vendored pure-JS PPTX engine is missing a pptx-svg version stamp.',
      `Installed pptx-svg: ${installed}`,
      'Run `npm run regen:pptx-js` and commit src/vendor/pptx-js-engine.mjs.',
    ].join('\n');
  }

  return [
    'The vendored pure-JS PPTX engine does not match the installed pptx-svg package.',
    `Installed pptx-svg: ${installed}`,
    `Vendored engine:  ${vendored}`,
    'Run `npm run regen:pptx-js` and commit src/vendor/pptx-js-engine.mjs.',
  ].join('\n');
}
