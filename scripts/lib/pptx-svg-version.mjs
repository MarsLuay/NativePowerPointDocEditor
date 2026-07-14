import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ENGINE_RELATIVE_PATH = 'src/powerpoint/backend/pptxJsEngine.mjs';
const ENGINE_VERSION_PATTERN = /\* Source: pptx-svg v([0-9]+\.[0-9]+\.[0-9]+)/;

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

export function readPptxJsEngineVersion(projectRoot) {
  const enginePath = path.join(projectRoot, ENGINE_RELATIVE_PATH);
  if (!existsSync(enginePath)) {
    throw new Error(`${ENGINE_RELATIVE_PATH} is missing. Run \`npm run regen:pptx-js\`.`);
  }

  const header = readFileSync(enginePath, 'utf8').slice(0, 4096);
  const match = header.match(ENGINE_VERSION_PATTERN);
  return match?.[1] ?? null;
}


export function formatPptxJsEngineVersionMismatch({ installed, local, vendored }) {
  const engineVersion = local ?? vendored;
  if (!engineVersion) {
    return [
      'The pure-JS PPTX engine is missing a pptx-svg version stamp.',
      `Installed pptx-svg: ${installed}`,
      `Run \`npm run regen:pptx-js\` and commit ${ENGINE_RELATIVE_PATH}.`,
    ].join('\n');
  }

  return [
    'The pure-JS PPTX engine does not match the installed pptx-svg package.',
    `Installed pptx-svg: ${installed}`,
    `Local engine:      ${engineVersion}`,
    `Run \`npm run regen:pptx-js\` and commit ${ENGINE_RELATIVE_PATH}.`,
  ].join('\n');
}
