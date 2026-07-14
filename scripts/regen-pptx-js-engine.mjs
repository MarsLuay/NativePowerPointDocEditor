// Regenerate the pure-JS PPTX engine fallback (src/powerpoint/backend/pptxJsEngine.mjs)
// end to end: clone the matching pptx-svg source, patch its package config to
// emit a JS build, compile with MoonBit, and rewrap the output into the local
// factory module. Run this whenever the `pptx-svg` dependency is upgraded.
//
//   node scripts/regen-pptx-js-engine.mjs
//
// Environment overrides:
//   PPTX_SVG_REF=v0.5.10   Clone a specific git ref instead of the tag that
//                          matches the installed pptx-svg version.
//   INSTALL_MOONBIT=1      Auto-install the MoonBit toolchain if `moon` is
//                          missing (otherwise the script prints instructions).
//   KEEP_CLONE=1           Keep the temporary clone for inspection.
//
// Why this exists: the fallback only runs on old runtimes, the npm package ships
// only the wasm-gc binary (no JS build), and producing a JS build needs a small
// patch to the upstream `moon.pkg`. Without this script that knowledge lives only
// in a maintainer's head. See CONTRIBUTING.md › "PowerPoint engine backends".

import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir, homedir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, '..');
const REPO_URL = 'https://github.com/t-ujiie-g/pptx-svg.git';

function log(message) {
  console.log(`[regen] ${message}`);
}

function fail(message) {
  console.error(`[regen] ERROR: ${message}`);
  process.exit(1);
}

function run(command, args, options = {}) {
  return execFileSync(command, args, { stdio: 'inherit', ...options });
}

function tryRun(command, args, options = {}) {
  try {
    execFileSync(command, args, { stdio: 'ignore', ...options });
    return true;
  } catch {
    return false;
  }
}

// 1. Resolve the pptx-svg version to clone so the JS engine matches the wasm one.
const installedPkgPath = path.join(projectRoot, 'node_modules/pptx-svg/package.json');
if (!existsSync(installedPkgPath)) {
  fail('node_modules/pptx-svg not found. Run `npm install` first.');
}
const installedVersion = JSON.parse(readFileSync(installedPkgPath, 'utf8')).version;
const ref = process.env.PPTX_SVG_REF || `v${installedVersion}`;
log(`Installed pptx-svg is ${installedVersion}; cloning ref ${ref}.`);

// 2. Ensure the MoonBit toolchain is available.
const moonBin = resolveMoon();
const childEnv = {
  ...process.env,
  PATH: `${path.join(homedir(), '.moon/bin')}${path.delimiter}${process.env.PATH ?? ''}`
};

// 3. Clone the matching source into a temp dir.
const cloneDir = mkdtempSync(path.join(tmpdir(), 'pptx-svg-regen-'));
let cleanedUp = false;
const cleanup = () => {
  if (cleanedUp || process.env.KEEP_CLONE === '1') return;
  cleanedUp = true;
  rmSync(cloneDir, { recursive: true, force: true });
};
process.on('exit', cleanup);

try {
  log(`Cloning ${REPO_URL} @ ${ref} ...`);
  try {
    run('git', ['clone', '--depth', '1', '--branch', ref, REPO_URL, cloneDir]);
  } catch {
    fail(`Could not clone ref "${ref}". Set PPTX_SVG_REF to a valid tag/branch ` +
      `(see https://github.com/t-ujiie-g/pptx-svg/tags).`);
  }

  // 4. Patch src/main/moon.pkg to emit a JS build mirroring the wasm-gc exports.
  patchMoonPkg(path.join(cloneDir, 'src/main/moon.pkg'));

  // 5. Compile the JS backend.
  log('Building MoonBit JS target (moon build --target js --release) ...');
  run(moonBin, ['build', '--target', 'js', '--release'], { cwd: cloneDir, env: childEnv });

  const jsBuild = path.join(cloneDir, '_build/js/release/build/main/main.js');
  if (!existsSync(jsBuild)) {
    fail(`Expected MoonBit JS output at ${jsBuild} but it is missing.`);
  }

  // 6. Rewrap the raw MoonBit output into the per-instance factory module.
  log('Rewrapping output into src/powerpoint/backend/pptxJsEngine.mjs ...');
  run('node', [path.join(projectRoot, 'scripts/build-pptx-js-engine.mjs'), jsBuild], {
    cwd: projectRoot,
    env: {
      ...childEnv,
      PPTX_SVG_VERSION: installedVersion,
      PPTX_SVG_REF: ref,
    },
  });

  // 7. Verify the regenerated engine renders.
  log('Verifying with smoke:pptx-js ...');
  run('npm', ['run', 'smoke:pptx-js'], { cwd: projectRoot, env: childEnv });

  log(`Done. Regenerated from pptx-svg ${ref}.`);
} finally {
  cleanup();
}

function resolveMoon() {
  const candidates = ['moon', path.join(homedir(), '.moon/bin/moon')];
  for (const candidate of candidates) {
    if (tryRun(candidate, ['version'])) {
      log(`Using MoonBit at: ${candidate}`);
      return candidate;
    }
  }

  if (process.env.INSTALL_MOONBIT === '1') {
    log('MoonBit not found; installing (INSTALL_MOONBIT=1) ...');
    run('bash', ['-c', 'curl -fsSL https://cli.moonbitlang.com/install/unix.sh | bash']);
    const installed = path.join(homedir(), '.moon/bin/moon');
    if (tryRun(installed, ['version'])) return installed;
    fail('MoonBit install completed but `moon` still not runnable.');
  }

  fail(
    'MoonBit toolchain (`moon`) not found. Install it once with:\n' +
    '    curl -fsSL https://cli.moonbitlang.com/install/unix.sh | bash\n' +
    '  then re-run this script (or re-run with INSTALL_MOONBIT=1).'
  );
}

/**
 * Add a `"js"` link block to the package config, mirroring the `"wasm-gc"`
 * block's exports (the npm package only declares exports for wasm-gc, so the JS
 * backend otherwise links to an empty module). Idempotent and fails loudly if the
 * expected structure is missing.
 */
function patchMoonPkg(moonPkgPath) {
  if (!existsSync(moonPkgPath)) {
    fail(`Could not find ${moonPkgPath} in the clone (pptx-svg layout changed?).`);
  }

  const original = readFileSync(moonPkgPath, 'utf8');
  if (/"js"\s*:\s*\{/.test(original)) {
    log('moon.pkg already declares a "js" link block; leaving as-is.');
    return;
  }

  // Capture the whole wasm-gc block including its trailing "},", using the
  // block's own indentation as the closing anchor.
  const blockMatch = original.match(/( *)"wasm-gc"\s*:\s*\{[\s\S]*?\n\1\},/);
  if (!blockMatch) {
    fail('Could not locate the "wasm-gc" link block in moon.pkg to mirror.');
  }

  const wasmBlock = blockMatch[0];
  let jsBlock = wasmBlock.replace('"wasm-gc"', '"js"');
  if (/"use-js-builtin-string"\s*:\s*true\s*,?/.test(jsBlock)) {
    jsBlock = jsBlock.replace(/"use-js-builtin-string"\s*:\s*true\s*,?/, '"format": "esm",');
  } else {
    jsBlock = jsBlock.replace(/\n( *)\},$/, '\n$1  "format": "esm",\n$1},');
  }

  const patched = original.replace(wasmBlock, `${wasmBlock}\n${jsBlock}`);
  writeFileSync(moonPkgPath, patched, 'utf8');
  log('Patched moon.pkg with a mirrored "js" link block.');
}
