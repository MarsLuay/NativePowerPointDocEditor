// Shared helpers for headless text-offset stamping harnesses (Electron / Chrome).

import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { spawn } from "node:child_process";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createRequire } from "node:module";
import { build } from "esbuild";
import { loadPresentationEngineModule } from "../../tests/helpers/load-plugin-modules.mjs";
import { toArrayBuffer } from "../../tests/helpers/renderer.mjs";

const require = createRequire(import.meta.url);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
export const projectRoot = path.resolve(__dirname, "../..");
export const electronMainPath = path.join(__dirname, "electron-eval-main.cjs");

export function resolveElectronBinary() {
  try {
    const binary = require("electron");
    if (typeof binary === "string" && existsSync(binary)) return binary;
  } catch {
    // not installed
  }
  return null;
}

export function findChrome() {
  const candidates = [
    process.env.CHROME_PATH,
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
  ].filter(Boolean);
  const chrome = candidates.find((candidate) => existsSync(candidate));
  if (!chrome) return null;
  return chrome;
}

export function runElectron(electronBinary, htmlFile, timeoutMs = 45000) {
  return new Promise((resolve, reject) => {
    const child = spawn(electronBinary, [electronMainPath], {
      stdio: ["ignore", "pipe", "pipe"],
      env: { ...process.env, HARNESS_HTML: htmlFile, ELECTRON_DISABLE_SECURITY_WARNINGS: "true" },
    });
    let stdout = "";
    let stderr = "";
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      reject(new Error(`Electron timed out after ${timeoutMs} ms: ${stderr || stdout}`));
    }, timeoutMs);

    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on("close", () => {
      clearTimeout(timer);
      const error = stdout.match(/^HARNESS_ERROR:(.*)$/m);
      if (error) {
        reject(new Error(`Electron harness error: ${error[1]}`));
        return;
      }
      const metrics = stdout.match(/^HARNESS_METRICS:(.*)$/m);
      if (!metrics) {
        reject(new Error(`Electron emitted no metrics. stderr: ${stderr.slice(-400)}`));
        return;
      }
      resolve(metrics[1]);
    });
  });
}

export async function bundleAnnotateModule() {
  const result = await build({
    entryPoints: [path.join(projectRoot, "src/powerpoint/annotateTextOffsets.ts")],
    bundle: true,
    format: "iife",
    globalName: "AnnotateTextOffsetsNS",
    platform: "browser",
    target: "es2020",
    write: false,
    logLevel: "silent",
    plugins: [{
      name: "stub-obsidian-annotate",
      setup(buildContext) {
        buildContext.onResolve({ filter: /^obsidian$/ }, () => ({ path: "obsidian", namespace: "stub-obsidian" }));
        buildContext.onLoad({ filter: /.*/, namespace: "stub-obsidian" }, () => ({
          contents: "export const Platform = { isMacOS: false };",
          loader: "js",
        }));
      },
    }],
  });
  return result.outputFiles[0].text;
}

export async function bundleTextUtilsModule() {
  const result = await build({
    entryPoints: [path.join(projectRoot, "src/powerpoint/textUtils.ts")],
    bundle: true,
    format: "iife",
    globalName: "TextUtilsNS",
    platform: "browser",
    target: "es2020",
    write: false,
    logLevel: "silent",
    plugins: [{
      name: "stub-obsidian-textutils",
      setup(buildContext) {
        buildContext.onResolve({ filter: /^obsidian$/ }, () => ({ path: "obsidian", namespace: "stub-obsidian" }));
        buildContext.onLoad({ filter: /.*/, namespace: "stub-obsidian" }, () => ({
          contents: "export const Platform = { isMacOS: false };",
          loader: "js",
        }));
      },
    }],
  });
  return result.outputFiles[0].text;
}

/** Build shape→paragraph→OOXML text map for a rendered slide. */
export async function buildOoxmlTextMap(engine, slideIndex) {
  const map = {};
  const svg = engine.renderSlide(slideIndex).svg;
  // Shape indices are discovered from the rendered SVG; paragraph indices are
  // probed until getParagraphRunText returns null (sparse paragraphs are fine).
  const shapeIdxPattern = /data-ooxml-shape-idx="(\d+)"/g;
  const shapeIndices = new Set();
  for (const match of svg.matchAll(shapeIdxPattern)) {
    shapeIndices.add(Number(match[1]));
  }

  for (const shapeIndex of shapeIndices) {
    map[shapeIndex] = {};
    for (let paragraphIndex = 0; paragraphIndex < 64; paragraphIndex++) {
      const text = engine.getParagraphRunText(slideIndex, shapeIndex, paragraphIndex);
      if (text === null) break;
      map[shapeIndex][paragraphIndex] = text;
    }
  }
  return map;
}

export async function loadDeckEngine(deckPath) {
  const resolved = path.resolve(process.cwd(), deckPath);
  assert.ok(existsSync(resolved), `deck not found: ${resolved}`);
  const bytes = await readFile(resolved);
  const { PresentationEngine } = await loadPresentationEngineModule();
  return PresentationEngine.load(toArrayBuffer(bytes));
}

export async function writeHarnessHtml({
  outputDir,
  filename,
  svg,
  ooxmlMap,
  annotateBundle,
  textUtilsBundle,
  driverScript,
}) {
  await mkdir(outputDir, { recursive: true });
  const htmlPath = path.join(outputDir, filename);
  const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Text offset stamping harness</title>
  <style>
    body { background: #f4f6fb; margin: 0; padding: 24px; font-family: Arial, sans-serif; }
    .native-powerpoint-canvas-pane { background: #d9dee9; display: inline-block; padding: 18px; }
    .native-powerpoint-slide-surface { background: white; }
    svg { display: block; }
  </style>
</head>
<body>
  <main>
    <div class="native-powerpoint-canvas-pane">
      <div class="native-powerpoint-slide-surface">${svg}</div>
    </div>
  </main>
  <script>window.OOXML_TEXT = ${JSON.stringify(ooxmlMap)};</script>
  <script>${textUtilsBundle}</script>
  <script>${annotateBundle}</script>
  <script>
    window.addEventListener('load', () => {
      try {
        ${driverScript}
      } catch (error) {
        document.body.dataset.metrics = encodeURIComponent(JSON.stringify({
          ok: false,
          error: String(error && error.stack || error),
        }));
      }
    });
  </script>
</body>
</html>`;
  await writeFile(htmlPath, html, "utf8");
  return htmlPath;
}

export async function runHeadlessHarness(htmlPath) {
  const electronBinary = resolveElectronBinary();
  if (electronBinary) {
    const encoded = await runElectron(electronBinary, htmlPath);
    return {
      runtime: `Electron ${require("electron/package.json").version}`,
      metrics: JSON.parse(decodeURIComponent(encoded)),
    };
  }

  const chromePath = findChrome();
  if (!chromePath) {
    return null;
  }

  const url = pathToFileURL(htmlPath).href;
  const userDataDir = path.join(os.tmpdir(), `native-powerpoint-offset-chrome-${process.pid}`);
  return new Promise((resolve, reject) => {
    const child = spawn(chromePath, [
      "--headless=new",
      "--disable-gpu",
      "--no-first-run",
      `--user-data-dir=${userDataDir}`,
      "--window-size=1700,1400",
      "--virtual-time-budget=1500",
      "--dump-dom",
      url,
    ], { stdio: ["ignore", "pipe", "pipe"] });

    let stdout = "";
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.on("close", (code) => {
      const match = stdout.match(/data-metrics="([^"]+)"/);
      if (!match) {
        reject(new Error(`Chrome harness emitted no metrics (exit ${code})`));
        return;
      }
      resolve({
        runtime: "headless Chrome",
        metrics: JSON.parse(decodeURIComponent(match[1])),
      });
    });
    child.on("error", reject);
  });
}
