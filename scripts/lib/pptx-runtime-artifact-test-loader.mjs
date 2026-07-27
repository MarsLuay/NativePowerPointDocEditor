import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { build } from "esbuild";
import { createInlinePptxSvgWasmPlugin } from "./patch-pptx-renderer.mjs";

const optionalPptxRuntimeArtifacts = [
  {
    artifact: "pptx-js-engine.mjs",
    source: "src/powerpoint/backend/pptxJsEngine.mjs",
    bundle: false,
  },
  {
    artifact: "pptx-wasm-renderer.mjs",
    source: "src/powerpoint/backend/pptxWasmRenderer.mjs",
    bundle: true,
  },
  {
    artifact: "heic-decode.mjs",
    source: "src/powerpoint/heicDecode.mjs",
    bundle: true,
  },
];

const inlinePptxSvgWasmPlugin = createInlinePptxSvgWasmPlugin();
const defaultProjectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

/**
 * Build the optional sidecars used by a temporary CommonJS engine bundle and
 * return the resolver it needs for dynamic imports. Production obtains these
 * URLs through Obsidian's vault adapter; smokes must supply file URLs instead.
 */
export async function createPptxRuntimeArtifactResolver(options) {
  const { projectRoot, outputDirectory } = typeof options === "string"
    ? { projectRoot: defaultProjectRoot, outputDirectory: options }
    : options;
  const artifactDirectory = path.join(outputDirectory, "pptx-runtime-artifacts");
  await Promise.all(optionalPptxRuntimeArtifacts.map(async ({ artifact, source, bundle }) => {
    await build({
      entryPoints: [path.join(projectRoot, source)],
      outfile: path.join(artifactDirectory, artifact),
      bundle,
      format: "esm",
      loader: { ".wasm": "binary" },
      logLevel: "silent",
      plugins: bundle ? [inlinePptxSvgWasmPlugin] : [],
      target: "es2020",
    });
  }));
  const artifactPaths = new Map(optionalPptxRuntimeArtifacts.map(({ artifact }) => [
    artifact,
    path.join(artifactDirectory, artifact),
  ]));
  return (artifact) => {
    const artifactPath = artifactPaths.get(artifact);
    if (!artifactPath) {
      throw new Error(`Missing test runtime artifact: ${artifact}`);
    }
    return {
      path: `test-runtime/${artifact}`,
      resourceUrl: pathToFileURL(artifactPath).href,
    };
  };
}
