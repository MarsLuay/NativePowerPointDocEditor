import { debugLog, errorLog } from '../logger';

export const PPTX_RUNTIME_ARTIFACTS = [
  'pptx-js-engine.mjs',
  'pptx-wasm-renderer.mjs',
  'heic-decode.mjs',
] as const;

export type PptxRuntimeArtifact = typeof PPTX_RUNTIME_ARTIFACTS[number];

export interface PptxRuntimeArtifactResource {
  path: string;
  resourceUrl: string;
}

type PptxRuntimeArtifactResolver = (
  artifact: PptxRuntimeArtifact,
) => PptxRuntimeArtifactResource;

let artifactResources = new Map<PptxRuntimeArtifact, PptxRuntimeArtifactResource>();
const artifactModules = new Map<PptxRuntimeArtifact, Promise<unknown>>();

/**
 * Optional PPTX runtimes are gzip-embedded in main.js and materialized beside
 * it on load so community releases stay to Obsidian's supported assets
 * (main.js / manifest.json / styles.css) while each on-disk file stays under
 * Sync Standard's 5 MB limit. Their URLs must be resolved by Obsidian: a
 * relative dynamic import from CommonJS would resolve at app://obsidian.md/,
 * not this plugin directory.
 */
export function configurePptxRuntimeArtifactLoader(
  resolveArtifact: PptxRuntimeArtifactResolver,
): void {
  const resources = PPTX_RUNTIME_ARTIFACTS.map((artifact) => [artifact, resolveArtifact(artifact)] as const);
  for (const [artifact, resource] of resources) {
    if (!resource.path || !resource.resourceUrl) {
      throw new Error(`PowerPoint runtime artifact ${artifact} resolved to an empty path or URL.`);
    }
  }

  artifactResources = new Map(resources);
  artifactModules.clear();
  debugLog('render', 'Configured optional PowerPoint runtime artifact loader', {
    artifacts: resources.map(([artifact, resource]) => ({
      artifact,
      path: resource.path,
      resourceUrl: resource.resourceUrl,
    })),
  });
}

export function getPptxRuntimeArtifactResource(
  artifact: PptxRuntimeArtifact,
): PptxRuntimeArtifactResource {
  const resource = artifactResources.get(artifact);
  if (!resource) {
    throw new Error(
      `PowerPoint runtime artifact loader was not configured before loading ${artifact}.`,
    );
  }
  return resource;
}

export async function loadPptxRuntimeArtifact<T>(artifact: PptxRuntimeArtifact): Promise<T> {
  const existing = artifactModules.get(artifact);
  if (existing) {
    return await existing as T;
  }

  const { path, resourceUrl } = getPptxRuntimeArtifactResource(artifact);
  const startedAt = Date.now();
  debugLog('render', 'Loading optional PowerPoint runtime artifact', {
    artifact,
    path,
    resourceUrl,
  });

  // eslint-disable-next-line no-unsanitized/method -- resourceUrl is obtained from Obsidian's vault adapter for this plugin's own artifact.
  const load: Promise<unknown> = (import(resourceUrl) as Promise<unknown>)
    .then((module): unknown => {
      debugLog('render', 'Loaded optional PowerPoint runtime artifact', {
        artifact,
        path,
        resourceUrl,
        durationMs: Date.now() - startedAt,
      });
      return module;
    })
    .catch((error: unknown): never => {
      artifactModules.delete(artifact);
      errorLog('render', 'Failed to load optional PowerPoint runtime artifact', {
        artifact,
        path,
        resourceUrl,
        error,
      });
      throw error;
    });

  artifactModules.set(artifact, load);
  return await load as T;
}
