import { performance } from "node:perf_hooks";
import { createDeck } from "../../tests/helpers/fixture-builder.mjs";
import { loadPresentationEngineModule } from "../../tests/helpers/load-plugin-modules.mjs";

/** Representative deck sizes used by the real thumbnail benchmark command. */
export const DEFAULT_PPTX_THUMBNAIL_TIERS = [50, 100, 250, 500];
export const THUMBNAIL_PRIORITY_RADIUS = 2;

function round(value) {
  return Math.round(value * 10) / 10;
}

function percentile(values, percentileValue) {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.min(sorted.length - 1, Math.floor((sorted.length - 1) * percentileValue))] ?? 0;
}

function summarize(values) {
  return {
    count: values.length,
    p50Ms: round(percentile(values, 0.5)),
    p95Ms: round(percentile(values, 0.95)),
    maxMs: round(Math.max(0, ...values)),
  };
}

function priorityThumbnailIndices(currentSlide, slideCount, radius = THUMBNAIL_PRIORITY_RADIUS) {
  const start = Math.max(0, currentSlide - radius);
  const end = Math.min(slideCount - 1, currentSlide + radius);
  const indices = [];
  for (let index = start; index <= end; index += 1) indices.push(index);
  return indices;
}

function remainingThumbnailIndices(slideCount, rendered) {
  const indices = [];
  for (let index = 0; index < slideCount; index += 1) {
    if (!rendered.has(index)) indices.push(index);
  }
  return indices;
}

function nextImmediate() {
  return new Promise((resolve) => setImmediate(resolve));
}

function startEventLoopProbe() {
  let stopped = false;
  let last = performance.now();
  const gaps = [];
  const tick = () => {
    const now = performance.now();
    gaps.push(now - last);
    last = now;
    if (!stopped) setImmediate(tick);
  };
  setImmediate(tick);
  return {
    async stop() {
      stopped = true;
      await nextImmediate();
      return gaps;
    },
  };
}

/**
 * Exercise the same PresentationEngine.renderSlide path used by
 * SlideFilmstripController for non-active thumbnails. The benchmark cache is
 * deliberately explicit so cache hits/misses and cold render timings cannot
 * be confused with DOM traversal measurements.
 */
export async function runPptxThumbnailBenchmarkTier(slideCount, options = {}) {
  if (!Number.isInteger(slideCount) || slideCount < 1) {
    throw new Error("slideCount must be a positive integer");
  }

  const { PresentationEngine } = await loadPresentationEngineModule();
  const pptxBuffer = createDeck({ format: "pptx", slideCount, richFirstSlide: true });
  const documentSizeBytes = pptxBuffer.byteLength;
  const engine = await PresentationEngine.load(
    pptxBuffer.buffer.slice(pptxBuffer.byteOffset, pptxBuffer.byteOffset + pptxBuffer.byteLength),
  );
  const currentSlide = Math.min(
    slideCount - 1,
    Math.max(0, Number.isInteger(options.currentSlide) ? options.currentSlide : Math.floor(slideCount / 2)),
  );
  const priorityIndices = priorityThumbnailIndices(currentSlide, slideCount);
  const nearbyIndices = priorityIndices.filter((index) => index !== currentSlide);
  const backgroundIndices = remainingThumbnailIndices(slideCount, new Set(priorityIndices));
  const cache = new Map();
  const measurements = [];
  let cacheHits = 0;
  let cacheMisses = 0;
  let renderedSvgCount = 0;

  const renderThumbnail = (index, role) => {
    const cached = cache.get(index);
    if (cached) {
      cacheHits += 1;
      measurements.push({ index, role, cacheHit: true, ms: 0 });
      return cached;
    }

    cacheMisses += 1;
    const started = performance.now();
    // This is the production slide renderer, not a mock DOM traversal.
    const rendered = engine.renderSlide(index);
    if (!rendered?.svg || !rendered.svg.startsWith("<svg")) {
      throw new Error(`Slide ${index} did not produce an SVG thumbnail render`);
    }
    const result = { svg: rendered.svg };
    cache.set(index, result);
    renderedSvgCount += 1;
    measurements.push({ index, role, cacheHit: false, ms: round(performance.now() - started) });
    return result;
  };

  const activeStarted = performance.now();
  renderThumbnail(currentSlide, "active");
  const activeSlideMs = round(performance.now() - activeStarted);

  const nearbyStarted = performance.now();
  for (const index of nearbyIndices) renderThumbnail(index, "nearby");
  const nearbyCompletionMs = round(performance.now() - nearbyStarted);

  const eventLoopProbe = startEventLoopProbe();
  const backgroundStarted = performance.now();
  for (const index of backgroundIndices) {
    renderThumbnail(index, "background");
    // Keep each actual render in its own task so the probe reports interaction
    // gaps rather than only one aggregate loop duration.
    await nextImmediate();
  }
  const backgroundCompletionMs = round(performance.now() - backgroundStarted);
  const eventLoopGaps = await eventLoopProbe.stop();

  // Prove the cache boundary without adding cold renders to the throughput
  // sample: active and one useful neighbor should be cache hits.
  renderThumbnail(currentSlide, "cache-probe");
  if (nearbyIndices[0] !== undefined) renderThumbnail(nearbyIndices[0], "cache-probe");

  const coldMeasurements = measurements.filter((measurement) => !measurement.cacheHit);
  const coldDurations = coldMeasurements.map((measurement) => measurement.ms);
  const nearbyDurations = coldMeasurements
    .filter((measurement) => measurement.role === "nearby")
    .map((measurement) => measurement.ms);
  const backgroundDurations = coldMeasurements
    .filter((measurement) => measurement.role === "background")
    .map((measurement) => measurement.ms);

  return {
    schemaVersion: 1,
    tier: slideCount,
    unit: "slides",
    fixture: "deterministic-pptx-thumbnails-v1",
    documentSizeBytes,
    activeSlideIndex: currentSlide,
    priorityIndices,
    nearbyIndices,
    backgroundCount: backgroundIndices.length,
    activeSlideMs,
    nearbyCompletionMs,
    backgroundCompletionMs,
    perThumbnail: summarize(coldDurations),
    nearbyPerThumbnail: summarize(nearbyDurations),
    backgroundPerThumbnail: summarize(backgroundDurations),
    cache: {
      hits: cacheHits,
      misses: cacheMisses,
      renderedSvgCount,
      entries: cache.size,
    },
    concurrency: {
      configuredBatchSize: 1,
      observedMax: 1,
    },
    interactionImpact: {
      maxEventLoopGapMs: round(Math.max(0, ...eventLoopGaps)),
      eventLoopGapCount: eventLoopGaps.length,
      gapsOver16Ms: eventLoopGaps.filter((gap) => gap > 16.7).length,
      gapsOver50Ms: eventLoopGaps.filter((gap) => gap > 50).length,
    },
    measurements,
  };
}

export async function runPptxThumbnailBenchmarks(options = {}) {
  const tiers = options.tiers ?? DEFAULT_PPTX_THUMBNAIL_TIERS;
  const results = [];
  for (const tier of tiers) {
    options.onProgress?.(`Running PPTX thumbnail tier ${tier} slides...`);
    results.push(await runPptxThumbnailBenchmarkTier(tier, options));
  }
  return {
    schemaVersion: 1,
    timestamp: new Date().toISOString(),
    environment: {
      platform: process.platform,
      arch: process.arch,
      nodeVersion: process.version,
      heapTotalBytes: process.memoryUsage?.().heapTotal ?? 0,
    },
    results,
  };
}
