import { performance } from "node:perf_hooks";
import { createRequire } from "node:module";
import { generateDeterministicPptx } from "./large-document-benchmarks.mjs";
import { loadPresentationEngineModule } from "../../tests/helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);
const { DOMParser } = require("@xmldom/xmldom");
const DEFAULT_TIERS = [10, 50, 100, 250, 500];
const DEFAULT_ITERATIONS = 1;
const MAX_ITERATIONS = 3;
const CACHE_STATES = ["cold", "warm"];
const DISTANCES = ["adjacent", "distant"];

function round(value) {
  return Math.round(value * 100) / 100;
}

function percentile(values, fraction) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * fraction) - 1))];
}

function stats(values) {
  return {
    count: values.length,
    p50: round(percentile(values, 0.5)),
    p95: round(percentile(values, 0.95)),
    max: round(Math.max(...values)),
  };
}

function targetFor(slideCount, distance) {
  return distance === "adjacent" ? Math.min(1, slideCount - 1) : Math.max(0, slideCount - 1);
}

function summarizeComplexity(slideXml, svgText) {
  return {
    slideXmlCharacters: slideXml.length,
    svgCharacters: svgText.length,
    shapeCount: (svgText.match(/data-ooxml-shape-idx=/g) ?? []).length,
  };
}

async function renderAndParse(engine, slideIndex, cache, cacheState) {
  const packageStarted = performance.now();
  const slideXml = cacheState === "warm" ? "" : engine.getSlideXml(slideIndex);
  const packageXmlMs = performance.now() - packageStarted;

  let svgText = cache.get(slideIndex)?.svgText ?? null;
  let renderMs = 0;
  if (cacheState === "cold" || !svgText) {
    const renderStarted = performance.now();
    svgText = engine.renderSlide(slideIndex).svg;
    renderMs = performance.now() - renderStarted;
    cache.set(slideIndex, { svgText, slideXml: slideXml || engine.getSlideXml(slideIndex) });
  }

  const parseStarted = performance.now();
  const parsed = new DOMParser().parseFromString(svgText, "image/svg+xml");
  const svgParseMs = performance.now() - parseStarted;
  const domSwapStarted = performance.now();
  // A detached document element models the canvas SVG replacement without
  // requiring a browser or retaining a live editor DOM in a benchmark.
  const visibleSvg = parsed.documentElement;
  const domSwapMs = performance.now() - domSwapStarted;
  const frameStarted = performance.now();
  await new Promise((resolve) => setImmediate(resolve));
  const stableAt = performance.now();

  return {
    packageXmlMs,
    renderMs,
    svgParseMs,
    domSwapMs,
    frameSchedulingDelayMs: stableAt - frameStarted,
    inputToVisibleMs: stableAt,
    visibleSvg,
    complexity: summarizeComplexity(cache.get(slideIndex)?.slideXml ?? slideXml, svgText),
  };
}

async function runScenario(engine, slideCount, distance, cacheState, iterations) {
  const target = targetFor(slideCount, distance);
  const samples = [];
  for (let iteration = 0; iteration < iterations; iteration += 1) {
    const cache = new Map();
    if (cacheState === "warm") {
      const warmup = engine.renderSlide(target).svg;
      cache.set(target, { svgText: warmup, slideXml: engine.getSlideXml(target) });
    }
    const startedAt = performance.now();
    const result = await renderAndParse(engine, target, cache, cacheState);
    const inputToVisibleMs = performance.now() - startedAt;
    // Keep a stable reference until the sample is serialized, mirroring the
    // active canvas ownership without retaining prior samples.
    void result.visibleSvg;
    samples.push({
      fromSlide: 0,
      toSlide: target,
      distance,
      distanceSlides: target,
      cacheState,
      inputToVisibleMs: round(inputToVisibleMs),
      stageTimingsMs: {
        packageXml: round(result.packageXmlMs),
        render: round(result.renderMs),
        svgParse: round(result.svgParseMs),
        domSwap: round(result.domSwapMs),
        frameSchedulingDelay: round(result.frameSchedulingDelayMs),
      },
      complexity: result.complexity,
      status: "stable",
      iteration,
    });
  }

  return {
    cacheState,
    distance,
    fromSlide: 0,
    toSlide: target,
    distanceSlides: target,
    iterations,
    complexity: samples[0].complexity,
    timing: stats(samples.map((sample) => sample.inputToVisibleMs)),
    stageTimingsMs: Object.fromEntries(
      Object.keys(samples[0].stageTimingsMs).map((stage) => [stage, stats(samples.map((sample) => sample.stageTimingsMs[stage]))])
    ),
    samples,
  };
}

export async function runPptxSlideSwitchBenchmarks(options = {}) {
  const tiers = options.tiers ?? DEFAULT_TIERS;
  const iterations = Math.min(MAX_ITERATIONS, Math.max(1, Math.floor(options.iterations ?? DEFAULT_ITERATIONS)));
  const { PresentationEngine } = await loadPresentationEngineModule();
  const results = [];

  for (const slideCount of tiers) {
    options.onProgress?.(`Loading ${slideCount}-slide PPTX fixture`);
    const buffer = generateDeterministicPptx(slideCount);
    const engine = await PresentationEngine.load(buffer.slice(0));
    const scenarios = [];
    for (const cacheState of CACHE_STATES) {
      for (const distance of DISTANCES) {
        options.onProgress?.(`Running ${slideCount}-slide ${cacheState} ${distance} switch`);
        scenarios.push(await runScenario(engine, slideCount, distance, cacheState, iterations));
      }
    }
    const allSamples = scenarios.flatMap((scenario) => scenario.samples);
    results.push({
      slideCount,
      scenarios,
      timing: stats(allSamples.map((sample) => sample.inputToVisibleMs)),
      worstSwitches: [...allSamples]
        .sort((left, right) => right.inputToVisibleMs - left.inputToVisibleMs)
        .slice(0, 10),
    });
  }

  return {
    schemaVersion: 1,
    benchmark: "pptx-slide-switch",
    fixture: "deterministic-pptx-slides-v1",
    generatedAt: new Date().toISOString(),
    tiers: results,
  };
}

export function comparePptxSlideSwitchBenchmarks(before, after) {
  const beforeByKey = new Map();
  for (const tier of before?.tiers ?? []) {
    for (const scenario of tier.scenarios ?? []) {
      beforeByKey.set(`${tier.slideCount}:${scenario.cacheState}:${scenario.distance}`, scenario);
    }
  }
  const comparisons = [];
  for (const tier of after?.tiers ?? []) {
    for (const scenario of tier.scenarios ?? []) {
      const previous = beforeByKey.get(`${tier.slideCount}:${scenario.cacheState}:${scenario.distance}`);
      if (!previous) continue;
      comparisons.push({
        slideCount: tier.slideCount,
        cacheState: scenario.cacheState,
        distance: scenario.distance,
        beforeP95Ms: previous.timing.p95,
        afterP95Ms: scenario.timing.p95,
        beforeMaxMs: previous.timing.max,
        afterMaxMs: scenario.timing.max,
        deltaP95Ms: round(scenario.timing.p95 - previous.timing.p95),
        deltaMaxMs: round(scenario.timing.max - previous.timing.max),
      });
    }
  }
  return { schemaVersion: 1, benchmark: "pptx-slide-switch-comparison", comparisons };
}

export { DEFAULT_TIERS, CACHE_STATES, DISTANCES };
