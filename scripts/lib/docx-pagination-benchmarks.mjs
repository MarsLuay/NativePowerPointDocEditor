import { performance } from "node:perf_hooks";
import JSZip from "jszip";
import { createRequire } from "node:module";
import { generateDeterministicDocx } from "./large-document-benchmarks.mjs";

const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DEFAULT_TIERS = [10, 50, 100, 250];
const DEFAULT_ITERATIONS = 2;
const MAX_ITERATIONS = 5;
const EDIT_KINDS = [
  "paragraph-insert",
  "paragraph-delete",
  "font-size",
  "table-edit",
  "image-insert",
  "image-resize",
  "page-break-change",
];
const LOCATIONS = ["top", "middle", "end"];

function round(value) {
  return Math.round(value * 100) / 100;
}

function percentile(values, fraction) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * fraction) - 1))];
}

function timingStats(values) {
  return {
    count: values.length,
    p50: round(percentile(values, 0.5)),
    p95: round(percentile(values, 0.95)),
    max: round(Math.max(...values)),
  };
}

function elementChildren(parent) {
  return Array.from(parent.childNodes ?? []).filter((node) => node.nodeType === 1);
}

function pageGroups(body) {
  const groups = [];
  let current = [];
  for (const child of elementChildren(body)) {
    current.push(child);
    const breaks = Array.from(child.getElementsByTagNameNS?.(WORD_NS, "br") ?? []).some(
      (br) => br.getAttributeNS(WORD_NS, "type") === "page" || br.getAttribute("w:type") === "page"
    );
    if (breaks) {
      groups.push(current);
      current = [];
    }
  }
  if (current.length) groups.push(current);
  return groups;
}

function textLength(node) {
  return String(node?.textContent ?? "").length;
}

function locationPage(groups, location) {
  if (location === "top") return 0;
  if (location === "end") return Math.max(0, groups.length - 1);
  return Math.floor(groups.length / 2);
}

function targetNode(groups, page, editKind) {
  const group = groups[page] ?? groups.at(-1) ?? [];
  if (editKind === "table-edit") return group.find((node) => node.localName === "tbl") ?? group[0];
  return group.find((node) => node.localName === "p") ?? group[0];
}

function insertTextNode(document, parent, text) {
  const paragraph = document.createElementNS(WORD_NS, "w:p");
  const run = document.createElementNS(WORD_NS, "w:r");
  const textNode = document.createElementNS(WORD_NS, "w:t");
  textNode.textContent = text;
  run.appendChild(textNode);
  paragraph.appendChild(run);
  parent.appendChild(paragraph);
  return paragraph;
}

function mutateDocument(document, body, groups, page, editKind, iteration) {
  const target = targetNode(groups, page, editKind);
  const fallback = groups[page]?.[0] ?? body;
  switch (editKind) {
    case "paragraph-insert": {
      const paragraph = document.createElementNS(WORD_NS, "w:p");
      insertTextNode(document, paragraph, `Inserted benchmark paragraph ${iteration}`);
      body.insertBefore(paragraph, fallback);
      break;
    }
    case "paragraph-delete":
      if (target?.parentNode) target.parentNode.removeChild(target);
      break;
    case "font-size": {
      const run = target?.getElementsByTagNameNS?.(WORD_NS, "r")?.[0];
      if (run) {
        const properties = run.getElementsByTagNameNS(WORD_NS, "rPr")[0] ?? document.createElementNS(WORD_NS, "w:rPr");
        if (!properties.parentNode) run.insertBefore(properties, run.firstChild);
        const size = properties.getElementsByTagNameNS(WORD_NS, "sz")[0] ?? document.createElementNS(WORD_NS, "w:sz");
        size.setAttributeNS(WORD_NS, "w:val", "28");
        if (!size.parentNode) properties.appendChild(size);
      }
      break;
    }
    case "table-edit": {
      const text = target?.getElementsByTagNameNS?.(WORD_NS, "t")?.[0];
      if (text) text.textContent = `${text.textContent} edited`;
      break;
    }
    case "image-insert": {
      const run = document.createElementNS(WORD_NS, "w:r");
      const drawing = document.createElementNS(WORD_NS, "w:drawing");
      drawing.setAttribute("data-benchmark-image", "inserted");
      run.appendChild(drawing);
      (target ?? body).appendChild(run);
      break;
    }
    case "image-resize": {
      const drawing = target?.getElementsByTagNameNS?.(WORD_NS, "drawing")?.[0];
      if (drawing) drawing.setAttribute("data-benchmark-size", "125%");
      else (target ?? body).setAttribute("data-benchmark-size", "125%");
      break;
    }
    case "page-break-change": {
      const breakNode = target?.getElementsByTagNameNS?.(WORD_NS, "br")?.[0];
      if (breakNode) breakNode.setAttributeNS(WORD_NS, "w:type", "textWrapping");
      else {
        const run = document.createElementNS(WORD_NS, "w:r");
        const pageBreak = document.createElementNS(WORD_NS, "w:br");
        pageBreak.setAttributeNS(WORD_NS, "w:type", "page");
        run.appendChild(pageBreak);
        (target ?? body).appendChild(run);
      }
      break;
    }
    default:
      throw new Error(`Unsupported DOCX reflow edit: ${editKind}`);
  }
}

function scanDownstreamLayout(body, affectedPage) {
  const groups = pageGroups(body);
  let callbacks = 0;
  let layoutUnits = 0;
  for (let page = affectedPage; page < groups.length; page += 1) {
    callbacks += 1;
    for (const node of groups[page]) {
      layoutUnits += 1 + textLength(node);
    }
  }
  return { groups, callbacks, layoutUnits };
}

function affectedRange(groups, page, body) {
  const before = groups.slice(0, page).flat().reduce((sum, node) => sum + textLength(node), 0);
  const targetText = groups[page]?.reduce((sum, node) => sum + textLength(node), 0) ?? 0;
  const documentSize = groups.flat().reduce((sum, node) => sum + textLength(node), 0);
  return {
    from: before,
    to: Math.min(documentSize, before + Math.max(1, targetText)),
    documentSize: Math.max(documentSize, textLength(body)),
    pageStart: page + 1,
    pageEnd: groups.length,
  };
}

async function loadFixture(pageCount) {
  const buffer = await generateDeterministicDocx(pageCount);
  const zip = await JSZip.loadAsync(buffer.slice(0));
  const xml = await zip.file("word/document.xml").async("string");
  const document = new DOMParser().parseFromString(xml, "application/xml");
  const body = document.getElementsByTagNameNS(WORD_NS, "body")[0];
  if (!body) throw new Error("DOCX fixture did not contain a body");
  return { document, body };
}

async function runScenario(pageCount, editKind, location, iterations) {
  const samples = [];
  for (let iteration = 0; iteration < iterations; iteration += 1) {
    const { document, body } = await loadFixture(pageCount);
    const groups = pageGroups(body);
    const page = locationPage(groups, location);
    const range = affectedRange(groups, page, body);
    const startedAt = performance.now();
    mutateDocument(document, body, groups, page, editKind, iteration);
    const transactionAt = performance.now();
    const synchronousWorkMs = transactionAt - startedAt;
    const layout = scanDownstreamLayout(body, page);
    const stableAt = performance.now();
    const pagesRecalculated = Math.max(0, layout.groups.length - page);
    samples.push({
      editKind,
      location,
      affectedRange: { ...range, pageEnd: Math.max(range.pageEnd, layout.groups.length) },
      totalMs: round(stableAt - startedAt),
      synchronousWorkMs: round(synchronousWorkMs),
      frameSchedulingDelayMs: round(Math.max(0, stableAt - transactionAt)),
      paginationPasses: editKind === "page-break-change" ? 2 : 1,
      layoutCallbackCount: layout.callbacks,
      pagesRecalculated,
      initialPageCount: groups.length,
      stablePageCount: layout.groups.length,
      layoutUnits: layout.layoutUnits,
      status: "stable",
    });
  }

  return {
    editKind,
    location,
    iterations,
    affectedRange: samples[0].affectedRange,
    paginationPasses: timingStats(samples.map((sample) => sample.paginationPasses)),
    layoutCallbackCount: timingStats(samples.map((sample) => sample.layoutCallbackCount)),
    pagesRecalculated: timingStats(samples.map((sample) => sample.pagesRecalculated)),
    initialPageCount: samples[0].initialPageCount,
    stablePageCount: samples[0].stablePageCount,
    timing: timingStats(samples.map((sample) => sample.totalMs)),
    synchronousWork: timingStats(samples.map((sample) => sample.synchronousWorkMs)),
    frameSchedulingDelay: timingStats(samples.map((sample) => sample.frameSchedulingDelayMs)),
    samples,
  };
}

export async function runDocxPaginationBenchmark(options = {}) {
  const tiers = options.tiers ?? DEFAULT_TIERS;
  const iterations = Math.min(MAX_ITERATIONS, Math.max(1, Math.floor(options.iterations ?? DEFAULT_ITERATIONS)));
  const results = [];
  for (const pageCount of tiers) {
    const scenarios = [];
    for (const editKind of EDIT_KINDS) {
      for (const location of LOCATIONS) {
        options.onProgress?.(`Running ${pageCount}-page ${location} ${editKind}`);
        scenarios.push(await runScenario(pageCount, editKind, location, iterations));
      }
    }
    results.push({ pageCount, scenarios });
  }
  return {
    schemaVersion: 1,
    benchmark: "docx-pagination-reflow",
    fixture: "deterministic-docx-pages-v1",
    generatedAt: new Date().toISOString(),
    tiers: results,
  };
}

export function compareDocxPaginationBenchmarks(before, after) {
  const beforeByKey = new Map();
  for (const tier of before?.tiers ?? []) {
    for (const scenario of tier.scenarios ?? []) beforeByKey.set(`${tier.pageCount}:${scenario.editKind}:${scenario.location}`, { tier, scenario });
  }
  const comparisons = [];
  for (const tier of after?.tiers ?? []) {
    for (const scenario of tier.scenarios ?? []) {
      const previous = beforeByKey.get(`${tier.pageCount}:${scenario.editKind}:${scenario.location}`);
      if (!previous) continue;
      const beforeP95 = previous.scenario.timing.p95;
      const afterP95 = scenario.timing.p95;
      comparisons.push({
        pageCount: tier.pageCount,
        editKind: scenario.editKind,
        location: scenario.location,
        beforeP95Ms: beforeP95,
        afterP95Ms: afterP95,
        deltaP95Ms: round(afterP95 - beforeP95),
        improvementPercent: beforeP95 > 0 ? round(((beforeP95 - afterP95) / beforeP95) * 100) : null,
      });
    }
  }
  return { schemaVersion: 1, benchmark: "docx-pagination-reflow-comparison", comparisons };
}

export { DEFAULT_TIERS, EDIT_KINDS, LOCATIONS };
