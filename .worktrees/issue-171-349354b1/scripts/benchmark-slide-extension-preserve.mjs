import { performance } from "node:perf_hooks";
import { DOMParser } from "@xmldom/xmldom";
import { buildZip } from "pptx-svg";
import { loadSlideExtensionPreserveModule } from "../tests/helpers/load-plugin-modules.mjs";

function createTestSlideXml(shapeCount, reverse = false) {
  let xml = `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>`;
  const indices = [];
  for (let i = 0; i < shapeCount; i++) indices.push(i);
  if (reverse) indices.reverse();
  for (const i of indices) {
    xml += `<p:sp><p:nvSpPr><p:cNvPr id="${i}" name="Shape ${i}"><a:extLst><a:ext uri="{123}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{ABC${i}}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${i * 10}" y="${i * 10}"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Shape Text ${i}</a:t></a:r></a:p></p:txBody></p:sp>`;
  }
  xml += `</p:spTree></p:cSld></p:sld>`;
  return xml;
}

// Helpers for old vs new comparison
function getDescendants(node, tag) {
  return Array.from(node.getElementsByTagName("*")).filter(el => el.localName === tag);
}
function getDirectChild(node, tag) {
  for (let c = node.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && c.localName === tag) return c;
  }
  return null;
}
function getShapeFingerprint(shape) {
  if (!shape) return "";
  const transform = getDescendants(shape, "xfrm")[0];
  const offset = transform ? getDirectChild(transform, "off") : null;
  const extent = transform ? getDirectChild(transform, "ext") : null;
  const geometry = [
    offset?.getAttribute("x") ?? "",
    offset?.getAttribute("y") ?? "",
    extent?.getAttribute("cx") ?? "",
    extent?.getAttribute("cy") ?? ""
  ].join(",");
  const text = (shape.textContent ?? "").replace(/\s+/g, " ").trim().slice(0, 64);
  return `${geometry}|${text}`;
}
function collectShapeIdentities(xmlDocument) {
  return getDescendants(xmlDocument, "cNvPr").map((cNvPr) => {
    const shape = cNvPr.parentNode?.parentNode;
    return {
      cNvPr,
      id: cNvPr.getAttribute("id") ?? "",
      name: cNvPr.getAttribute("name") ?? "",
      extensionList: getDirectChild(cNvPr, "extLst"),
      fingerprint: getShapeFingerprint(shape),
    };
  });
}

function oldMatchingAlgorithm(previousShapes, exportedShapes) {
  const pairs = [];
  const remaining = [...previousShapes];
  for (const exported of exportedShapes) {
    const matchIndex = remaining.findIndex(
      (candidate) => candidate.fingerprint === exported.fingerprint
    );
    const previous = matchIndex >= 0 ? remaining[matchIndex] : undefined;
    if (previous) {
      pairs.push([previous, exported]);
      remaining.splice(matchIndex, 1);
    }
  }
  return pairs;
}

function newMatchingAlgorithm(previousShapes, exportedShapes) {
  const remainingByFingerprint = new Map();
  for (const previous of previousShapes) {
    let entry = remainingByFingerprint.get(previous.fingerprint);
    if (!entry) {
      entry = { items: [], index: 0 };
      remainingByFingerprint.set(previous.fingerprint, entry);
    }
    entry.items.push(previous);
  }

  const pairs = [];
  for (const exported of exportedShapes) {
    const entry = remainingByFingerprint.get(exported.fingerprint);
    if (entry && entry.index < entry.items.length) {
      const previous = entry.items[entry.index++];
      if (previous) {
        pairs.push([previous, exported]);
      }
    }
  }
  return pairs;
}

async function runBenchmark() {
  const parser = new DOMParser();
  const shapeCounts = [200, 500, 1500];

  console.log("=== Direct Matching Algorithm Comparison ===");
  for (const count of shapeCounts) {
    const prevDoc = parser.parseFromString(createTestSlideXml(count, false), "text/xml");
    const expDoc = parser.parseFromString(createTestSlideXml(count - 1, true), "text/xml");

    const previousShapes = collectShapeIdentities(prevDoc);
    const exportedShapes = collectShapeIdentities(expDoc);

    const iterations = 50;

    // Benchmark Old
    const startOld = performance.now();
    for (let i = 0; i < iterations; i++) {
      oldMatchingAlgorithm(previousShapes, exportedShapes);
    }
    const durationOld = performance.now() - startOld;

    // Benchmark New
    const startNew = performance.now();
    for (let i = 0; i < iterations; i++) {
      newMatchingAlgorithm(previousShapes, exportedShapes);
    }
    const durationNew = performance.now() - startNew;

    const speedup = (durationOld / durationNew).toFixed(1);
    const pctReduction = (((durationOld - durationNew) / durationOld) * 100).toFixed(1);

    console.log(`\nShape count: ${count} (${iterations} iterations)`);
    console.log(`  Old (O(N^2) findIndex + splice): ${durationOld.toFixed(2)} ms (avg ${(durationOld / iterations).toFixed(3)} ms/op)`);
    console.log(`  New (O(N) Map lookup):           ${durationNew.toFixed(2)} ms (avg ${(durationNew / iterations).toFixed(3)} ms/op)`);
    console.log(`  Speedup:                         ${speedup}x faster (${pctReduction}% time reduction)`);
  }
}

runBenchmark().catch((err) => {
  console.error(err);
  process.exit(1);
});
