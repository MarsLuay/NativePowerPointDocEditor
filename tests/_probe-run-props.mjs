// Empirical probe: which authored run properties does the pptx-svg Wasm model
// drop when a Wasm-primitive edit (here: updateShapeTransform) re-serializes a
// slide? Authors a run decorated with every "likely suspect" rPr attribute and
// child, then diffs the run's <a:rPr> before vs after the transform.
//
//   node tests/_probe-run-props.mjs
import { buildZip, extractZip } from "pptx-svg";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";

const fixtureTitleParagraph =
  '<a:p><a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>Native PowerPoint fixture</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>';

// A run wearing one of every property we suspect the renderer doesn't model.
// rPr child order follows CT_TextCharacterProperties (ln, fill, effectLst,
// highlight, latin, ea, cs, sym, hlinkClick, ...).
const decoratedParagraph =
  '<a:p>' +
  '<a:r>' +
  '<a:rPr lang="en-US" sz="1800" b="1" i="1" u="sng" strike="sngStrike"' +
  ' spc="300" kern="1200" baseline="30000" cap="all" normalizeH="1" spcPts="0">' +
  '<a:ln w="3175"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>' +
  '<a:solidFill><a:srgbClr val="123456"/></a:solidFill>' +
  '<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>' +
  '<a:uFill><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:uFill>' +
  '<a:latin typeface="Arial"/>' +
  '<a:ea typeface="MS Mincho"/>' +
  '<a:cs typeface="Arial"/>' +
  '<a:sym typeface="Wingdings"/>' +
  '<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/>' +
  '</a:rPr>' +
  '<a:t>Decorated run</a:t>' +
  '</a:r>' +
  '<a:endParaRPr lang="en-US"/>' +
  '</a:p>';

// rPr child effectLst lives just after fill per the schema; PowerPoint also
// accepts it, but pptx-svg can be picky about order. Insert it right after ln.
const decoratedWithEffect = decoratedParagraph.replace(
  '<a:solidFill><a:srgbClr val="123456"/></a:solidFill>',
  '<a:solidFill><a:srgbClr val="123456"/></a:solidFill>' +
    '<a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>'
);

function firstRunRpr(ooxml) {
  const doc = new globalThis.DOMParser().parseFromString(ooxml, "application/xml");
  // Find the run whose text is "Decorated run".
  const runs = Array.from(doc.getElementsByTagName("a:r"));
  const run = runs.find(
    (r) => (r.getElementsByTagName("a:t")[0]?.textContent || "") === "Decorated run"
  );
  if (!run) return null;
  return run.getElementsByTagName("a:rPr")[0] ?? null;
}

const ATTRS = ["b", "i", "u", "strike", "spc", "kern", "baseline", "cap", "normalizeH"];
const CHILDREN = ["ln", "solidFill", "effectLst", "highlight", "uFill", "latin", "ea", "cs", "sym", "hlinkClick"];

function describe(rPr) {
  if (!rPr) return null;
  const out = {};
  for (const a of ATTRS) out[a] = rPr.getAttribute(a);
  for (const c of CHILDREN) {
    out[c] = Array.from(rPr.childNodes).some((n) => n.nodeType === 1 && n.localName === c);
  }
  return out;
}

async function run() {
  const input = await readDeck("features.pptx");
  const source = toArrayBuffer(input);
  const zip = await extractZip(source);
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = zip.textFiles.get(slidePath);
  const patched = await buildZip(
    source,
    new Map([[slidePath, slideXml.replace(fixtureTitleParagraph, decoratedWithEffect)]])
  );

  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(patched);
  const slide = 0;

  const beforeRpr = describe(firstRunRpr(engine.renderer.getSlideOoxml(slide)));
  if (!beforeRpr) {
    console.log("Decorated run not found in the loaded model — the renderer may have rejected the authored rPr.");
    return;
  }

  // The decorated run sits on shape 0 (the former title). Find its renderer
  // shape index by scanning for the text, then move that shape.
  let shapeIndex = 0;
  for (let i = 0; i < 32; i++) {
    if (engine.getParagraphRunText(slide, i, 0)?.includes("Decorated run")) {
      shapeIndex = i;
      break;
    }
  }

  await engine.updateShapeTransform(slide, shapeIndex, { x: 120000, y: 120000, cx: 4000000, cy: 1000000, rot: 0 });

  // The renderer's live model after the edit (what it drops).
  const liveRpr = describe(firstRunRpr(engine.renderer.getSlideOoxml(slide)));
  // What survives once the engine reconciles + the file is reopened (the fix).
  const exported = await engine.export();
  const { PresentationEngine: PE2 } = await loadPresentationEngineModule();
  const reopened = await PE2.load(exported);
  const savedRpr = describe(firstRunRpr(reopened.renderer.getSlideOoxml(slide)));

  const val = (rpr, key) =>
    !rpr ? "—" : ATTRS.includes(key) ? rpr[key] ?? "—" : rpr[key] ? "present" : "DROPPED";

  const rows = [];
  for (const key of [...ATTRS, ...CHILDREN]) {
    const wasPresent = ATTRS.includes(key) ? beforeRpr[key] !== null : beforeRpr[key];
    if (!wasPresent) continue;
    rows.push({
      property: ATTRS.includes(key) ? `@${key}` : `<a:${key}>`,
      authored: ATTRS.includes(key) ? beforeRpr[key] : "present",
      "renderer (live)": val(liveRpr, key),
      "engine (saved)": val(savedRpr, key),
    });
  }

  console.log("\nRun property survival: authored -> renderer live model -> engine save/reopen\n");
  console.table(rows);
  const rendererDropped = rows.filter((r) => r["renderer (live)"] === "DROPPED" || r["renderer (live)"] === "—");
  const savedRecovered = rendererDropped.filter((r) => r["engine (saved)"] !== "DROPPED" && r["engine (saved)"] !== "—");
  console.log("\nRenderer drops:", rendererDropped.map((r) => r.property).join(", ") || "(none)");
  console.log("Engine recovers:", savedRecovered.map((r) => r.property).join(", ") || "(none)");
}

run();
