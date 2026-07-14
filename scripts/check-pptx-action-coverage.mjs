import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const testsDirectory = path.join(projectRoot, "tests");

const CHECKLIST = [
  { operation: "open", file: "matrix-pptx-open.test.mjs", patterns: [/loadEngine/, /malformed/] },
  { operation: "insert", file: "matrix-pptx-insert.test.mjs", patterns: [/addShapeGeometry/, /addImage/, /addTable/, /addChart/, /addTextBox/] },
  { operation: "arrange", file: "matrix-pptx-arrange.test.mjs", patterns: [/nudgeShapes/, /reorderShapes/, /groupShapes/], skipPattern: /t\.skip/ },
  { operation: "text", file: "matrix-pptx-text.test.mjs", patterns: [/updateParagraphText/, /setRunStyle/, /setParagraphAlignment/] },
  { operation: "clipboard", file: "matrix-pptx-clipboard-find.test.mjs", patterns: [/copyShape/, /pasteShape/, /duplicateShape/] },
  { operation: "find", file: "matrix-pptx-clipboard-find.test.mjs", patterns: [/replaceText/] },
  { operation: "history", file: "matrix-pptx-history.test.mjs", patterns: [/HistoryController/, /\.undo\(\)/, /\.redo\(\)/] },
  { operation: "save", file: "matrix-pptx-save-export.test.mjs", patterns: [/DocumentSaveCoordinator/, /saveCurrentPresentation/, /SaveController/] },
  { operation: "export", file: "matrix-pptx-save-export.test.mjs", patterns: [/assertExportRoundTrips/, /validatePowerPointExport/] },
  { operation: "slides", file: "matrix-pptx-slides-charts.test.mjs", patterns: [/addSlide/, /duplicateSlide/, /deleteSlide/, /moveSlide/] },
  { operation: "charts", file: "matrix-pptx-slides-charts.test.mjs", patterns: [/getChartDataGrid/, /updateChartData/], skipPattern: /t\.skip/ },
  { operation: "package", file: "matrix-pptx-package.test.mjs", patterns: [/PresentationEngine\.load/, /assertExportRoundTrips/], skipPattern: /t\.skip/ },
];

const fileContents = new Map();
const missingFiles = [];
const missingCoverage = [];
const warnings = [];

for (const item of CHECKLIST) {
  let source = fileContents.get(item.file);
  if (source === undefined) {
    try {
      source = await readFile(path.join(testsDirectory, item.file), "utf8");
      fileContents.set(item.file, source);
    } catch (error) {
      if (error?.code === "ENOENT") {
        missingFiles.push(item.file);
        continue;
      }
      throw error;
    }
  }

  const missingPatterns = item.patterns.filter((pattern) => !pattern.test(source));
  if (missingPatterns.length > 0) {
    missingCoverage.push(
      `${item.operation} (${item.file}): ${missingPatterns.map(String).join(", ")}`,
    );
  }
  if (item.skipPattern?.test(source)) {
    warnings.push(`${item.operation} (${item.file}) contains a skipped matrix scenario.`);
  }
}

const uniqueMissingFiles = [...new Set(missingFiles)];
const uniqueWarnings = [...new Set(warnings)];

if (uniqueWarnings.length > 0) {
  console.warn("PPTX action coverage warnings:");
  for (const warning of uniqueWarnings) console.warn(`- ${warning}`);
}

if (uniqueMissingFiles.length > 0 || missingCoverage.length > 0) {
  console.error("PPTX action coverage failed:");
  for (const file of uniqueMissingFiles) console.error(`- Missing matrix file: tests/${file}`);
  for (const gap of missingCoverage) console.error(`- Missing matrix coverage: ${gap}`);
  process.exitCode = 1;
} else {
  console.log(`PPTX action coverage passed (${CHECKLIST.length} operations checked).`);
}
