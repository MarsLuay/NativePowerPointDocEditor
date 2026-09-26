import { extractZip } from 'pptx-svg';

export type ChartDataLayoutMode = 'chart' | 'excel';

export interface ChartDataTableDraft {
  categoryLabel: string;
  categories: string[];
  series: Array<{
    name: string;
    pointLabels: string[] | null;
    values: string[];
  }>;
}

function parseXml(contents: string, partPath: string): XMLDocument {
  const doc = new DOMParser().parseFromString(contents, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`Could not parse spreadsheet part: ${partPath}`);
  }
  return doc;
}

function getDescendants(element: Element | XMLDocument, localName: string): Element[] {
  return Array.from(element.getElementsByTagNameNS('*', localName));
}

function parseColumnNumber(columnName: string): number {
  let value = 0;
  for (const character of columnName.toUpperCase()) {
    value = value * 26 + character.charCodeAt(0) - 64;
  }
  return value;
}

function parseCellReference(reference: string): { column: number; row: number } | null {
  const match = reference.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  return { column: parseColumnNumber(match[1] ?? ''), row: Number(match[2]) };
}

function detectDelimiter(text: string): string {
  const firstLine = text.split(/\r?\n/, 1)[0] ?? '';
  const commas = (firstLine.match(/,/g) ?? []).length;
  const tabs = (firstLine.match(/\t/g) ?? []).length;
  const semis = (firstLine.match(/;/g) ?? []).length;
  if (tabs >= commas && tabs >= semis && tabs > 0) return '\t';
  if (semis > commas) return ';';
  return ',';
}

function parseCsvLine(line: string, delimiter: string): string[] {
  const cells: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let index = 0; index < line.length; index++) {
    const character = line[index] ?? '';
    if (character === '"') {
      if (inQuotes && line[index + 1] === '"') {
        current += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (character === delimiter && !inQuotes) {
      cells.push(current);
      current = '';
      continue;
    }

    current += character;
  }

  cells.push(current);
  return cells;
}

/** Parse CSV/TSV text into a dense matrix (ragged rows padded). */
export function parseDelimitedChartMatrix(text: string, delimiter?: string): string[][] {
  const resolvedDelimiter = delimiter ?? detectDelimiter(text);
  const rows = text
    .replace(/^\uFEFF/, '')
    .split(/\r?\n/)
    .map((line) => line.trimEnd())
    .filter((line, index, all) => line.length > 0 || index < all.length - 1)
    .filter((line) => line.length > 0)
    .map((line) => parseCsvLine(line, resolvedDelimiter));

  if (rows.length === 0) {
    throw new Error('The spreadsheet is empty.');
  }

  const width = Math.max(...rows.map((row) => row.length));
  return rows.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ''));
}

function readSharedStrings(workbookFiles: Awaited<ReturnType<typeof extractZip>>): string[] {
  const sharedXml = workbookFiles.textFiles.get('xl/sharedStrings.xml');
  if (!sharedXml) return [];

  const doc = parseXml(sharedXml, 'xl/sharedStrings.xml');
  return getDescendants(doc, 'si').map((item) =>
    getDescendants(item, 't')
      .map((text) => text.textContent ?? '')
      .join('')
  );
}

function readCellDisplayValue(cell: Element, sharedStrings: string[]): string {
  const type = cell.getAttribute('t') ?? '';
  if (type === 'inlineStr') {
    return getDescendants(cell, 't')
      .map((text) => text.textContent ?? '')
      .join('');
  }

  const valueText = getDescendants(cell, 'v')[0]?.textContent ?? '';
  if (type === 's') {
    const index = Number(valueText);
    return Number.isInteger(index) ? sharedStrings[index] ?? '' : '';
  }

  if (type === 'b') {
    return valueText === '1' ? 'TRUE' : 'FALSE';
  }

  return valueText;
}

function getFirstWorksheetPath(workbookFiles: Awaited<ReturnType<typeof extractZip>>): string {
  const workbookXml = workbookFiles.textFiles.get('xl/workbook.xml');
  const relationshipsXml = workbookFiles.textFiles.get('xl/_rels/workbook.xml.rels');
  if (!workbookXml || !relationshipsXml) {
    throw new Error('The Excel workbook is missing sheet metadata.');
  }

  const workbookDoc = parseXml(workbookXml, 'xl/workbook.xml');
  const sheet = getDescendants(workbookDoc, 'sheet')[0];
  const relationshipId =
    sheet?.getAttributeNS(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'id'
    ) || sheet?.getAttribute('r:id');
  if (!relationshipId) {
    throw new Error('The Excel workbook has no worksheets.');
  }

  const relationships = getDescendants(parseXml(relationshipsXml, 'xl/_rels/workbook.xml.rels'), 'Relationship');
  const relationship = relationships.find((element) => element.getAttribute('Id') === relationshipId);
  const target = relationship?.getAttribute('Target');
  if (!target) {
    throw new Error('Could not resolve the first Excel worksheet.');
  }

  const normalized = target.replace(/^\//, '');
  return normalized.startsWith('xl/') ? normalized : `xl/${normalized}`;
}

/** Read the first worksheet of an .xlsx into a dense string matrix. */
export async function parseExcelChartMatrix(buffer: ArrayBuffer): Promise<string[][]> {
  const workbookFiles = await extractZip(buffer);
  const sheetPath = getFirstWorksheetPath(workbookFiles);
  const sheetXml = workbookFiles.textFiles.get(sheetPath);
  if (!sheetXml) {
    throw new Error(`Missing Excel worksheet: ${sheetPath}`);
  }

  const sharedStrings = readSharedStrings(workbookFiles);
  const sheetDoc = parseXml(sheetXml, sheetPath);
  let maxRow = 0;
  let maxColumn = 0;
  const cells = new Map<string, string>();

  for (const cell of getDescendants(sheetDoc, 'c')) {
    const reference = cell.getAttribute('r');
    if (!reference) continue;
    const location = parseCellReference(reference);
    if (!location) continue;
    maxRow = Math.max(maxRow, location.row);
    maxColumn = Math.max(maxColumn, location.column);
    cells.set(`${location.row}:${location.column}`, readCellDisplayValue(cell, sharedStrings));
  }

  if (maxRow < 1 || maxColumn < 1) {
    throw new Error('The Excel worksheet is empty.');
  }

  const matrix: string[][] = [];
  for (let row = 1; row <= maxRow; row++) {
    const values: string[] = [];
    for (let column = 1; column <= maxColumn; column++) {
      values.push(cells.get(`${row}:${column}`) ?? '');
    }
    matrix.push(values);
  }

  while (matrix.length > 0 && matrix[matrix.length - 1]?.every((value) => value.trim() === '')) {
    matrix.pop();
  }

  if (matrix.length === 0) {
    throw new Error('The Excel worksheet is empty.');
  }

  return matrix;
}

function looksNumeric(value: string): boolean {
  const trimmed = value.trim().replace(/,/g, '');
  if (trimmed === '') return false;
  return Number.isFinite(Number(trimmed));
}

function isBlank(value: string | undefined): boolean {
  return (value ?? '').trim() === '';
}

function transposeMatrix(matrix: string[][]): string[][] {
  const height = matrix.length;
  const width = Math.max(...matrix.map((row) => row.length), 0);
  const result: string[][] = [];
  for (let column = 0; column < width; column++) {
    const row: string[] = [];
    for (let rowIndex = 0; rowIndex < height; rowIndex++) {
      row.push(matrix[rowIndex]?.[column] ?? '');
    }
    result.push(row);
  }
  return result;
}

function shouldTransposeImportedMatrix(matrix: string[][]): boolean {
  if (matrix.length < 2 || (matrix[0]?.length ?? 0) < 2) return false;

  const headerRow = matrix[0] ?? [];
  const firstColumn = matrix.map((row) => row[0] ?? '');
  const headerTail = headerRow.slice(1).filter((value) => !isBlank(value));
  const columnTail = firstColumn.slice(1).filter((value) => !isBlank(value));
  if (headerTail.length === 0 || columnTail.length === 0) return false;

  const headerNumeric = headerTail.every(looksNumeric);
  const columnNumeric = columnTail.every(looksNumeric);
  const bodyValues = matrix.slice(1).flatMap((row) => row.slice(1)).filter((value) => !isBlank(value));
  const bodyNumeric = bodyValues.length > 0 && bodyValues.every(looksNumeric);

  // Categories across the top, series names down the left.
  if (headerNumeric && !columnNumeric && bodyNumeric) {
    return true;
  }

  // Wider-than-tall blank-corner sheet usually means categories across.
  if (
    isBlank(headerRow[0])
    && !headerNumeric
    && !columnNumeric
    && bodyNumeric
    && headerTail.length > columnTail.length
  ) {
    return true;
  }

  return false;
}

/**
 * Convert a spreadsheet matrix into chart draft data.
 * Supports chart layout (Category | Series…) and Excel datasheet layout (blank A1).
 */
export function matrixToChartDraft(
  matrix: string[][],
  options: {
    defaultCategoryLabel?: string;
    preferLayout?: ChartDataLayoutMode;
  } = {}
): ChartDataTableDraft {
  if (matrix.length < 2 || (matrix[0]?.length ?? 0) < 2) {
    throw new Error('Chart data needs a header row and at least one data row.');
  }

  let working = matrix.map((row) => [...row]);
  if (shouldTransposeImportedMatrix(working)) {
    working = transposeMatrix(working);
  }

  const header = working[0] ?? [];
  const excelLayout = options.preferLayout === 'excel' || isBlank(header[0]);
  const categoryLabel = excelLayout
    ? options.defaultCategoryLabel || 'Category'
    : header[0]!.trim() || options.defaultCategoryLabel || 'Category';

  const seriesNames = header.slice(1).map((name, index) => name.trim() || `Series ${index + 1}`);
  if (seriesNames.length === 0) {
    throw new Error('Chart data needs at least one series column.');
  }

  const categories: string[] = [];
  const seriesValues = seriesNames.map(() => [] as string[]);

  for (const row of working.slice(1)) {
    const category = (row[0] ?? '').trim();
    if (category === '' && row.slice(1).every((value) => isBlank(value))) {
      continue;
    }

    categories.push(category || `Category ${categories.length + 1}`);
    seriesNames.forEach((_, seriesIndex) => {
      seriesValues[seriesIndex]?.push(row[seriesIndex + 1] ?? '');
    });
  }

  if (categories.length === 0) {
    throw new Error('Chart data needs at least one category row.');
  }

  return {
    categoryLabel,
    categories,
    series: seriesNames.map((name, index) => ({
      name,
      pointLabels: null,
      values: seriesValues[index] ?? [],
    })),
  };
}

export function draftFromChartGrid(
  grid: {
    categoryLabel: string;
    categories: string[];
    series: Array<{ name: string; pointLabels: string[] | null; values: string[] }>;
  }
): ChartDataTableDraft {
  return {
    categoryLabel: grid.categoryLabel,
    categories: [...grid.categories],
    series: grid.series.map((series) => ({
      name: series.name,
      pointLabels: series.pointLabels === null ? null : [...series.pointLabels],
      values: [...series.values],
    })),
  };
}

export async function importChartDataFromFile(
  file: File,
  options: {
    defaultCategoryLabel?: string;
    preferLayout?: ChartDataLayoutMode;
  } = {}
): Promise<ChartDataTableDraft> {
  const name = file.name.toLowerCase();
  const buffer = await file.arrayBuffer();

  let matrix: string[][];
  if (name.endsWith('.csv') || name.endsWith('.tsv') || name.endsWith('.txt')) {
    const text = new TextDecoder('utf-8').decode(buffer);
    matrix = parseDelimitedChartMatrix(text, name.endsWith('.tsv') ? '\t' : undefined);
  } else if (name.endsWith('.xlsx') || name.endsWith('.xlsm')) {
    matrix = await parseExcelChartMatrix(buffer);
  } else {
    throw new Error('Choose an Excel workbook (.xlsx) or CSV/TSV file.');
  }

  return matrixToChartDraft(matrix, options);
}
