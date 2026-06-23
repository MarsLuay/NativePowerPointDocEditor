import {
  DRAWINGML_NAMESPACE,
  RELATIONSHIP_NAMESPACE,
  SHAPE_ELEMENT_NAMES,
  getDescendants,
  getElementChildren,
  getSlidePath,
  getSlideRelationshipsPath,
  hasAncestor,
  normalizeLabelText,
  parseXml,
  resolvePartPath,
} from './ooxmlXml';

type ChartAxisOrientation = 'horizontal' | 'vertical';

export interface ChartAxisFormat {
  orientation: ChartAxisOrientation;
  formatCode: string;
  min: number;
  max: number;
  majorUnit: number | null;
  date1904: boolean;
}

interface ChartTickRun {
  orientation: ChartAxisOrientation;
  elements: SVGTextElement[];
}

function getShapeElement(slideDoc: XMLDocument, shapeIndex: number): Element {
  const shapeTree = getDescendants(slideDoc, 'spTree')[0];
  const shape = getElementChildren(shapeTree)
    .filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName))[shapeIndex];
  if (!shape) {
    throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
  }
  return shape;
}

function findChartPartPath(
  textFiles: Map<string, string>,
  slideIndex: number,
  shapeIndex: number
): string {
  const slidePath = getSlidePath(slideIndex);
  const slideXml = textFiles.get(slidePath);
  if (!slideXml) {
    throw new Error(`Missing slide XML part: ${slidePath}`);
  }

  const shape = getShapeElement(parseXml(slideXml, slidePath), shapeIndex);
  const chart = getDescendants(shape, 'chart')[0];
  const relationshipId =
    chart?.getAttributeNS(RELATIONSHIP_NAMESPACE, 'id') ||
    chart?.getAttribute('r:id');
  if (!relationshipId) {
    throw new Error('Could not find the embedded chart relationship.');
  }

  const relationshipsPath = getSlideRelationshipsPath(slideIndex);
  const relationshipsXml = textFiles.get(relationshipsPath);
  if (!relationshipsXml) {
    throw new Error(`Missing slide relationship part: ${relationshipsPath}`);
  }

  const relationships = getDescendants(parseXml(relationshipsXml, relationshipsPath), 'Relationship');
  const relationship = relationships.find((element) => element.getAttribute('Id') === relationshipId);
  const target = relationship?.getAttribute('Target');
  if (!target || relationship?.getAttribute('TargetMode') === 'External') {
    throw new Error('Could not resolve the embedded chart XML part.');
  }

  return resolvePartPath(slidePath, target);
}


export function getChartTextSources(chartDoc: XMLDocument): Element[] {
  const richText = getDescendants(chartDoc, 't')
    .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  const cachedTextContainers = new Set(['strCache', 'strLit']);
  const cachedText = getDescendants(chartDoc, 'v').filter((element) => {
    return element.parentElement?.localName === 'tx' || hasAncestor(element, cachedTextContainers);
  });

  return [...richText, ...cachedText];
}

function getChartTextValues(textFiles: Map<string, string>, slideIndex: number, shapeIndex: number): string[] {
  const chartPath = findChartPartPath(textFiles, slideIndex, shapeIndex);
  const chartXml = textFiles.get(chartPath);
  if (!chartXml) {
    throw new Error(`Missing chart XML part: ${chartPath}`);
  }

  return getChartTextSources(parseXml(chartXml, chartPath))
    .map((element) => normalizeLabelText(element.textContent || ''))
    .filter(Boolean);
}

function getValAttribute(element: Element, localName: string): string | null {
  return getDescendants(element, localName)[0]?.getAttribute('val') ?? null;
}

function parseFiniteNumber(value: string | null): number | null {
  if (value === null || value.trim() === '') {
    return null;
  }

  const parsedValue = Number(value);
  return Number.isFinite(parsedValue) ? parsedValue : null;
}

function getCachedNumbers(chartDoc: XMLDocument, containerNames: string[]): number[] {
  const values: number[] = [];

  for (const containerName of containerNames) {
    for (const container of getDescendants(chartDoc, containerName)) {
      for (const point of getDescendants(container, 'pt')) {
        const value = parseFiniteNumber(
          getElementChildren(point).find((element) => element.localName === 'v')?.textContent ?? null
        );

        if (value !== null) {
          values.push(value);
        }
      }
    }
  }

  return values;
}

function getNiceStep(range: number): number {
  if (!Number.isFinite(range) || range <= 0) {
    return 1;
  }

  const roughStep = range / 5;
  const magnitude = 10 ** Math.floor(Math.log10(roughStep));
  const normalizedStep = roughStep / magnitude;

  if (normalizedStep <= 1) return magnitude;
  if (normalizedStep <= 2) return 2 * magnitude;
  if (normalizedStep <= 5) return 5 * magnitude;
  return 10 * magnitude;
}

function deriveAxisBounds(
  axisElement: Element,
  values: number[],
  includeZero: boolean
): { min: number; max: number; majorUnit: number | null } {
  const scaling = getDescendants(axisElement, 'scaling')[0] ?? axisElement;
  const explicitMin = parseFiniteNumber(getValAttribute(scaling, 'min'));
  const explicitMax = parseFiniteNumber(getValAttribute(scaling, 'max'));
  const explicitMajorUnit = parseFiniteNumber(getValAttribute(axisElement, 'majorUnit'));
  const fallbackMin = values.length > 0 ? Math.min(...values) : 0;
  const fallbackMax = values.length > 0 ? Math.max(...values) : 1;
  const dataMin = includeZero ? Math.min(0, fallbackMin) : fallbackMin;
  const dataMax = includeZero ? Math.max(0, fallbackMax) : fallbackMax;
  const step = explicitMajorUnit ?? getNiceStep(Math.max(dataMax - dataMin, Number.EPSILON));
  let min = explicitMin ?? (explicitMajorUnit === null ? dataMin : Math.floor(dataMin / step) * step);
  let max = explicitMax ?? (explicitMajorUnit === null ? dataMax : Math.ceil(dataMax / step) * step);

  if (!Number.isFinite(min)) {
    min = 0;
  }

  if (!Number.isFinite(max) || max <= min) {
    max = min + step * 5;
  }

  return { min, max, majorUnit: explicitMajorUnit };
}

function getChartAxisFormats(textFiles: Map<string, string>, slideIndex: number, shapeIndex: number): ChartAxisFormat[] {
  const chartPath = findChartPartPath(textFiles, slideIndex, shapeIndex);
  const chartXml = textFiles.get(chartPath);
  if (!chartXml) {
    throw new Error(`Missing chart XML part: ${chartPath}`);
  }

  const chartDoc = parseXml(chartXml, chartPath);
  const date1904 = getValAttribute(chartDoc.documentElement, 'date1904') === '1';
  const xValues = getCachedNumbers(chartDoc, ['xVal']);
  const categoryValues = getCachedNumbers(chartDoc, ['cat']);
  const yValues = getCachedNumbers(chartDoc, ['yVal']);
  const seriesValues = getCachedNumbers(chartDoc, ['val']);
  const formats: ChartAxisFormat[] = [];

  for (const axisName of ['valAx', 'dateAx']) {
    for (const axisElement of getDescendants(chartDoc, axisName)) {
      const axisPosition = getValAttribute(axisElement, 'axPos');
      const orientation: ChartAxisOrientation =
        axisPosition === 'l' || axisPosition === 'r' ? 'vertical' : 'horizontal';
      let values: number[];

      if (orientation === 'horizontal') {
        values = xValues.length > 0 ? xValues : categoryValues;
        if (values.length === 0 && axisName === 'valAx') {
          values = seriesValues;
        }
      } else {
        values = yValues.length > 0 ? yValues : seriesValues;
        if (values.length === 0 && axisName === 'dateAx') {
          values = categoryValues;
        }
      }

      const bounds = deriveAxisBounds(axisElement, values, axisName === 'valAx');
      formats.push({
        orientation,
        formatCode: getDescendants(axisElement, 'numFmt')[0]?.getAttribute('formatCode') ?? 'General',
        min: bounds.min,
        max: bounds.max,
        majorUnit: bounds.majorUnit,
        date1904
      });
    }
  }

  return formats;
}

function getDecimalPlaces(step: number): number {
  if (!Number.isFinite(step) || step <= 0) {
    return 0;
  }

  return Math.min(10, Math.max(0, Math.ceil(-Math.log10(step) - 1e-10)));
}

function formatFixedNumber(value: number, decimalPlaces: number, useThousandsSeparator: boolean): string {
  const threshold = 0.5 * 10 ** -decimalPlaces;
  const normalizedValue = Math.abs(value) < threshold ? 0 : value;
  const fixedValue = normalizedValue.toFixed(decimalPlaces);
  const [integerPart = '0', decimalPart] = fixedValue.split('.');
  const formattedInteger = useThousandsSeparator
    ? integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
    : integerPart;

  return decimalPart === undefined ? formattedInteger : `${formattedInteger}.${decimalPart}`;
}

function formatGeneralNumber(value: number, step: number): string {
  if (value !== 0 && (Math.abs(value) >= 1e12 || Math.abs(value) < 1e-8)) {
    return value.toExponential(Math.max(0, getDecimalPlaces(step))).replace('e', 'E');
  }

  return formatFixedNumber(value, getDecimalPlaces(step), false)
    .replace(/(\.\d*?)0+$/, '$1')
    .replace(/\.$/, '');
}

function excelSerialToDate(serial: number, date1904: boolean): Date {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return new Date(epoch + Math.round(serial * 24 * 60 * 60 * 1000));
}

function formatExcelDate(serial: number, formatCode: string, date1904: boolean): string {
  const date = excelSerialToDate(serial, date1904);
  const months = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ];
  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();
  const year = date.getUTCFullYear();
  const monthName = months[month - 1] ?? '';

  return (formatCode.split(';')[0] ?? formatCode)
    .replace(/\[[^\]]+\]/g, '')
    .replace(/\\(.)/g, '$1')
    .replace(/"([^"]*)"/g, '$1')
    .replace(/yyyy|yy|mmmm|mmm|mm|m|dd|d/gi, (token) => {
      switch (token.toLowerCase()) {
        case 'yyyy':
          return String(year);
        case 'yy':
          return String(year).slice(-2);
        case 'mmmm':
          return monthName;
        case 'mmm':
          return monthName.slice(0, 3);
        case 'mm':
          return String(month).padStart(2, '0');
        case 'm':
          return String(month);
        case 'dd':
          return String(day).padStart(2, '0');
        default:
          return String(day);
      }
    });
}

export function formatChartAxisValue(
  value: number,
  formatCode: string,
  step: number,
  date1904 = false
): string {
  const primaryFormat = formatCode.split(';')[0] || 'General';
  const normalizedFormat = primaryFormat.replace(/\[[^\]]+\]/g, '').replace(/"[^"]*"/g, '');

  if (/[dmy]/i.test(normalizedFormat) && !/[#0](?:\.[#0]+)?%?/i.test(normalizedFormat)) {
    return formatExcelDate(value, primaryFormat, date1904);
  }

  const isPercentage = normalizedFormat.includes('%');
  const scaledValue = isPercentage ? value * 100 : value;
  const scaledStep = isPercentage ? step * 100 : step;
  const decimalMatch = normalizedFormat.match(/\.([0#]+)/);
  const useThousandsSeparator = /[0#],[0#]/.test(normalizedFormat);
  let formattedValue: string;

  if (/E[+-]?0+/i.test(normalizedFormat)) {
    formattedValue = scaledValue
      .toExponential(decimalMatch?.[1]?.length ?? 0)
      .replace('e', 'E')
      .replace(/E(\d)/, 'E+$1');
  } else if (decimalMatch !== null || /[0#]/.test(normalizedFormat) && normalizedFormat !== 'General') {
    formattedValue = formatFixedNumber(scaledValue, decimalMatch?.[1]?.length ?? 0, useThousandsSeparator);
  } else {
    formattedValue = formatGeneralNumber(scaledValue, scaledStep);
  }

  return isPercentage ? `${formattedValue}%` : formattedValue;
}

function getChartTickRuns(chartGroup: Element): ChartTickRun[] {
  const runs: ChartTickRun[] = [];
  let currentRun: ChartTickRun | null = null;
  let currentKey: string | null = null;
  let previousPosition: number | null = null;

  const finishRun = (): void => {
    if (currentRun !== null && currentRun.elements.length >= 2) {
      runs.push(currentRun);
    }

    currentRun = null;
    currentKey = null;
    previousPosition = null;
  };

  for (const textElement of getDescendants(chartGroup, 'text') as SVGTextElement[]) {
    if (textElement.getAttribute('fill')?.toLowerCase() !== '#666666') {
      finishRun();
      continue;
    }

    const anchor = textElement.getAttribute('text-anchor');
    const x = textElement.getAttribute('x');
    const y = textElement.getAttribute('y');

    if (x === null || y === null || anchor !== 'middle' && anchor !== 'end') {
      finishRun();
      continue;
    }

    const orientation: ChartAxisOrientation = anchor === 'middle' ? 'horizontal' : 'vertical';
    const key = orientation === 'horizontal' ? `h:${y}` : `v:${x}`;
    const position = Number(orientation === 'horizontal' ? x : y);
    const hasReset =
      previousPosition !== null &&
      Number.isFinite(position) &&
      (orientation === 'horizontal' ? position <= previousPosition : position >= previousPosition);

    if (key !== currentKey || hasReset) {
      finishRun();
      currentRun = { orientation, elements: [] };
      currentKey = key;
    }

    currentRun?.elements.push(textElement);
    previousPosition = position;
  }

  finishRun();
  return runs;
}

function removeRedundantTickRuns(runs: ChartTickRun[], axis: ChartAxisFormat): ChartTickRun[] {
  if (runs.length < 2) {
    return runs;
  }

  const keptRuns: ChartTickRun[] = [];

  for (const run of runs) {
    const equivalentIndex = keptRuns.findIndex((keptRun) => {
      if (keptRun.orientation !== run.orientation) {
        return false;
      }

      const keptFirst = keptRun.elements[0];
      const keptLast = keptRun.elements[keptRun.elements.length - 1];
      const first = run.elements[0];
      const last = run.elements[run.elements.length - 1];
      const coordinate = run.orientation === 'horizontal' ? 'x' : 'y';

      if (!keptFirst || !keptLast || !first || !last) {
        return false;
      }

      return (
        keptFirst.getAttribute(coordinate) === first.getAttribute(coordinate) &&
        keptLast.getAttribute(coordinate) === last.getAttribute(coordinate)
      );
    });

    if (equivalentIndex === -1) {
      keptRuns.push(run);
      continue;
    }

    const keptRun = keptRuns[equivalentIndex];
    if (!keptRun) {
      keptRuns.push(run);
      continue;
    }

    const expectedCount =
      axis.majorUnit === null ? null : Math.round((axis.max - axis.min) / axis.majorUnit) + 1;
    const shouldReplace =
      expectedCount !== null &&
      Math.abs(run.elements.length - expectedCount) < Math.abs(keptRun.elements.length - expectedCount);
    const redundantRun = shouldReplace ? keptRun : run;

    for (const element of redundantRun.elements) {
      element.parentNode?.removeChild(element);
    }

    if (shouldReplace) {
      keptRuns[equivalentIndex] = run;
    }
  }

  return keptRuns;
}

export { findChartPartPath, getChartAxisFormats, getChartTextValues, getChartTickRuns, removeRedundantTickRuns };
