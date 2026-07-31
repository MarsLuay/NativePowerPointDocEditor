/**
 * Insertable chart types + OOXML part builders for the Insert Chart picker.
 * Classic charts use `c:chartSpace`; modern types use chartEx (`cx:chartSpace`).
 */

import {
  CHART_INSERT_CHART_RELS_XML,
  CHART_INSERT_WORKBOOK_BASE64,
} from '../chartInsertTemplate';

export const CLASSIC_CHART_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const CHARTEX_NAMESPACE = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
export const CLASSIC_CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const CHARTEX_CONTENT_TYPE = 'application/vnd.ms-office.chartex+xml';
export const CLASSIC_CHART_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
export const CHARTEX_RELATIONSHIP_TYPE =
  'http://schemas.microsoft.com/office/2014/relationships/chartEx';

export type InsertableChartType =
  | 'column'
  | 'line'
  | 'pie'
  | 'bar'
  | 'area'
  | 'scatter'
  | 'stock'
  | 'surface'
  | 'radar'
  | 'treemap'
  | 'sunburst'
  | 'histogram'
  | 'boxWhisker'
  | 'waterfall'
  | 'combo';

export interface ChartTypeMenuEntry {
  readonly id: InsertableChartType;
  /** Locale suffix under powerpoint:toolbar.chartType.* */
  readonly labelKey: string;
  readonly icon: string;
  readonly kind: 'classic' | 'chartex';
  /** Expected plot element localName for tests (c: or cx: series layoutId). */
  readonly fingerprint: string;
}

export interface ChartTemplateMenuEntry {
  readonly id: string;
  readonly labelKey: string;
  readonly icon: string;
  readonly chartType: InsertableChartType;
}

export interface ChartInsertParts {
  readonly chartType: InsertableChartType;
  readonly kind: 'classic' | 'chartex';
  readonly frameXml: string;
  readonly chartXml: string;
  readonly chartRelsXml: string;
  readonly workbookBase64: string;
  readonly contentType: string;
  readonly relationshipType: string;
  readonly graphicUri: string;
}

const CAT_AX = 12345678;
const VAL_AX = 87654321;
const VAL_AX_2 = 87654322;

const SERIES_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000'] as const;

export const INSERTABLE_CHART_TYPES: readonly ChartTypeMenuEntry[] = [
  { id: 'column', labelKey: 'column', icon: 'bar-chart-3', kind: 'classic', fingerprint: 'barChart' },
  { id: 'line', labelKey: 'line', icon: 'trending-up', kind: 'classic', fingerprint: 'lineChart' },
  { id: 'pie', labelKey: 'pie', icon: 'pie-chart', kind: 'classic', fingerprint: 'pieChart' },
  { id: 'bar', labelKey: 'bar', icon: 'bar-chart-horizontal', kind: 'classic', fingerprint: 'barChart' },
  { id: 'area', labelKey: 'area', icon: 'area-chart', kind: 'classic', fingerprint: 'areaChart' },
  { id: 'scatter', labelKey: 'scatter', icon: 'circle-dot', kind: 'classic', fingerprint: 'scatterChart' },
  { id: 'stock', labelKey: 'stock', icon: 'activity', kind: 'classic', fingerprint: 'stockChart' },
  { id: 'surface', labelKey: 'surface', icon: 'box', kind: 'classic', fingerprint: 'surfaceChart' },
  { id: 'radar', labelKey: 'radar', icon: 'hexagon', kind: 'classic', fingerprint: 'radarChart' },
  { id: 'treemap', labelKey: 'treemap', icon: 'layout-grid', kind: 'chartex', fingerprint: 'treemap' },
  { id: 'sunburst', labelKey: 'sunburst', icon: 'sun', kind: 'chartex', fingerprint: 'sunburst' },
  { id: 'histogram', labelKey: 'histogram', icon: 'bar-chart-big', kind: 'chartex', fingerprint: 'histogram' },
  { id: 'boxWhisker', labelKey: 'boxWhisker', icon: 'separator-vertical', kind: 'chartex', fingerprint: 'boxWhisker' },
  { id: 'waterfall', labelKey: 'waterfall', icon: 'stairs', kind: 'chartex', fingerprint: 'waterfall' },
  { id: 'combo', labelKey: 'combo', icon: 'layers', kind: 'classic', fingerprint: 'barChart' },
] as const;

export const CHART_TEMPLATE_ENTRIES: readonly ChartTemplateMenuEntry[] = [
  { id: 'template-column', labelKey: 'templateColumn', icon: 'bar-chart-3', chartType: 'column' },
  { id: 'template-line', labelKey: 'templateLine', icon: 'trending-up', chartType: 'line' },
  { id: 'template-pie', labelKey: 'templatePie', icon: 'pie-chart', chartType: 'pie' },
] as const;

const CHART_TYPE_IDS = new Set<string>(INSERTABLE_CHART_TYPES.map((entry) => entry.id));

export function isInsertableChartType(value: unknown): value is InsertableChartType {
  return typeof value === 'string' && CHART_TYPE_IDS.has(value);
}

export function normalizeInsertableChartType(value: unknown): InsertableChartType {
  return isInsertableChartType(value) ? value : 'column';
}

export function getChartTypeMenuEntry(type: InsertableChartType): ChartTypeMenuEntry {
  const entry = INSERTABLE_CHART_TYPES.find((candidate) => candidate.id === type);
  if (!entry) {
    throw new Error(`Unknown chart type: ${type}`);
  }
  return entry;
}

const RECENT_CHART_TYPES_KEY = 'npde-recent-chart-types';
const RECENT_CHART_TYPES_LIMIT = 5;

export function readRecentChartTypes(): InsertableChartType[] {
  try {
    const raw = window.localStorage?.getItem(RECENT_CHART_TYPES_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(isInsertableChartType).slice(0, RECENT_CHART_TYPES_LIMIT);
  } catch {
    return [];
  }
}

export function rememberRecentChartType(type: InsertableChartType): InsertableChartType[] {
  const next = [type, ...readRecentChartTypes().filter((entry) => entry !== type)]
    .slice(0, RECENT_CHART_TYPES_LIMIT);
  try {
    window.localStorage?.setItem(RECENT_CHART_TYPES_KEY, JSON.stringify(next));
  } catch {
    // Ignore quota / private-mode failures; insert still succeeds.
  }
  return next;
}

function strCache(values: string[]): string {
  return [
    '<c:strCache>',
    `<c:ptCount val="${values.length}"/>`,
    ...values.map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`),
    '</c:strCache>',
  ].join('');
}

function numCache(values: number[]): string {
  return [
    '<c:numCache>',
    '<c:formatCode>General</c:formatCode>',
    `<c:ptCount val="${values.length}"/>`,
    ...values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`),
    '</c:numCache>',
  ].join('');
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function solidFill(color: string): string {
  return `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
}

function catRef(formula: string, values: string[]): string {
  return `<c:cat><c:strRef><c:f>${formula}</c:f>${strCache(values)}</c:strRef></c:cat>`;
}

function valRef(formula: string, values: number[]): string {
  return `<c:val><c:numRef><c:f>${formula}</c:f>${numCache(values)}</c:numRef></c:val>`;
}

function txRef(formula: string, name: string): string {
  return `<c:tx><c:strRef><c:f>${formula}</c:f>${strCache([name])}</c:strRef></c:tx>`;
}

function classicSeries(options: {
  index: number;
  name: string;
  nameFormula: string;
  categories: string[];
  catFormula: string;
  values: number[];
  valFormula: string;
  color?: string;
}): string {
  const color = options.color ?? SERIES_COLORS[options.index % SERIES_COLORS.length]!;
  return [
    '<c:ser>',
    `<c:idx val="${options.index}"/>`,
    `<c:order val="${options.index}"/>`,
    txRef(options.nameFormula, options.name),
    solidFill(color),
    catRef(options.catFormula, options.categories),
    valRef(options.valFormula, options.values),
    '</c:ser>',
  ].join('');
}

function scatterSeries(options: {
  index: number;
  name: string;
  nameFormula: string;
  xValues: number[];
  xFormula: string;
  yValues: number[];
  yFormula: string;
}): string {
  const color = SERIES_COLORS[options.index % SERIES_COLORS.length]!;
  return [
    '<c:ser>',
    `<c:idx val="${options.index}"/>`,
    `<c:order val="${options.index}"/>`,
    txRef(options.nameFormula, options.name),
    solidFill(color),
    `<c:xVal><c:numRef><c:f>${options.xFormula}</c:f>${numCache(options.xValues)}</c:numRef></c:xVal>`,
    `<c:yVal><c:numRef><c:f>${options.yFormula}</c:f>${numCache(options.yValues)}</c:numRef></c:yVal>`,
    '</c:ser>',
  ].join('');
}

function dataLabels(): string {
  return [
    '<c:dLbls>',
    '<c:showLegendKey val="0"/>',
    '<c:showVal val="0"/>',
    '<c:showCatName val="0"/>',
    '<c:showSerName val="0"/>',
    '<c:showPercent val="0"/>',
    '<c:showBubbleSize val="0"/>',
    '</c:dLbls>',
  ].join('');
}

function catAx(axId: number, crossAx: number, axPos: 'l' | 'r' | 'b' | 't' = 'b'): string {
  return [
    '<c:catAx>',
    `<c:axId val="${axId}"/>`,
    '<c:scaling><c:orientation val="minMax"/></c:scaling>',
    '<c:delete val="0"/>',
    `<c:axPos val="${axPos}"/>`,
    '<c:numFmt formatCode="General" sourceLinked="1"/>',
    '<c:majorTickMark val="out"/>',
    '<c:minorTickMark val="none"/>',
    '<c:tickLblPos val="nextTo"/>',
    `<c:crossAx val="${crossAx}"/>`,
    '<c:crosses val="autoZero"/>',
    '<c:auto val="1"/>',
    '<c:lblAlgn val="ctr"/>',
    '<c:lblOffset val="100"/>',
    '<c:noMultiLvlLbl val="0"/>',
    '</c:catAx>',
  ].join('');
}

function valAx(axId: number, crossAx: number, axPos: 'l' | 'r' | 'b' | 't' = 'l'): string {
  return [
    '<c:valAx>',
    `<c:axId val="${axId}"/>`,
    '<c:scaling><c:orientation val="minMax"/></c:scaling>',
    '<c:delete val="0"/>',
    `<c:axPos val="${axPos}"/>`,
    '<c:numFmt formatCode="General" sourceLinked="1"/>',
    '<c:majorTickMark val="out"/>',
    '<c:minorTickMark val="none"/>',
    '<c:tickLblPos val="nextTo"/>',
    `<c:crossAx val="${crossAx}"/>`,
    '<c:crosses val="autoZero"/>',
    '<c:crossBetween val="between"/>',
    '</c:valAx>',
  ].join('');
}

function serAx(axId: number, crossAx: number): string {
  return [
    '<c:serAx>',
    `<c:axId val="${axId}"/>`,
    '<c:scaling><c:orientation val="minMax"/></c:scaling>',
    '<c:delete val="0"/>',
    '<c:axPos val="b"/>',
    '<c:majorTickMark val="out"/>',
    '<c:minorTickMark val="none"/>',
    '<c:tickLblPos val="nextTo"/>',
    `<c:crossAx val="${crossAx}"/>`,
    '<c:crosses val="autoZero"/>',
    '</c:serAx>',
  ].join('');
}

function wrapClassicChartSpace(plotInner: string): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<c:chartSpace xmlns:c="${CLASSIC_CHART_NAMESPACE}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`,
    '<c:lang val="en-US"/>',
    '<c:roundedCorners val="0"/>',
    '<c:chart>',
    '<c:plotArea>',
    '<c:layout/>',
    plotInner,
    '</c:plotArea>',
    '<c:legend><c:legendPos val="r"/><c:layout/></c:legend>',
    '<c:plotVisOnly val="1"/>',
    '<c:dispBlanksAs val="gap"/>',
    '<c:showDLblsOverMax val="0"/>',
    '</c:chart>',
    '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>',
    '</c:chartSpace>',
  ].join('');
}

const SAMPLE_CATS = ['1g', '2g', '3g'];
const SAMPLE_VALS_A = [1.04, 2.08, 1.56];
const SAMPLE_VALS_B = [2.2, 1.7, 2.9];

function buildClassicPlotXml(type: InsertableChartType): string {
  const seriesA = classicSeries({
    index: 0,
    name: 'Samples',
    nameFormula: 'Sheet1!$B$1',
    categories: SAMPLE_CATS,
    catFormula: 'Sheet1!$A$2:$A$4',
    values: SAMPLE_VALS_A,
    valFormula: 'Sheet1!$B$2:$B$4',
  });
  const seriesB = classicSeries({
    index: 1,
    name: 'Series 2',
    nameFormula: 'Sheet1!$C$1',
    categories: SAMPLE_CATS,
    catFormula: 'Sheet1!$A$2:$A$4',
    values: SAMPLE_VALS_B,
    valFormula: 'Sheet1!$C$2:$C$4',
    color: SERIES_COLORS[1],
  });

  switch (type) {
    case 'column':
      return [
        '<c:barChart>',
        '<c:barDir val="col"/>',
        '<c:grouping val="clustered"/>',
        seriesA,
        dataLabels(),
        '<c:gapWidth val="150"/>',
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX}"/>`,
        '</c:barChart>',
        catAx(CAT_AX, VAL_AX),
        valAx(VAL_AX, CAT_AX),
      ].join('');
    case 'bar':
      return [
        '<c:barChart>',
        '<c:barDir val="bar"/>',
        '<c:grouping val="clustered"/>',
        seriesA,
        dataLabels(),
        '<c:gapWidth val="150"/>',
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX}"/>`,
        '</c:barChart>',
        catAx(CAT_AX, VAL_AX, 'l'),
        valAx(VAL_AX, CAT_AX, 'b'),
      ].join('');
    case 'line':
      return [
        '<c:lineChart>',
        '<c:grouping val="standard"/>',
        seriesA,
        dataLabels(),
        '<c:marker val="1"/>',
        '<c:smooth val="0"/>',
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX}"/>`,
        '</c:lineChart>',
        catAx(CAT_AX, VAL_AX),
        valAx(VAL_AX, CAT_AX),
      ].join('');
    case 'area':
      return [
        '<c:areaChart>',
        '<c:grouping val="standard"/>',
        seriesA,
        dataLabels(),
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX}"/>`,
        '</c:areaChart>',
        catAx(CAT_AX, VAL_AX),
        valAx(VAL_AX, CAT_AX),
      ].join('');
    case 'pie':
      return [
        '<c:pieChart>',
        '<c:varyColors val="1"/>',
        seriesA,
        dataLabels(),
        '<c:firstSliceAng val="0"/>',
        '</c:pieChart>',
      ].join('');
    case 'radar':
      return [
        '<c:radarChart>',
        '<c:radarStyle val="marker"/>',
        '<c:varyColors val="0"/>',
        seriesA,
        dataLabels(),
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX}"/>`,
        '</c:radarChart>',
        catAx(CAT_AX, VAL_AX),
        valAx(VAL_AX, CAT_AX),
      ].join('');
    case 'scatter':
      return [
        '<c:scatterChart>',
        '<c:scatterStyle val="marker"/>',
        '<c:varyColors val="0"/>',
        scatterSeries({
          index: 0,
          name: 'Samples',
          nameFormula: 'Sheet1!$B$1',
          xValues: [1, 2, 3],
          xFormula: 'Sheet1!$A$2:$A$4',
          yValues: SAMPLE_VALS_A,
          yFormula: 'Sheet1!$B$2:$B$4',
        }),
        dataLabels(),
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX}"/>`,
        '</c:scatterChart>',
        valAx(CAT_AX, VAL_AX, 'b'),
        valAx(VAL_AX, CAT_AX),
      ].join('');
    case 'stock': {
      const high = classicSeries({
        index: 0,
        name: 'High',
        nameFormula: 'Sheet1!$B$1',
        categories: SAMPLE_CATS,
        catFormula: 'Sheet1!$A$2:$A$4',
        values: [2.5, 3.1, 2.8],
        valFormula: 'Sheet1!$B$2:$B$4',
        color: SERIES_COLORS[0],
      });
      const low = classicSeries({
        index: 1,
        name: 'Low',
        nameFormula: 'Sheet1!$C$1',
        categories: SAMPLE_CATS,
        catFormula: 'Sheet1!$A$2:$A$4',
        values: [1.0, 1.5, 1.2],
        valFormula: 'Sheet1!$C$2:$C$4',
        color: SERIES_COLORS[1],
      });
      const close = classicSeries({
        index: 2,
        name: 'Close',
        nameFormula: 'Sheet1!$D$1',
        categories: SAMPLE_CATS,
        catFormula: 'Sheet1!$A$2:$A$4',
        values: [2.1, 2.7, 2.0],
        valFormula: 'Sheet1!$D$2:$D$4',
        color: SERIES_COLORS[2],
      });
      return [
        '<c:stockChart>',
        high,
        low,
        close,
        '<c:hiLowLines/>',
        dataLabels(),
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX}"/>`,
        '</c:stockChart>',
        catAx(CAT_AX, VAL_AX),
        valAx(VAL_AX, CAT_AX),
      ].join('');
    }
    case 'surface':
      return [
        '<c:surfaceChart>',
        '<c:wireframe val="0"/>',
        seriesA,
        seriesB,
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX}"/>`,
        `<c:axId val="${VAL_AX_2}"/>`,
        '</c:surfaceChart>',
        catAx(CAT_AX, VAL_AX),
        valAx(VAL_AX, CAT_AX),
        serAx(VAL_AX_2, VAL_AX),
      ].join('');
    case 'combo':
      return [
        '<c:barChart>',
        '<c:barDir val="col"/>',
        '<c:grouping val="clustered"/>',
        seriesA,
        dataLabels(),
        '<c:gapWidth val="150"/>',
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX}"/>`,
        '</c:barChart>',
        '<c:lineChart>',
        '<c:grouping val="standard"/>',
        seriesB,
        dataLabels(),
        '<c:marker val="1"/>',
        '<c:smooth val="0"/>',
        `<c:axId val="${CAT_AX}"/>`,
        `<c:axId val="${VAL_AX_2}"/>`,
        '</c:lineChart>',
        catAx(CAT_AX, VAL_AX),
        valAx(VAL_AX, CAT_AX),
        valAx(VAL_AX_2, CAT_AX, 'r'),
      ].join('');
    default:
      throw new Error(`Classic builder does not support chart type: ${type}`);
  }
}

function cxPts(values: Array<string | number>): string {
  return values
    .map((value, index) => `<cx:pt idx="${index}">${escapeXml(String(value))}</cx:pt>`)
    .join('');
}

function buildChartExXml(layoutId: string, categories: string[], values: number[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<cx:chartSpace xmlns:cx="${CHARTEX_NAMESPACE}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`,
    '<cx:chartData>',
    '<cx:data id="0">',
    '<cx:strDim type="cat">',
    `<cx:lvl>${cxPts(categories)}</cx:lvl>`,
    '</cx:strDim>',
    '<cx:numDim type="val">',
    `<cx:lvl>${cxPts(values)}</cx:lvl>`,
    '</cx:numDim>',
    '</cx:data>',
    '</cx:chartData>',
    '<cx:chart>',
    '<cx:plotArea>',
    '<cx:plotAreaRegion>',
    `<cx:series layoutId="${layoutId}">`,
    '<cx:dataId val="0"/>',
    layoutId === 'waterfall'
      ? '<cx:layoutPr><cx:subtotals><cx:idx val="3"/></cx:subtotals></cx:layoutPr>'
      : '',
    '</cx:series>',
    '</cx:plotAreaRegion>',
    '<cx:axis id="0"><cx:catScaling/></cx:axis>',
    '<cx:axis id="1"><cx:valScaling/></cx:axis>',
    '</cx:plotArea>',
    '<cx:legend pos="r"/>',
    '</cx:chart>',
    '</cx:chartSpace>',
  ].join('');
}

function buildChartExXmlForType(type: InsertableChartType): string {
  switch (type) {
    case 'treemap':
      return buildChartExXml('treemap', ['Alpha', 'Beta', 'Gamma', 'Delta'], [40, 30, 20, 10]);
    case 'sunburst':
      return buildChartExXml('sunburst', ['A', 'B', 'C', 'D'], [40, 30, 20, 10]);
    case 'histogram':
      return buildChartExXml('histogram', ['0-1', '1-2', '2-3', '3-4'], [3, 7, 5, 2]);
    case 'boxWhisker':
      return buildChartExXml('boxWhisker', ['Sample'], [1, 2, 3, 4, 5, 6, 7, 8, 9]);
    case 'waterfall':
      return buildChartExXml('waterfall', ['Start', 'Q1', 'Q2', 'Total'], [100, 20, -15, 105]);
    default:
      throw new Error(`ChartEx builder does not support chart type: ${type}`);
  }
}

function buildFrameXml(kind: 'classic' | 'chartex'): string {
  if (kind === 'classic') {
    return [
      '<p:graphicFrame>',
      '<p:nvGraphicFramePr>',
      '<p:cNvPr id="2" name="Chart"/>',
      '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>',
      '<p:nvPr/>',
      '</p:nvGraphicFramePr>',
      '<p:xfrm>',
      '<a:off x="914400" y="685800"/>',
      '<a:ext cx="7315200" cy="3657600"/>',
      '</p:xfrm>',
      '<a:graphic>',
      `<a:graphicData uri="${CLASSIC_CHART_NAMESPACE}">`,
      `<c:chart xmlns:c="${CLASSIC_CHART_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>`,
      '</a:graphicData>',
      '</a:graphic>',
      '</p:graphicFrame>',
    ].join('');
  }

  return [
    '<p:graphicFrame>',
    '<p:nvGraphicFramePr>',
    '<p:cNvPr id="2" name="Chart"/>',
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>',
    '<p:nvPr/>',
    '</p:nvGraphicFramePr>',
    '<p:xfrm>',
    '<a:off x="914400" y="685800"/>',
    '<a:ext cx="7315200" cy="3657600"/>',
    '</p:xfrm>',
    '<a:graphic>',
    `<a:graphicData uri="${CHARTEX_NAMESPACE}">`,
    `<cx:chart xmlns:cx="${CHARTEX_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>`,
    '</a:graphicData>',
    '</a:graphic>',
    '</p:graphicFrame>',
  ].join('');
}

export function buildChartInsertParts(type: InsertableChartType): ChartInsertParts {
  const entry = getChartTypeMenuEntry(type);
  if (entry.kind === 'classic') {
    return {
      chartType: type,
      kind: 'classic',
      frameXml: buildFrameXml('classic'),
      chartXml: wrapClassicChartSpace(buildClassicPlotXml(type)),
      chartRelsXml: CHART_INSERT_CHART_RELS_XML,
      workbookBase64: CHART_INSERT_WORKBOOK_BASE64,
      contentType: CLASSIC_CHART_CONTENT_TYPE,
      relationshipType: CLASSIC_CHART_RELATIONSHIP_TYPE,
      graphicUri: CLASSIC_CHART_NAMESPACE,
    };
  }

  return {
    chartType: type,
    kind: 'chartex',
    frameXml: buildFrameXml('chartex'),
    chartXml: buildChartExXmlForType(type),
    chartRelsXml: CHART_INSERT_CHART_RELS_XML,
    workbookBase64: CHART_INSERT_WORKBOOK_BASE64,
    contentType: CHARTEX_CONTENT_TYPE,
    relationshipType: CHARTEX_RELATIONSHIP_TYPE,
    graphicUri: CHARTEX_NAMESPACE,
  };
}
