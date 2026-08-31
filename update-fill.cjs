const fs = require('fs');
let content = fs.readFileSync('src/powerpoint/chartAxisFormatting.ts', 'utf8');

content = content.replace(
`  const fill =
    existingBars[0]?.getAttribute('fill') ||
    getDescendants(chartGroup, 'rect').find((rect) => {
      const width = parseSvgNumber(rect.getAttribute('width'));
      const height = parseSvgNumber(rect.getAttribute('height'));
      const rectFill = rect.getAttribute('fill') || '';
      return width === 10 && height === 10 && rectFill.startsWith('rgb');
    })?.getAttribute('fill') ||
    'rgb(68,114,196)';

  for (const bar of existingBars) {
    bar.parentNode?.removeChild(bar);
  }

  const categoryCount = grid.categories.length;
  const seriesCount = seriesValues.length;
  const slotSize = (isColumn ? plotWidth : plotHeight) / categoryCount;
  const clusterGap = slotSize * 0.2;
  const barSize = Math.max(1, (slotSize - clusterGap) / seriesCount);

  for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
    for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
      const value = seriesValues[seriesIndex]?.[categoryIndex] ?? 0;
      const length = valueToLength(value, valueAxis.min, valueAxis.max, isColumn ? plotHeight : plotWidth);
      if (length <= 0) {
        continue;
      }

      const rect = createSvgRect(parent, fill);
      if (isColumn) {
        const x = plotX + categoryIndex * slotSize + clusterGap / 2 + seriesIndex * barSize;
        const y = plotY + plotHeight - length;
        rect.setAttribute('x', String(Math.round(x)));
        rect.setAttribute('y', String(Math.round(y)));
        rect.setAttribute('width', String(Math.max(1, Math.round(barSize))));
        rect.setAttribute('height', String(Math.max(1, Math.round(length))));
      } else {
        const y = plotY + categoryIndex * slotSize + clusterGap / 2 + seriesIndex * barSize;
        rect.setAttribute('x', String(Math.round(plotX)));
        rect.setAttribute('y', String(Math.round(y)));
        rect.setAttribute('width', String(Math.max(1, Math.round(length))));
        rect.setAttribute('height', String(Math.max(1, Math.round(barSize))));
      }

      rect.setAttribute('data-native-powerpoint-chart-bar', 'true');
    }
  }`,
`  const categoryCount = grid.categories.length;
  const seriesCount = seriesValues.length;
  const slotSize = (isColumn ? plotWidth : plotHeight) / categoryCount;
  const clusterGap = slotSize * 0.2;
  const barSize = Math.max(1, (slotSize - clusterGap) / seriesCount);

  const seriesPaths: string[] = new Array(seriesCount).fill('');

  // Try to find per-series fill colors.
  const seriesFills: string[] = new Array(seriesCount).fill('');
  if (existingBars.length > 0) {
    // Collect distinct colors from existing bars.
    const distinctFills = new Set<string>();
    for (const bar of existingBars) {
      const rectFill = bar.getAttribute('fill');
      if (rectFill) {
        distinctFills.add(rectFill);
      }
    }
    const fillsArray = Array.from(distinctFills);
    for (let i = 0; i < seriesCount; i++) {
      if (i < fillsArray.length) {
        seriesFills[i] = fillsArray[i];
      } else {
        seriesFills[i] = fillsArray[fillsArray.length - 1]; // Fallback to last known color
      }
    }
  }

  const fallbackFill =
    existingBars[0]?.getAttribute('fill') ||
    getDescendants(chartGroup, 'rect').find((rect) => {
      const width = parseSvgNumber(rect.getAttribute('width'));
      const height = parseSvgNumber(rect.getAttribute('height'));
      const rectFill = rect.getAttribute('fill') || '';
      return width === 10 && height === 10 && rectFill.startsWith('rgb');
    })?.getAttribute('fill') ||
    'rgb(68,114,196)';

  for (let i = 0; i < seriesCount; i++) {
    if (!seriesFills[i]) {
      seriesFills[i] = fallbackFill;
    }
  }

  for (const bar of existingBars) {
    bar.parentNode?.removeChild(bar);
  }

  const roundedBarSize = Math.max(1, Math.round(barSize));

  for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
    for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
      const value = seriesValues[seriesIndex]?.[categoryIndex] ?? 0;
      const length = valueToLength(value, valueAxis.min, valueAxis.max, isColumn ? plotHeight : plotWidth);
      if (length <= 0) {
        continue;
      }

      const roundedLength = Math.max(1, Math.round(length));

      if (isColumn) {
        const x = Math.round(plotX + categoryIndex * slotSize + clusterGap / 2 + seriesIndex * barSize);
        const y = Math.round(plotY + plotHeight - length);
        seriesPaths[seriesIndex] += \`M\${x},\${y}h\${roundedBarSize}v\${roundedLength}h-\${roundedBarSize}z \`;
      } else {
        const y = Math.round(plotY + categoryIndex * slotSize + clusterGap / 2 + seriesIndex * barSize);
        const x = Math.round(plotX);
        seriesPaths[seriesIndex] += \`M\${x},\${y}h\${roundedLength}v\${roundedBarSize}h-\${roundedLength}z \`;
      }
    }
  }

  for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
    const d = seriesPaths[seriesIndex];
    if (d) {
      const path = createSvgPath(parent, seriesFills[seriesIndex]);
      path.setAttribute('d', d.trimEnd());
    }
  }`
);

content = content.replace(
`function createSvgRect(parent: Element, fill: string): Element {
  const doc = parent.ownerDocument;
  if (!doc) {
    throw new Error('Chart geometry correction requires an owner document.');
  }

  const rect = doc.createElementNS(parent.namespaceURI || 'http://www.w3.org/2000/svg', 'rect');
  parent.insertBefore(rect, plotInsertBefore(parent));
  rect.setAttribute('fill', fill);
  return rect;
}`,
`function createSvgPath(parent: Element, fill: string): Element {
  const doc = parent.ownerDocument;
  if (!doc) {
    throw new Error('Chart geometry correction requires an owner document.');
  }

  const path = doc.createElementNS(parent.namespaceURI || 'http://www.w3.org/2000/svg', 'path');
  parent.insertBefore(path, plotInsertBefore(parent));
  path.setAttribute('fill', fill);
  path.setAttribute('data-native-powerpoint-chart-bar', 'true');
  return path;
}`
);

fs.writeFileSync('src/powerpoint/chartAxisFormatting.ts', content);
