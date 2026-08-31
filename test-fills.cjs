const fs = require('fs');
const content = fs.readFileSync('src/powerpoint/chartAxisFormatting.ts', 'utf8');
const search = `  for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
    for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
      const value = seriesValues[seriesIndex]?.[categoryIndex] ?? 0;
      const length = valueToLength(value, valueAxis.min, valueAxis.max, isColumn ? plotHeight : plotWidth);
      if (length <= 0) {
        continue;
      }

      const rect = createSvgRect(parent, fill);`;

console.log(content.includes(search));
