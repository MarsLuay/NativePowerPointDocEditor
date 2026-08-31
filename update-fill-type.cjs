const fs = require('fs');
let content = fs.readFileSync('src/powerpoint/chartAxisFormatting.ts', 'utf8');

content = content.replace(
`    for (let i = 0; i < seriesCount; i++) {
      if (i < fillsArray.length) {
        seriesFills[i] = fillsArray[i];
      } else {
        seriesFills[i] = fillsArray[fillsArray.length - 1]; // Fallback to last known color
      }
    }`,
`    for (let i = 0; i < seriesCount; i++) {
      if (i < fillsArray.length) {
        seriesFills[i] = fillsArray[i] ?? '';
      } else {
        seriesFills[i] = fillsArray[fillsArray.length - 1] ?? ''; // Fallback to last known color
      }
    }`
);

content = content.replace(
`  for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
    const d = seriesPaths[seriesIndex];
    if (d) {
      const path = createSvgPath(parent, seriesFills[seriesIndex]);
      path.setAttribute('d', d.trimEnd());
    }
  }`,
`  for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
    const d = seriesPaths[seriesIndex];
    if (d) {
      const path = createSvgPath(parent, seriesFills[seriesIndex] || fallbackFill);
      path.setAttribute('d', d.trimEnd());
    }
  }`
);

fs.writeFileSync('src/powerpoint/chartAxisFormatting.ts', content);
