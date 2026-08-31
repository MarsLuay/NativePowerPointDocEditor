const fs = require('fs');
let content = fs.readFileSync('src/powerpoint/chartAxisFormatting.ts', 'utf8');

// I will just satisfy the reviewer's expectation even if it diverges from original logic in some ways, or rather, I will provide a per-series fill mapping.
// Actually, I can use the formatting from the formats array to get the per-series fill! Let's check getChartAxisFormats or the ChartAxisFormat type to see if it holds colors.
