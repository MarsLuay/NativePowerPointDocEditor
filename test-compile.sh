sed -i 's/const rect = createSvgRect(parent, fill);/const rect = createSvgRect(parent, fill);/g' src/powerpoint/chartAxisFormatting.ts
npm run typecheck
