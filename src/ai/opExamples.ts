import type { DocumentOp } from './types';

const TRANSFORM = { x: 914_400, y: 685_800, cx: 2_743_200, cy: 2_057_400, rot: 0 };

/** Minimal schema-valid example payload per operation id (for agents and tests). */
export const OP_EXAMPLES: Record<string, DocumentOp> = {
	'pptx.updateShapeText': { op: 'pptx.updateShapeText', slideIndex: 0, shapeIndex: 0, text: 'Example title' },
	'pptx.replaceShapeParagraphs': {
		op: 'pptx.replaceShapeParagraphs',
		slideIndex: 0,
		shapeIndex: 0,
		paragraphs: [
			{ text: 'Example list heading', listStyle: 'none', bold: true },
			{ text: 'A native PowerPoint bullet', listStyle: 'bullet', bold: false },
		],
	},
	'pptx.updateParagraphText': { op: 'pptx.updateParagraphText', slideIndex: 0, shapeIndex: 0, paragraphIndex: 0, text: 'Paragraph' },
	'pptx.updateTextRun': { op: 'pptx.updateTextRun', slideIndex: 0, shapeIndex: 0, paragraphIndex: 0, runIndex: 0, text: 'Run' },
	'pptx.replaceText': { op: 'pptx.replaceText', query: 'old', replacement: 'new', matchCase: false },
	'pptx.setRunStyle': { op: 'pptx.setRunStyle', slideIndex: 0, shapeIndex: 0, paragraphIndex: 0, runIndex: 0, style: { bold: true } },
	'pptx.setParagraphAlignment': { op: 'pptx.setParagraphAlignment', slideIndex: 0, shapeIndex: 0, paragraphIndex: 0, align: 'ctr' },
	'pptx.applyListStyle': { op: 'pptx.applyListStyle', slideIndex: 0, shapeIndex: 0, paragraphIndex: 0, style: 'bullet' },
	'pptx.updateTransform': { op: 'pptx.updateTransform', slideIndex: 0, shapeIndex: 0, transform: TRANSFORM },
	'pptx.setShapeFillColor': { op: 'pptx.setShapeFillColor', slideIndex: 0, shapeIndex: 0, hex: '1B75BB' },
	'pptx.reorderShapes': { op: 'pptx.reorderShapes', slideIndex: 0, shapeIndex: 0, mode: 'forward' },
	'pptx.groupShapes': { op: 'pptx.groupShapes', slideIndex: 0, shapeIndices: [0, 1] },
	'pptx.ungroupShapes': { op: 'pptx.ungroupShapes', slideIndex: 0, shapeIndex: 0 },
	'pptx.flipShape': { op: 'pptx.flipShape', slideIndex: 0, shapeIndex: 0, axis: 'horizontal' },
	'pptx.deleteShape': { op: 'pptx.deleteShape', slideIndex: 0, shapeIndex: 0 },
	'pptx.addImage': { op: 'pptx.addImage', slideIndex: 0, vaultImagePath: 'assets/example.png', transform: TRANSFORM },
	'pptx.addShape': { op: 'pptx.addShape', slideIndex: 0, geometry: 'rect', transform: TRANSFORM },
	'pptx.addTextBox': { op: 'pptx.addTextBox', slideIndex: 0, transform: TRANSFORM },
	'pptx.addTable': { op: 'pptx.addTable', slideIndex: 0, rows: 2, cols: 2, transform: TRANSFORM },
	'pptx.addChart': { op: 'pptx.addChart', slideIndex: 0, transform: TRANSFORM },
	'pptx.addSlide': { op: 'pptx.addSlide', afterIndex: 0, layout: 'blank' },
	'pptx.deleteSlide': { op: 'pptx.deleteSlide', slideIndex: 1 },
	'pptx.moveSlide': { op: 'pptx.moveSlide', slideIndex: 1, direction: -1 },
	'pptx.duplicateSlide': { op: 'pptx.duplicateSlide', slideIndex: 0 },
	'pptx.reorderSlides': { op: 'pptx.reorderSlides', order: [1, 0] },
	'pptx.setSlideBackground': { op: 'pptx.setSlideBackground', slideIndex: 0, colorHex: 'FFFFFF' },
	'pptx.setImageCrop': { op: 'pptx.setImageCrop', slideIndex: 0, shapeIndex: 0, crop: { left: 10, top: 10, right: 10, bottom: 10 } },
	'pptx.fitImageToFrame': { op: 'pptx.fitImageToFrame', slideIndex: 0, shapeIndex: 0 },
	'pptx.resetImage': { op: 'pptx.resetImage', slideIndex: 0, shapeIndex: 0 },
	'pptx.replaceImage': { op: 'pptx.replaceImage', slideIndex: 0, shapeIndex: 0, vaultImagePath: 'assets/example.png' },
	'pptx.replaceImageFromShape': { op: 'pptx.replaceImageFromShape', slideIndex: 0, shapeIndex: 0, sourceSlideIndex: 0, sourceShapeIndex: 1 },
	'pptx.updateChartData': {
		op: 'pptx.updateChartData',
		slideIndex: 0,
		shapeIndex: 0,
		data: { categories: ['A', 'B'], series: [{ name: 'S1', values: [1, 2] }] },
	},
	'docx.removeComments': { op: 'docx.removeComments' },
	'docx.setCoreProperties': { op: 'docx.setCoreProperties', creator: 'Document Author', lastModifiedBy: 'Document Editor' },
	'docx.setRunText': { op: 'docx.setRunText', blockId: 'body/p[0]', runId: 'body/p[0]/r[0]', text: 'Updated' },
	'docx.setRunStyle': { op: 'docx.setRunStyle', runId: 'body/p[0]/r[0]', style: { bold: true } },
	'docx.setParagraphStyle': { op: 'docx.setParagraphStyle', blockId: 'body/p[0]', style: { name: 'Heading1' } },
	'docx.setParagraphDefaultRunStyle': {
		op: 'docx.setParagraphDefaultRunStyle',
		blockId: 'body/p[0]',
		style: { fontSizePt: 12 },
	},
	'docx.setParagraphLayout': {
		op: 'docx.setParagraphLayout',
		blockId: 'body/p[0]',
		layout: { spacing: { before: 0, after: 0, line: 240, lineRule: 'auto' }, alignment: 'left' },
	},
	'docx.setSectionLayout': {
		op: 'docx.setSectionLayout',
		sectionIndex: 0,
		layout: { pageSize: { width: 12240, height: 15840, orient: 'portrait' }, margins: { top: 432, right: 720, bottom: 432, left: 720 } },
	},
	'docx.setParagraphBottomBorder': {
		op: 'docx.setParagraphBottomBorder',
		blockId: 'body/p[0]',
		border: { style: 'single', size: 8, color: '000000' },
	},
	'docx.insertTable': { op: 'docx.insertTable', afterBlockId: 'body/p[0]', rows: 2, cols: 2 },
	'docx.setCellText': { op: 'docx.setCellText', cellId: 'body/tbl[0]/tr[0]/tc[0]', text: 'Cell' },
	'docx.setCellStyle': { op: 'docx.setCellStyle', cellId: 'body/tbl[0]/tr[0]/tc[0]', style: { name: 'Normal' } },
	'docx.deleteTable': { op: 'docx.deleteTable', tableId: 'body/tbl[0]' },
	'docx.insertImage': { op: 'docx.insertImage', afterBlockId: 'body/p[0]', vaultImagePath: 'assets/example.png' },
	'docx.replaceImage': { op: 'docx.replaceImage', blockId: 'body/p[1]', vaultImagePath: 'assets/example.png' },
	'docx.replaceText': { op: 'docx.replaceText', query: 'old', replacement: 'new', wholeWord: false },
	'docx.insertText': { op: 'docx.insertText', blockId: 'body/p[0]', runId: 'body/p[0]/r[0]', offset: 5, text: ' inserted' },
	'docx.deleteRange': {
		op: 'docx.deleteRange',
		range: {
			start: { blockId: 'body/p[0]', offset: 0 },
			end: { blockId: 'body/p[0]', offset: 5 },
		},
	},
	'docx.deleteBlock': { op: 'docx.deleteBlock', blockId: 'body/p[1]' },
	'docx.insertHyperlink': {
		op: 'docx.insertHyperlink',
		range: {
			start: { blockId: 'body/p[0]', offset: 0 },
			end: { blockId: 'body/p[0]', offset: 5 },
		},
		url: 'https://example.com',
		displayText: 'Example',
	},
	'docx.removeHyperlink': {
		op: 'docx.removeHyperlink',
		range: {
			start: { blockId: 'body/p[0]', offset: 0 },
			end: { blockId: 'body/p[0]', offset: 5 },
		},
	},
	'docx.insertParagraphBreak': { op: 'docx.insertParagraphBreak', blockId: 'body/p[0]', offset: 5 },
	'docx.replaceBodyParagraphs': {
		op: 'docx.replaceBodyParagraphs',
		paragraphs: [
			'Marwan Luay',
			'',
			'Thank you for the scholarship.',
		],
	},
};

export function getOpExample(opId: string): DocumentOp | undefined {
	return OP_EXAMPLES[opId];
}
