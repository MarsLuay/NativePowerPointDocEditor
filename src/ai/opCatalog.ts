import type { JsonSchema, OpDefinition } from './types';

const SLIDE_INDEX: JsonSchema = {
	type: 'integer',
	minimum: 0,
	description: '0-based slide index.',
};

const SHAPE_INDEX: JsonSchema = {
	type: 'integer',
	description: 'Renderer shape index from describe(). Negative indices are not editable.',
};

const TRANSFORM: JsonSchema = {
	type: 'object',
	required: ['x', 'y', 'cx', 'cy', 'rot'],
	properties: {
		x: { type: 'number', description: 'Left offset in EMU.' },
		y: { type: 'number', description: 'Top offset in EMU.' },
		cx: { type: 'number', description: 'Width in EMU.' },
		cy: { type: 'number', description: 'Height in EMU.' },
		rot: { type: 'number', description: 'Rotation in OOXML 60,000ths of a degree.' },
	},
	additionalProperties: false,
};

function slideShapeParams(
	extra: Record<string, JsonSchema> = {},
	requiredExtra: readonly string[] = Object.keys(extra),
): JsonSchema {
	return {
		type: 'object',
		required: ['slideIndex', 'shapeIndex', ...requiredExtra],
		properties: {
			slideIndex: SLIDE_INDEX,
			shapeIndex: SHAPE_INDEX,
			...extra,
		},
		additionalProperties: false,
	};
}

function pptxOp(
	id: string,
	featureArea: string,
	description: string,
	parameters: JsonSchema,
): OpDefinition {
	return {
		id: `pptx.${id}`,
		namespace: 'pptx',
		featureArea,
		description,
		status: 'implemented',
		parameters,
	};
}

function docxOp(
	id: string,
	featureArea: string,
	description: string,
	parameters: JsonSchema,
): OpDefinition {
	return {
		id: `docx.${id}`,
		namespace: 'docx',
		featureArea,
		description,
		status: 'implemented',
		parameters,
	};
}

const DOCX_PARAGRAPH_OFFSET: JsonSchema = {
	type: 'integer',
	minimum: 0,
	description: '0-based character offset within paragraph plain text (describe runs[].text concatenated).',
};

const DOCX_TEXT_POSITION: JsonSchema = {
	type: 'object',
	required: ['blockId', 'offset'],
	properties: {
		blockId: { type: 'string', description: 'Paragraph block id from describe(), e.g. body/p[0].' },
		runId: { type: 'string', description: 'Optional anchor run id; must belong to blockId when provided.' },
		offset: DOCX_PARAGRAPH_OFFSET,
	},
	additionalProperties: false,
};

const DOCX_TEXT_RANGE: JsonSchema = {
	type: 'object',
	required: ['start', 'end'],
	properties: {
		start: DOCX_TEXT_POSITION,
		end: DOCX_TEXT_POSITION,
	},
	additionalProperties: false,
};

/** Canonical operation catalog for agent discovery and schema validation. */
export const OP_CATALOG: readonly OpDefinition[] = [
	// PPTX — text-editing
	pptxOp('updateShapeText', 'text-editing', 'Replace all text in a shape.', slideShapeParams({
		text: {
			type: 'string',
			description: 'Plain text for one shape paragraph. Use replaceShapeParagraphs for multiple paragraphs or lists.',
		},
	})),
	pptxOp('replaceShapeParagraphs', 'text-editing', 'Replace a shape with native PowerPoint paragraphs and list markers. Use for bullets; do not put literal bullet glyphs in text.', slideShapeParams({
		paragraphs: {
			type: 'array',
			description: 'One entry per real PowerPoint paragraph. Each text value must not contain line breaks.',
			items: {
				type: 'object',
				required: ['text', 'listStyle'],
				properties: {
					text: { type: 'string' },
					listStyle: { type: 'string', enum: ['none', 'bullet', 'number'] },
					bold: {
						type: 'boolean',
						description: 'Optional run bold. Heading-first shapes: set true on the title paragraph; omit/false on body so bold does not leak from the template.',
					},
				},
				additionalProperties: false,
			},
		},
	})),
	pptxOp('updateParagraphText', 'text-editing', 'Replace one paragraph of text.', slideShapeParams({
		paragraphIndex: { type: 'integer', minimum: 0 },
		text: { type: 'string' },
	})),
	pptxOp('updateTextRun', 'text-editing', 'Replace one text run.', slideShapeParams({
		paragraphIndex: { type: 'integer', minimum: 0 },
		runIndex: { type: 'integer', minimum: 0 },
		text: { type: 'string' },
	})),
	pptxOp('replaceText', 'text-editing', 'Find and replace text within a slide or deck scope.', {
		type: 'object',
		required: ['query', 'replacement'],
		properties: {
			slideIndex: { ...SLIDE_INDEX, description: 'Optional slide scope. Omit to search all slides.' },
			query: { type: 'string' },
			replacement: { type: 'string' },
			matchCase: { type: 'boolean' },
		},
		additionalProperties: false,
	}),

	// PPTX — text-formatting
	pptxOp('setRunStyle', 'text-formatting', 'Apply bold/italic/underline/font to a run or range.', slideShapeParams({
		paragraphIndex: { type: 'integer', minimum: 0 },
		runIndex: { type: 'integer', minimum: 0 },
		style: { type: 'object', additionalProperties: true, description: 'Run style patch object.' },
	})),
	pptxOp('setParagraphAlignment', 'text-formatting', 'Set paragraph alignment.', slideShapeParams({
		paragraphIndex: { type: 'integer', minimum: 0 },
		align: { type: 'string', enum: ['l', 'ctr', 'r', 'just'] },
	})),
	pptxOp('applyListStyle', 'text-formatting', 'Apply bullet, numbered, or no list style.', slideShapeParams({
		paragraphIndex: { type: 'integer', minimum: 0 },
		style: { type: 'string', enum: ['bullet', 'number', 'none'] },
	})),

	// PPTX — arrange / inspector
	pptxOp('updateTransform', 'arrange', 'Move, resize, or rotate a shape.', slideShapeParams({
		transform: TRANSFORM,
	})),
	pptxOp('setShapeFillColor', 'arrange', 'Set an explicit solid fill color on an editable shape.', slideShapeParams({
		hex: {
			type: 'string',
			description: 'Six-digit RGB hex color, with or without #.',
		},
	})),
	pptxOp('reorderShapes', 'arrange', 'Change z-order of shapes on a slide. Reordering changes renderer shape indices, so describe the slide again before targeting another shape.', {
		type: 'object',
		required: ['slideIndex', 'shapeIndex', 'mode'],
		properties: {
			slideIndex: SLIDE_INDEX,
			shapeIndex: SHAPE_INDEX,
			mode: { type: 'string', enum: ['front', 'back', 'forward', 'backward'] },
		},
		additionalProperties: false,
	}),
	pptxOp('groupShapes', 'arrange', 'Group shapes on a slide.', {
		type: 'object',
		required: ['slideIndex', 'shapeIndices'],
		properties: {
			slideIndex: SLIDE_INDEX,
			shapeIndices: { type: 'array', items: { type: 'integer' }, description: 'At least two shape indices.' },
		},
		additionalProperties: false,
	}),
	pptxOp('ungroupShapes', 'arrange', 'Ungroup a grouped shape.', slideShapeParams()),
	pptxOp('flipShape', 'arrange', 'Flip a shape horizontally or vertically.', slideShapeParams({
		axis: { type: 'string', enum: ['horizontal', 'vertical'] },
	})),
	pptxOp('deleteShape', 'arrange', 'Delete an editable shape. When deleting several shapes, use descending shape indices and describe the slide again.', slideShapeParams()),

	// PPTX — insert
	pptxOp('addImage', 'insert', 'Insert an image on a slide. Default fit is contain: preserve the source aspect ratio and resize it within the requested transform; use fit "stretch" only when explicitly requested.', {
		type: 'object',
		required: ['slideIndex', 'vaultImagePath', 'transform'],
		properties: {
			slideIndex: SLIDE_INDEX,
			vaultImagePath: { type: 'string', description: 'Vault path to a raster image file.' },
			transform: TRANSFORM,
			fit: { type: 'string', enum: ['contain', 'stretch'], description: 'Optional image fit. Defaults to aspect-ratio-preserving resize.' },
		},
		additionalProperties: false,
	}),
	pptxOp('addShape', 'insert', 'Insert a geometry shape.', {
		type: 'object',
		required: ['slideIndex', 'geometry', 'transform'],
		properties: {
			slideIndex: SLIDE_INDEX,
			geometry: { type: 'string', description: 'Shape geometry preset name.' },
			transform: TRANSFORM,
		},
		additionalProperties: false,
	}),
	pptxOp('addTextBox', 'insert', 'Insert an empty text box.', {
		type: 'object',
		required: ['slideIndex', 'transform'],
		properties: {
			slideIndex: SLIDE_INDEX,
			transform: TRANSFORM,
		},
		additionalProperties: false,
	}),
	pptxOp('addTable', 'insert', 'Insert a table.', {
		type: 'object',
		required: ['slideIndex', 'rows', 'cols', 'transform'],
		properties: {
			slideIndex: SLIDE_INDEX,
			rows: { type: 'integer', minimum: 1 },
			cols: { type: 'integer', minimum: 1 },
			transform: TRANSFORM,
		},
		additionalProperties: false,
	}),
	pptxOp('addChart', 'insert', 'Insert a chart of the requested type (default: column).', {
		type: 'object',
		required: ['slideIndex', 'transform'],
		properties: {
			slideIndex: SLIDE_INDEX,
			transform: TRANSFORM,
			chartType: {
				type: 'string',
				enum: [
					'column', 'line', 'pie', 'bar', 'area', 'scatter', 'stock', 'surface', 'radar',
					'treemap', 'sunburst', 'histogram', 'boxWhisker', 'waterfall', 'combo',
				],
				description: 'Chart family to insert. Defaults to column.',
			},
		},
		additionalProperties: false,
	}),

	// PPTX — slide operations
	pptxOp('addSlide', 'slide-operations', 'Add a blank slide after an index.', {
		type: 'object',
		required: ['afterIndex'],
		properties: {
			afterIndex: { type: 'integer', minimum: -1, description: '-1 inserts at start.' },
			layout: { type: 'string', enum: ['blank', 'title', 'titleBody'] },
		},
		additionalProperties: false,
	}),
	pptxOp('deleteSlide', 'slide-operations', 'Delete a slide.', {
		type: 'object',
		required: ['slideIndex'],
		properties: { slideIndex: SLIDE_INDEX },
		additionalProperties: false,
	}),
	pptxOp('moveSlide', 'slide-operations', 'Move a slide left or right.', {
		type: 'object',
		required: ['slideIndex', 'direction'],
		properties: {
			slideIndex: SLIDE_INDEX,
			direction: { type: 'integer', enum: [-1, 1] },
		},
		additionalProperties: false,
	}),
	pptxOp('duplicateSlide', 'slide-operations', 'Duplicate a slide.', {
		type: 'object',
		required: ['slideIndex'],
		properties: { slideIndex: SLIDE_INDEX },
		additionalProperties: false,
	}),
	pptxOp('reorderSlides', 'slide-operations', 'Reorder slides by index list.', {
		type: 'object',
		required: ['order'],
		properties: {
			order: { type: 'array', items: { type: 'integer', minimum: 0 } },
		},
		additionalProperties: false,
	}),
	pptxOp('setSlideBackground', 'slide-operations', 'Set slide background color or image.', {
		type: 'object',
		required: ['slideIndex'],
		properties: {
			slideIndex: SLIDE_INDEX,
			colorHex: { type: 'string', description: 'RRGGBB without #.' },
			vaultImagePath: { type: 'string' },
		},
		additionalProperties: false,
	}),

	// PPTX — image
	pptxOp('setImageCrop', 'image', 'Crop an image shape.', slideShapeParams({
		crop: { type: 'object', additionalProperties: true, description: 'Image crop fractions.' },
	})),
	pptxOp('fitImageToFrame', 'image', 'Center-crop an existing image to fill its current frame while preserving its source aspect ratio.', slideShapeParams()),
	pptxOp('resetImage', 'image', 'Reset image crop and effects.', slideShapeParams()),
	pptxOp('replaceImage', 'image', 'Replace picture media, or convert a non-picture shape/placeholder into a picture that fills the same transform box. Default fit is cover: preserve the source aspect ratio and center-crop it to fill the frame; use fit "stretch" only when explicitly requested.', slideShapeParams({
		vaultImagePath: { type: 'string' },
		fit: { type: 'string', enum: ['cover', 'stretch'], description: 'Optional image fit. Defaults to aspect-ratio-preserving center crop.' },
	}, ['vaultImagePath'])),
	pptxOp('replaceImageFromShape', 'image', 'Replace picture media from another embedded picture shape in the presentation.', slideShapeParams({
		sourceSlideIndex: { ...SLIDE_INDEX, description: '0-based slide index containing the source picture.' },
		sourceShapeIndex: { ...SHAPE_INDEX, description: 'Renderer shape index of the source picture from describe().' },
	})),

	// PPTX — charts
	pptxOp('updateChartData', 'charts', 'Update chart series data.', slideShapeParams({
		data: { type: 'object', additionalProperties: true, description: 'Chart data grid patch.' },
	})),

	// DOCX — review / metadata
	docxOp('removeComments', 'review', 'Remove every comment annotation, inline anchor, and related DOCX package part.', {
		type: 'object',
		required: [],
		properties: {},
		additionalProperties: false,
	}),
	docxOp('setCoreProperties', 'metadata', 'Set the DOCX author and last modifier in docProps/core.xml.', {
		type: 'object',
		required: ['creator', 'lastModifiedBy'],
		properties: {
			creator: { type: 'string', minLength: 1, description: 'Document author (dc:creator).' },
			lastModifiedBy: { type: 'string', minLength: 1, description: 'Most recent editor (cp:lastModifiedBy).' },
		},
		additionalProperties: false,
	}),

	// DOCX — text / font
	docxOp('setRunText', 'font', 'Set run text by stable block/run id (body, header, footer, footnotes, endnotes).', {
		type: 'object',
		required: ['blockId', 'runId', 'text'],
		properties: {
			blockId: { type: 'string', description: 'e.g. body/p[12], header/1/p[0], footnotes/fn[1]/p[0]' },
			runId: { type: 'string', description: 'e.g. body/p[12]/r[0]' },
			text: { type: 'string' },
		},
		additionalProperties: false,
	}),
	docxOp('setRunStyle', 'font', 'Apply run style by stable id.', {
		type: 'object',
		required: ['runId', 'style'],
		properties: {
			runId: { type: 'string' },
			style: { type: 'object', additionalProperties: true },
		},
		additionalProperties: false,
	}),
	docxOp('setParagraphStyle', 'font', 'Apply paragraph style by block id.', {
		type: 'object',
		required: ['blockId', 'style'],
		properties: {
			blockId: { type: 'string' },
			style: { type: 'object', additionalProperties: true },
		},
		additionalProperties: false,
	}),
	docxOp('setParagraphBottomBorder', 'font', 'Set the bottom paragraph border without replacing the paragraph or other border sides.', {
		type: 'object',
		required: ['blockId', 'border'],
		properties: {
			blockId: { type: 'string', description: 'Paragraph block id from describe(), e.g. body/p[0].' },
			border: {
				type: 'object',
				required: ['style'],
				properties: {
					style: { type: 'string', description: 'OOXML border style, e.g. single or double.' },
					size: { type: 'integer', minimum: 0, description: 'Border width in eighths of a point.' },
					space: { type: 'integer', minimum: 0, description: 'Spacing from paragraph text in points.' },
					color: { type: 'string', description: 'Six-digit RGB hex color without #.' },
				},
				additionalProperties: false,
			},
		},
		additionalProperties: false,
	}),

	// DOCX — table
	docxOp('insertTable', 'table', 'Insert a table at a block anchor.', {
		type: 'object',
		required: ['afterBlockId', 'rows', 'cols'],
		properties: {
			afterBlockId: { type: 'string' },
			rows: { type: 'integer', minimum: 1 },
			cols: { type: 'integer', minimum: 1 },
		},
		additionalProperties: false,
	}),
	docxOp('setCellText', 'table', 'Set table cell text.', {
		type: 'object',
		required: ['cellId', 'text'],
		properties: {
			cellId: { type: 'string', description: 'e.g. body/tbl[3]/tr[1]/tc[2]' },
			text: { type: 'string' },
		},
		additionalProperties: false,
	}),
	docxOp('setCellStyle', 'table', 'Set table cell style.', {
		type: 'object',
		required: ['cellId', 'style'],
		properties: {
			cellId: { type: 'string' },
			style: { type: 'object', additionalProperties: true },
		},
		additionalProperties: false,
	}),
	docxOp('deleteTable', 'table', 'Delete a table block.', {
		type: 'object',
		required: ['tableId'],
		properties: {
			tableId: { type: 'string', description: 'e.g. body/tbl[0]' },
		},
		additionalProperties: false,
	}),

	// DOCX — image
	docxOp('insertImage', 'image', 'Insert an image after a block.', {
		type: 'object',
		required: ['afterBlockId', 'vaultImagePath'],
		properties: {
			afterBlockId: { type: 'string' },
			vaultImagePath: { type: 'string' },
		},
		additionalProperties: false,
	}),
	docxOp('replaceImage', 'image', 'Replace an inline image. blockId must be a describe() image block (body/p[N] with an embedded drawing).', {
		type: 'object',
		required: ['blockId', 'vaultImagePath'],
		properties: {
			blockId: { type: 'string', description: 'Image block id from describe(), e.g. body/p[1]' },
			vaultImagePath: { type: 'string' },
		},
		additionalProperties: false,
	}),

	docxOp('insertText', 'font', 'Insert text at a paragraph offset identified by describe() block/run ids.', {
		type: 'object',
		required: ['blockId', 'offset', 'text'],
		properties: {
			blockId: { type: 'string', description: 'Paragraph block id, e.g. body/p[0].' },
			runId: { type: 'string', description: 'Optional anchor run id, e.g. body/p[0]/r[0].' },
			offset: DOCX_PARAGRAPH_OFFSET,
			text: { type: 'string' },
		},
		additionalProperties: false,
	}),
	docxOp('deleteRange', 'font', 'Delete a text range within one paragraph or across consecutive paragraphs in the same part.', {
		type: 'object',
		required: ['range'],
		properties: {
			range: DOCX_TEXT_RANGE,
		},
		additionalProperties: false,
	}),
	docxOp('deleteBlock', 'font', 'Delete one complete paragraph block without merging it into adjacent paragraphs.', {
		type: 'object',
		required: ['blockId'],
		properties: {
			blockId: { type: 'string', description: 'Paragraph block id from describe(), e.g. body/p[0].' },
		},
		additionalProperties: false,
	}),
	docxOp('insertHyperlink', 'font', 'Wrap a single-paragraph text range in an external OOXML hyperlink.', {
		type: 'object',
		required: ['range', 'url'],
		properties: {
			range: DOCX_TEXT_RANGE,
			url: { type: 'string', description: 'http, https, or mailto URL.' },
			displayText: { type: 'string', description: 'Optional replacement text for the linked range.' },
			tooltip: { type: 'string' },
		},
		additionalProperties: false,
	}),
	docxOp('removeHyperlink', 'font', 'Remove hyperlink wrapping in a range while keeping visible text.', {
		type: 'object',
		required: ['range'],
		properties: {
			range: DOCX_TEXT_RANGE,
		},
		additionalProperties: false,
	}),
	docxOp('insertParagraphBreak', 'font', 'Split a paragraph at an offset without rewriting the whole part. Prefer docx.replaceBodyParagraphs when writing a full multi-paragraph letter from scratch.', {
		type: 'object',
		required: ['blockId', 'offset'],
		properties: {
			blockId: { type: 'string', description: 'Paragraph block id, e.g. body/p[0].' },
			runId: { type: 'string', description: 'Optional anchor run id, e.g. body/p[0]/r[0].' },
			offset: DOCX_PARAGRAPH_OFFSET,
		},
		additionalProperties: false,
	}),
	docxOp(
		'replaceBodyParagraphs',
		'font',
		'Replace all body paragraphs/tables with plain paragraphs (preserves trailing sectPr). Prefer this over chaining insertParagraphBreak for new letters.',
		{
			type: 'object',
			required: ['paragraphs'],
			properties: {
				paragraphs: {
					type: 'array',
					items: { type: 'string' },
					description: 'Ordered paragraph plain text. Empty strings become blank paragraphs. Empty array becomes one blank paragraph.',
				},
			},
			additionalProperties: false,
		},
	),

	// DOCX — find-replace
	docxOp('replaceText', 'find-replace', 'Find and replace text across body, headers, footers, footnotes, and endnotes.', {
		type: 'object',
		required: ['query', 'replacement'],
		properties: {
			query: { type: 'string' },
			replacement: { type: 'string' },
			matchCase: { type: 'boolean' },
			wholeWord: { type: 'boolean' },
		},
		additionalProperties: false,
	}),
];

export const OP_IDS = OP_CATALOG.map((op) => op.id);
