import type { ChartDataGrid } from '../ChartData';
import type { ImageCrop, PresentationEngine } from '../PresentationEngine';
import { parseRenderedSlideSvg } from '../powerpoint/parseRenderedSlideSvg';
import { getShapeIndex, isEditableShapeIndex } from '../powerpoint/svgUtils';
import type { ShapeTransform } from 'pptx-svg';
import { pptxParagraphId, pptxRunId, pptxShapeId } from './pptxIds';

export interface PptxDescribedRun {
	id: string;
	text: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	fontFamily?: string | null;
	fontSizePt?: number | null;
	color?: string | null;
	highlight?: string | null;
}

export interface PptxDescribedParagraph {
	id: string;
	text: string;
	/** Explicit native list marker, or null when the paragraph has none. */
	listStyle: 'none' | 'bullet' | 'number' | null;
	align?: string | null;
	runs?: PptxDescribedRun[];
}

export interface PptxDescribedShapeStyle {
	fill: string | null;
	stroke: string | null;
	strokeWidthPt: number | null;
}

export interface PptxDescribedSlideBackground {
	colorHex: string | null;
	imageHref: string | null;
	crop: ImageCrop | null;
}

export interface PptxDescribedShape {
	id: string;
	index: number;
	kind: string;
	editable: boolean;
	text: string | null;
	paragraphs?: PptxDescribedParagraph[];
	transform: ShapeTransform;
	style: PptxDescribedShapeStyle | null;
	crop?: ImageCrop | null;
	chartData?: ChartDataGrid;
}

export interface PptxDescribedSlide {
	index: number;
	background?: PptxDescribedSlideBackground;
	shapes: PptxDescribedShape[];
}

export interface PptxDescribeSnapshot {
	format: 'pptx';
	file: string;
	slideCount: number;
	runtime?: {
		rendererBackend: 'wasm-gc' | 'js';
	};
	slides: PptxDescribedSlide[];
}

function resolveShapeKind(
	engine: PresentationEngine,
	slideIndex: number,
	shapeIndex: number,
	shapeEl: SVGGElement,
): string {
	const type = shapeEl.getAttribute('data-ooxml-shape-type');
	if (type === 'table') return 'table';
	if (type === 'chart') return 'chart';
	if (type === 'group') return 'group';
	if (engine.isImageShape(slideIndex, shapeIndex) || shapeEl.querySelector('image')) {
		return 'image';
	}
	if (shapeEl.querySelector('text')) {
		return 'textbox';
	}
	return type ?? 'shape';
}

function mapRunStyle(
	slideIndex: number,
	shapeIndex: number,
	paragraphIndex: number,
	runIndex: number,
	text: string,
	style: NonNullable<ReturnType<PresentationEngine['getRunStyle']>>,
): PptxDescribedRun {
	return {
		id: pptxRunId(slideIndex, shapeIndex, paragraphIndex, runIndex),
		text,
		...(style.bold ? { bold: true } : {}),
		...(style.italic ? { italic: true } : {}),
		...(style.underline ? { underline: true } : {}),
		...(style.fontFamily ? { fontFamily: style.fontFamily } : {}),
		...(style.fontSizePt !== null ? { fontSizePt: style.fontSizePt } : {}),
		...(style.color ? { color: style.color } : {}),
		...(style.highlight ? { highlight: style.highlight } : {}),
	};
}

function describeShapeRuns(
	engine: PresentationEngine,
	slideIndex: number,
	shapeIndex: number,
	paragraphIndex: number,
): PptxDescribedRun[] {
	const runs: PptxDescribedRun[] = [];
	for (let runIndex = 0; runIndex < 200; runIndex++) {
		const style = engine.getRunStyle(slideIndex, shapeIndex, paragraphIndex, runIndex);
		if (!style) break;
		const text = engine.getTextRunText(slideIndex, shapeIndex, paragraphIndex, runIndex) ?? '';
		runs.push(mapRunStyle(slideIndex, shapeIndex, paragraphIndex, runIndex, text, style));
	}
	return runs;
}

function describeShapeParagraphs(
	engine: PresentationEngine,
	slideIndex: number,
	shapeIndex: number,
): PptxDescribedParagraph[] {
	const paragraphs: PptxDescribedParagraph[] = [];
	for (let paragraphIndex = 0; paragraphIndex < 200; paragraphIndex++) {
		const text = engine.getParagraphRunText(slideIndex, shapeIndex, paragraphIndex);
		if (text === null) break;
		const runs = describeShapeRuns(engine, slideIndex, shapeIndex, paragraphIndex);
		const align = engine.getRunStyle(slideIndex, shapeIndex, paragraphIndex, 0)?.alignment ?? null;
		paragraphs.push({
			id: pptxParagraphId(slideIndex, shapeIndex, paragraphIndex),
			text,
			listStyle: engine.getParagraphListStyle(slideIndex, shapeIndex, paragraphIndex),
			align,
			runs: runs.length > 0 ? runs : undefined,
		});
	}
	return paragraphs;
}

function shapeHasEditableText(
	engine: PresentationEngine,
	slideIndex: number,
	shapeIndex: number,
): boolean {
	if (!isEditableShapeIndex(shapeIndex)) {
		return false;
	}
	return engine.getParagraphRunText(slideIndex, shapeIndex, 0) !== null;
}

function describeShape(
	engine: PresentationEngine,
	slideIndex: number,
	shapeEl: SVGGElement,
): PptxDescribedShape | null {
	const shapeIndex = getShapeIndex(shapeEl);
	if (shapeIndex === null) return null;

	const editable = isEditableShapeIndex(shapeIndex);
	const transform = engine.getShapeTransform(shapeEl);
	const kind = resolveShapeKind(engine, slideIndex, shapeIndex, shapeEl);
	const paragraphs = shapeHasEditableText(engine, slideIndex, shapeIndex)
		? describeShapeParagraphs(engine, slideIndex, shapeIndex)
		: [];
	const text = paragraphs.length > 0
		? paragraphs.map((paragraph) => paragraph.text).join('\n')
		: null;
	const style = editable ? engine.getShapeVisualStyle(slideIndex, shapeIndex) : null;
	const crop = editable && kind === 'image'
		? engine.getImageCrop(slideIndex, shapeIndex)
		: undefined;
	const chartData = editable && kind === 'chart'
		? engine.getChartDataGrid(slideIndex, shapeIndex) ?? undefined
		: undefined;

	return {
		id: pptxShapeId(slideIndex, shapeIndex),
		index: shapeIndex,
		kind,
		editable,
		text,
		paragraphs: paragraphs.length > 0 ? paragraphs : undefined,
		transform,
		style,
		...(crop !== undefined ? { crop } : {}),
		...(chartData ? { chartData } : {}),
	};
}

function describeSlideBackground(engine: PresentationEngine, slideIndex: number): PptxDescribedSlideBackground {
	return engine.getSlideBackgroundDescribe(slideIndex);
}

export function describePptxFromEngine(
	engine: PresentationEngine,
	filePath: string,
): PptxDescribeSnapshot {
	const slides: PptxDescribedSlide[] = [];
	for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
		const rendered = engine.renderSlide(slideIndex);
		const svg = parseRenderedSlideSvg(rendered.svg);
		const background = describeSlideBackground(engine, slideIndex);
		const shapes = engine.getShapes(svg)
			.map((shapeEl) => describeShape(engine, slideIndex, shapeEl))
			.filter((shape): shape is PptxDescribedShape => shape !== null)
			.sort((left, right) => left.index - right.index);

		slides.push({
			index: slideIndex,
			...(background.colorHex || background.imageHref || background.crop
				? { background }
				: {}),
			shapes,
		});
	}

	return {
		format: 'pptx',
		file: filePath,
		slideCount: engine.slideCount,
		runtime: { rendererBackend: engine.getRendererBackend() },
		slides,
	};
}
