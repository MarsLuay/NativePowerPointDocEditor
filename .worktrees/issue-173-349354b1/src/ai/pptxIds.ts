export function pptxShapeId(slideIndex: number, shapeIndex: number): string {
	return `slide:${slideIndex}/shape:${shapeIndex}`;
}

export function pptxParagraphId(slideIndex: number, shapeIndex: number, paragraphIndex: number): string {
	return `${pptxShapeId(slideIndex, shapeIndex)}/p:${paragraphIndex}`;
}

export function pptxRunId(
	slideIndex: number,
	shapeIndex: number,
	paragraphIndex: number,
	runIndex: number,
): string {
	return `${pptxParagraphId(slideIndex, shapeIndex, paragraphIndex)}/r:${runIndex}`;
}
