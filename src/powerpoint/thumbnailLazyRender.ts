/** Slide count at or below this renders every thumbnail during initial filmstrip build. */
export const FULL_THUMBNAIL_RENDER_MAX_SLIDES = 12;

/** On large decks, render this many slides on each side of the active slide first. */
export const THUMBNAIL_PRIORITY_RADIUS = 2;

/** Off-screen thumbnails filled per idle callback after the priority pass. */
export const THUMBNAIL_IDLE_BATCH_SIZE = 2;

export function shouldUseLazyThumbnails(slideCount: number): boolean {
	return slideCount > FULL_THUMBNAIL_RENDER_MAX_SLIDES;
}

export function priorityThumbnailIndices(
	currentSlide: number,
	slideCount: number,
	radius: number = THUMBNAIL_PRIORITY_RADIUS,
): number[] {
	if (slideCount <= 0) return [];

	const start = Math.max(0, currentSlide - radius);
	const end = Math.min(slideCount - 1, currentSlide + radius);
	const indices: number[] = [];
	for (let index = start; index <= end; index += 1) {
		indices.push(index);
	}
	return indices;
}

export function remainingThumbnailIndices(
	slideCount: number,
	rendered: ReadonlySet<number>,
): number[] {
	const indices: number[] = [];
	for (let index = 0; index < slideCount; index += 1) {
		if (!rendered.has(index)) {
			indices.push(index);
		}
	}
	return indices;
}

export function sortThumbnailIndicesByProximity(indices: number[], currentSlide: number): number[] {
	return [...indices].sort((left, right) => {
		const leftDistance = Math.abs(left - currentSlide);
		const rightDistance = Math.abs(right - currentSlide);
		return leftDistance - rightDistance || left - right;
	});
}
