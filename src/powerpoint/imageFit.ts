import type { RasterImageDimensions } from './imageDimensions';

/** Inset crop percentages for a centered PowerPoint picture crop. */
export interface CenteredImageCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * Center-crop a source image to cover a frame without distorting it.
 * A zero crop means the source already matches the frame aspect ratio.
 */
export function computeCenteredCoverCrop(
  imageSize: RasterImageDimensions | null,
  frameWidth: number,
  frameHeight: number,
): CenteredImageCrop | null {
  if (!imageSize || frameWidth <= 0 || frameHeight <= 0) return null;

  const imageAspect = imageSize.width / imageSize.height;
  const frameAspect = frameWidth / frameHeight;
  const epsilon = 0.01;
  if (Math.abs(imageAspect - frameAspect) < epsilon) {
    return { left: 0, top: 0, right: 0, bottom: 0 };
  }

  if (imageAspect > frameAspect) {
    const visible = frameAspect / imageAspect;
    const side = ((1 - visible) / 2) * 100;
    return { left: side, top: 0, right: side, bottom: 0 };
  }

  const visible = imageAspect / frameAspect;
  const side = ((1 - visible) / 2) * 100;
  return { left: 0, top: side, right: 0, bottom: side };
}

/** Fit a source image within a default insertion box while preserving its ratio. */
export function fitImageWithinBounds(
  imageSize: RasterImageDimensions | null,
  maximumWidth: number,
  maximumHeight: number,
): RasterImageDimensions {
  const fallback = {
    width: Math.max(1, Math.round(maximumWidth)),
    height: Math.max(1, Math.round(maximumHeight)),
  };
  if (!imageSize || maximumWidth <= 0 || maximumHeight <= 0) return fallback;

  const scale = Math.min(maximumWidth / imageSize.width, maximumHeight / imageSize.height);
  return {
    width: Math.max(1, Math.round(imageSize.width * scale)),
    height: Math.max(1, Math.round(imageSize.height * scale)),
  };
}
