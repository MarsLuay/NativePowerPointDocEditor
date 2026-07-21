/** Axis-aligned box in SVG user units (picture frame or expanded source image). */
export interface PicturePreviewBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Keep srcRect crop proportions while the visible frame moves/resizes.
 *
 * Cropped pictures render as an expanded {@link SVGImageElement} plus a clip
 * rect the size of the OOXML frame. Live resize must scale both together;
 * fitting only the image AABB to the overlay collapses the crop.
 */
export function scaleCroppedPicturePreview(
  image: PicturePreviewBox,
  clip: PicturePreviewBox,
  nextClip: PicturePreviewBox,
): { image: PicturePreviewBox; clip: PicturePreviewBox } {
  if (
    !(image.width > 0)
    || !(image.height > 0)
    || !(clip.width > 0)
    || !(clip.height > 0)
    || !(nextClip.width > 0)
    || !(nextClip.height > 0)
  ) {
    return { image: { ...image }, clip: { ...nextClip } };
  }

  const leftRatio = (clip.x - image.x) / image.width;
  const topRatio = (clip.y - image.y) / image.height;
  const visibleWidthRatio = clip.width / image.width;
  const visibleHeightRatio = clip.height / image.height;
  if (!(visibleWidthRatio > 0) || !(visibleHeightRatio > 0)) {
    return { image: { ...image }, clip: { ...nextClip } };
  }

  const width = nextClip.width / visibleWidthRatio;
  const height = nextClip.height / visibleHeightRatio;
  return {
    clip: { ...nextClip },
    image: {
      x: nextClip.x - width * leftRatio,
      y: nextClip.y - height * topRatio,
      width,
      height,
    },
  };
}
