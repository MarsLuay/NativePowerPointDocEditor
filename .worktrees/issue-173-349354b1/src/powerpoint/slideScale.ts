// Pure fit-then-zoom sizing for the PowerPoint slide surface. Kept free of DOM
// so tests can prove scrollbar-width feedback would thrash display pixels.

export const DEFAULT_SCROLLBAR_SIZE_PX = 15;

export function computeSlideDisplaySize(args: {
  intrinsicWidth: number;
  intrinsicHeight: number;
  availableWidth: number;
  availableHeight: number;
  zoomLevel: number;
}): { fitScale: number; scale: number; width: number; height: number } {
  const availableWidth = Math.max(1, args.availableWidth);
  const availableHeight = Math.max(1, args.availableHeight);
  const fitScale = Math.min(
    1,
    availableWidth / Math.max(1, args.intrinsicWidth),
    availableHeight / Math.max(1, args.intrinsicHeight),
  );
  const scale = Math.max(0.05, fitScale * args.zoomLevel);
  return {
    fitScale,
    scale,
    width: Math.max(1, Math.floor(args.intrinsicWidth * scale)),
    height: Math.max(1, Math.floor(args.intrinsicHeight * scale)),
  };
}

/**
 * Size the slide against a scrollbar-stable available box.
 *
 * `clientWidth`/`clientHeight` shrink when overflow bars appear. For a tall
 * poster near ~200% zoom that toggles the *horizontal* bar, height flips by
 * ~15px → fitScale thrash. Callers should pass border-box dims (offsetWidth /
 * offsetHeight), then this reserves bar thickness when either axis overflows.
 */
export function computeStableSlideDisplaySize(args: {
  intrinsicWidth: number;
  intrinsicHeight: number;
  availableWidth: number;
  availableHeight: number;
  zoomLevel: number;
  scrollbarSize?: number;
}): {
  fitScale: number;
  scale: number;
  width: number;
  height: number;
  reservedScrollbars: boolean;
  availableWidth: number;
  availableHeight: number;
} {
  const scrollbarSize = Math.max(0, args.scrollbarSize ?? DEFAULT_SCROLLBAR_SIZE_PX);
  const first = computeSlideDisplaySize(args);
  const overflowsWidth = first.width > args.availableWidth;
  const overflowsHeight = first.height > args.availableHeight;
  if (!overflowsWidth && !overflowsHeight) {
    return {
      ...first,
      reservedScrollbars: false,
      availableWidth: args.availableWidth,
      availableHeight: args.availableHeight,
    };
  }

  // Vertical bar eats width; horizontal bar eats height. Near dual-overflow
  // thresholds reserve both so the result does not flip next frame.
  const reservedWidth = Math.max(1, args.availableWidth - scrollbarSize);
  const reservedHeight = Math.max(1, args.availableHeight - scrollbarSize);
  const second = computeSlideDisplaySize({
    ...args,
    availableWidth: reservedWidth,
    availableHeight: reservedHeight,
  });
  return {
    ...second,
    reservedScrollbars: true,
    availableWidth: reservedWidth,
    availableHeight: reservedHeight,
  };
}

/** True when toggling a typical scrollbar width on available dims flips floor(size). */
export function slideDisplaySizeOscillatesWithScrollbar(args: {
  intrinsicWidth: number;
  intrinsicHeight: number;
  availableWidth: number;
  availableHeight: number;
  zoomLevel: number;
  scrollbarSize?: number;
}): boolean {
  const scrollbarSize = args.scrollbarSize ?? DEFAULT_SCROLLBAR_SIZE_PX;
  const without = computeSlideDisplaySize(args);
  const withBars = computeSlideDisplaySize({
    ...args,
    availableWidth: Math.max(1, args.availableWidth - scrollbarSize),
    availableHeight: Math.max(1, args.availableHeight - scrollbarSize),
  });
  return without.width !== withBars.width || without.height !== withBars.height;
}
