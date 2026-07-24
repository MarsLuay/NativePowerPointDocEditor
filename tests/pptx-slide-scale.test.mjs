import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { build } from 'esbuild';

const require = createRequire(import.meta.url);

async function loadSlideScale() {
  const dir = await mkdtemp(join(tmpdir(), 'npde-slide-scale-'));
  const outfile = join(dir, 'slide-scale.cjs');
  try {
    await build({
      entryPoints: ['src/powerpoint/slideScale.ts'],
      bundle: true,
      format: 'cjs',
      platform: 'node',
      outfile,
      logLevel: 'silent',
    });
    return require(outfile);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

// 24x36 poster at 96dpi (matches Job/PNNL poster EMUs → SVG px).
const POSTER = { width: 2304, height: 3456 };

test('poster at ~206% flips floor size when scrollbar width is subtracted from available box', async () => {
  const {
    computeSlideDisplaySize,
    slideDisplaySizeOscillatesWithScrollbar,
  } = await loadSlideScale();

  // 850×650 pane minus 32px padding each side → 786×586 available.
  const availableWidth = 850 - 64;
  const availableHeight = 650 - 64;
  const zoomLevel = 2.06;

  assert.equal(
    slideDisplaySizeOscillatesWithScrollbar({
      intrinsicWidth: POSTER.width,
      intrinsicHeight: POSTER.height,
      availableWidth,
      availableHeight,
      zoomLevel,
      scrollbarSize: 15,
    }),
    true,
    'documents the pre-gutter bug class for this poster/zoom',
  );

  const stable = computeSlideDisplaySize({
    intrinsicWidth: POSTER.width,
    intrinsicHeight: POSTER.height,
    availableWidth,
    availableHeight,
    zoomLevel,
  });
  const again = computeSlideDisplaySize({
    intrinsicWidth: POSTER.width,
    intrinsicHeight: POSTER.height,
    availableWidth,
    availableHeight,
    zoomLevel,
  });
  assert.deepEqual(again, stable, 'stable available box must not thrash display size');
});

test('stable sizing settles when horizontal bar would flip height-limited fitScale', async () => {
  const {
    computeSlideDisplaySize,
    computeStableSlideDisplaySize,
  } = await loadSlideScale();

  // Repro from live logs at zoom 2.05: availableHeight 663 vs 648 flips width
  // across the 889 client threshold (906 ↔ 885).
  const rawAvailableWidth = 889;
  const rawAvailableHeight = 663;
  const zoomLevel = 2.05;

  const naive = computeSlideDisplaySize({
    intrinsicWidth: POSTER.width,
    intrinsicHeight: POSTER.height,
    availableWidth: rawAvailableWidth,
    availableHeight: rawAvailableHeight,
    zoomLevel,
  });
  assert.ok(naive.width > rawAvailableWidth, 'naive pass overflows width (needs h-bar)');

  const shrunk = computeSlideDisplaySize({
    intrinsicWidth: POSTER.width,
    intrinsicHeight: POSTER.height,
    availableWidth: rawAvailableWidth,
    availableHeight: rawAvailableHeight - 15,
    zoomLevel,
  });
  assert.notEqual(naive.width, shrunk.width, 'h-bar height shrink flips display width');

  const stable = computeStableSlideDisplaySize({
    intrinsicWidth: POSTER.width,
    intrinsicHeight: POSTER.height,
    availableWidth: rawAvailableWidth,
    availableHeight: rawAvailableHeight,
    zoomLevel,
    scrollbarSize: 15,
  });
  assert.equal(stable.reservedScrollbars, true);
  const again = computeStableSlideDisplaySize({
    intrinsicWidth: POSTER.width,
    intrinsicHeight: POSTER.height,
    availableWidth: rawAvailableWidth,
    availableHeight: rawAvailableHeight,
    zoomLevel,
    scrollbarSize: 15,
  });
  assert.deepEqual(
    { width: again.width, height: again.height },
    { width: stable.width, height: stable.height },
  );
  // Settled size must match the reserved-bar pass, not the oscillating naive one.
  assert.equal(stable.width, shrunk.width);
  assert.equal(stable.height, shrunk.height);
});

test('computeSlideDisplaySize floors width/height from fitScale * zoom', async () => {
  const { computeSlideDisplaySize } = await loadSlideScale();
  const result = computeSlideDisplaySize({
    intrinsicWidth: 1000,
    intrinsicHeight: 500,
    availableWidth: 800,
    availableHeight: 600,
    zoomLevel: 1,
  });
  assert.equal(result.fitScale, 0.8);
  assert.equal(result.width, 800);
  assert.equal(result.height, 400);
});
