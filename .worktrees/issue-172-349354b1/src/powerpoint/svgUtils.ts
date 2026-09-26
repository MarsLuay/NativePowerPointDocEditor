// Pure SVG and shape-transform helpers for the PowerPoint view. Extracted from
// NativePowerPointView.ts; these operate only on the elements/values passed in.

import type { GeneratedTextKind } from '../PresentationEngine';
import type { ShapeTransform } from 'pptx-svg';
import { applyBackgroundAwareTextHalos } from '../TextHalo';
import { GENERATED_GRID_SELECTOR } from './constants';
import { normalizeSearchText } from './textUtils';
import type { SlideSize } from './types';
import { applyShapeFlipTransforms } from './shapeFlipTransforms';

export function cloneTransform(transform: ShapeTransform): ShapeTransform {
  return {
    x: transform.x,
    y: transform.y,
    cx: transform.cx,
    cy: transform.cy,
    rot: transform.rot
  };
}

export function getShapeIndex(shape: Element | null): number | null {
  const raw = shape?.getAttribute('data-ooxml-shape-idx');
  if (!raw) return null;

  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

/**
 * True when an already-known SVG group is the selected top-level shape. This
 * lets drag commits reuse their preview element instead of searching a large
 * slide SVG again.
 */
export function isTopLevelShapeForIndex(
  shape: Element | null | undefined,
  shapeIndex: number,
): boolean {
  return getShapeIndex(shape ?? null) === shapeIndex
    && !shape?.parentElement?.closest('g[data-ooxml-shape-idx]');
}

/** A translated multi-selection preview is already the final visual state. */
export function canKeepLiveGroupMovePreview(
  mode: 'move' | 'resize' | 'rotate',
  flipAxes: { horizontal: boolean; vertical: boolean },
): boolean {
  return mode === 'move' && !flipAxes.horizontal && !flipAxes.vertical;
}

/**
 * Whether a renderer shape index can be persisted back to slide OOXML. The
 * pptx-svg renderer emits negative `data-ooxml-shape-idx` values for shapes it
 * renders but does not store directly on the slide (e.g. inherited layout/master
 * content); those cannot be transformed via the renderer or the OOXML fallback.
 */
export function isEditableShapeIndex(index: number | null): index is number {
  return index !== null && Number.isInteger(index) && index >= 0;
}

/** Layout/master decorations use negative indices and should not receive clicks. */
export function isSelectableShapeIndex(index: number | null): index is number {
  return isEditableShapeIndex(index);
}

export function parseSvgLength(value: string | null): number | null {
  if (!value || value.includes('%')) return null;

  const match = value.trim().match(/^([0-9]*\.?[0-9]+)(?:px)?$/i);
  if (!match) return null;

  const parsed = Number(match[1]);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
}

export function getSvgIntrinsicSize(svg: SVGSVGElement): SlideSize | null {
  const viewBox = svg.viewBox.baseVal;
  if (viewBox.width > 0 && viewBox.height > 0) {
    return { width: viewBox.width, height: viewBox.height };
  }

  const width = parseSvgLength(svg.getAttribute('width'));
  const height = parseSvgLength(svg.getAttribute('height'));
  if (width && height) {
    return { width, height };
  }

  return null;
}

/**
 * `pptx-svg` shape fragments may start with a slide-level `<defs>` block.
 * A caller that swaps only the returned shape group must fall back to a full
 * slide render: dropping those definitions can make the replacement invisible.
 */
export function svgFragmentHasDefinitions(fragment: string): boolean {
  return /<\s*defs(?:\s|>)/i.test(fragment);
}

export function transformsMatch(a: ShapeTransform, b: ShapeTransform): boolean {
  return a.x === b.x && a.y === b.y && a.cx === b.cx && a.cy === b.cy && a.rot === b.rot;
}

export function ensureSvgViewBox(svg: SVGSVGElement): void {
  const viewBox = svg.viewBox.baseVal;
  if (viewBox.width > 0 && viewBox.height > 0) return;

  const width = parseSvgLength(svg.getAttribute('width'));
  const height = parseSvgLength(svg.getAttribute('height'));
  if (!width || !height) return;

  svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
  if (!svg.hasAttribute('preserveAspectRatio')) {
    svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
  }
}

export function bringGridTextToFront(svg: Element): void {
  const gridGroups = svg.querySelectorAll(GENERATED_GRID_SELECTOR);

  for (const gridGroup of Array.from(gridGroups)) {
    const kind = gridGroup.getAttribute('data-ooxml-shape-type') as GeneratedTextKind;
    const occurrences = new Map<string, number>();
    applyBackgroundAwareTextHalos(gridGroup, kind);
    gridGroup.querySelectorAll('text').forEach((text, labelIndex) => {
      const normalizedText = normalizeSearchText(text.textContent || '');
      const occurrence = occurrences.get(normalizedText) ?? 0;
      occurrences.set(normalizedText, occurrence + 1);

      text.classList.add('native-powerpoint-grid-label');
      text.setAttribute('data-native-powerpoint-generated-kind', kind);
      text.setAttribute('data-native-powerpoint-label-index', String(labelIndex));
      text.setAttribute('data-native-powerpoint-label-occurrence', String(occurrence));
    });

    const containers = [gridGroup, ...Array.from(gridGroup.querySelectorAll('g'))];
    for (const container of containers) {
      const textChildren = Array.from(container.children).filter(
        (child) => child.tagName.toLowerCase() === 'text'
      );
      for (const textChild of textChildren) {
        container.appendChild(textChild);
      }
    }
  }
}

export function markEditableTextRuns(svg: Element): void {
  svg.querySelectorAll('tspan[data-ooxml-run-idx]').forEach((run) => {
    if (!run.closest(GENERATED_GRID_SELECTOR)) {
      run.classList.add('native-powerpoint-editable-text');
    }
  });
}

export function normalizePictureStretchImages(svg: Element): void {
  svg.querySelectorAll('g[data-ooxml-shape-type="picture"][data-ooxml-blip-stretch="1"] image')
    .forEach((image) => {
      image.setAttribute('preserveAspectRatio', 'none');
    });
}

export function normalizeSvgForDisplay(svg: SVGSVGElement): void {
  ensureSvgViewBox(svg);
  applyShapeFlipTransforms(svg);
  normalizePictureStretchImages(svg);
  bringGridTextToFront(svg);
  markEditableTextRuns(svg);
}

/**
 * Normalize one freshly rendered shape without walking an unchanged slide.
 * The shape is already attached to its slide, so flip geometry can read the
 * root SVG's scale while every other pass remains local to the replacement.
 */
export function normalizeShapeForDisplay(shape: SVGGElement): void {
  applyShapeFlipTransforms(shape);
  normalizePictureStretchImages(shape);
  bringGridTextToFront(shape);
  markEditableTextRuns(shape);
}
