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

export function bringGridTextToFront(svg: SVGSVGElement): void {
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

export function markEditableTextRuns(svg: SVGSVGElement): void {
  svg.querySelectorAll('tspan[data-ooxml-run-idx]').forEach((run) => {
    if (!run.closest(GENERATED_GRID_SELECTOR)) {
      run.classList.add('native-powerpoint-editable-text');
    }
  });
}

export function normalizeSvgForDisplay(svg: SVGSVGElement): void {
  ensureSvgViewBox(svg);
  applyShapeFlipTransforms(svg);
  bringGridTextToFront(svg);
  markEditableTextRuns(svg);
}
