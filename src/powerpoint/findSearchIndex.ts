import type { PowerPointFindMatch } from './types';
import {
  getDescendants,
  getElementChildren,
  getShapeChildren,
  getShapeTree,
  parseXml,
  SHAPE_ELEMENT_NAMES,
} from './ooxmlXml';

function normalizeSearchText(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

export interface PowerPointFindSearchIndexSlide {
  slideIndex: number;
  shapeMatches: PowerPointFindMatch[];
  fallbackText: string;
}

/**
 * Extracts searchable text from slide OOXML without invoking the SVG renderer.
 * Rendering every slide just to search its text makes Find unusably slow for
 * large poster templates and image-heavy decks.
 */
export function createFindSearchIndexSlideFromOoxml(
  slideIndex: number,
  slideXml: string,
): PowerPointFindSearchIndexSlide {
  const slideDocument = parseXml(slideXml, `slide ${slideIndex + 1}`);
  const shapeMatches: PowerPointFindMatch[] = [];
  const addShape = (shape: Element, shapeIndex: number): void => {
    const text = normalizeSearchText(
      getDescendants(shape, 't').map((element) => element.textContent ?? '').join(''),
    );
    if (text) shapeMatches.push({ slideIndex, shapeIndex, text });
  };

  const shapes = getShapeChildren(getShapeTree(slideDocument));
  shapes.forEach((shape, shapeIndex) => {
    addShape(shape, shapeIndex);
    if (shape.localName !== 'grpSp') return;

    getElementChildren(shape)
      .filter((child) => SHAPE_ELEMENT_NAMES.has(child.localName))
      .forEach((child, childIndex) => addShape(child, (shapeIndex * 1000) + childIndex));
  });

  return {
    slideIndex,
    shapeMatches,
    fallbackText: normalizeSearchText(slideDocument.documentElement.textContent ?? ''),
  };
}

/**
 * Filters already-extracted deck text. The fallback preserves the renderer
 * behavior for text that is not inside an editable shape group.
 */
export function collectFindMatchesFromSearchIndex(
  searchIndex: readonly PowerPointFindSearchIndexSlide[],
  query: string,
): PowerPointFindMatch[] {
  const normalizedQuery = query.trim().toLocaleLowerCase();
  if (!normalizedQuery) return [];

  const matches: PowerPointFindMatch[] = [];
  for (const slide of searchIndex) {
    const shapeMatches = slide.shapeMatches.filter((match) => (
      match.text.toLocaleLowerCase().includes(normalizedQuery)
    ));
    if (shapeMatches.length > 0) {
      matches.push(...shapeMatches);
    } else if (slide.fallbackText.toLocaleLowerCase().includes(normalizedQuery)) {
      matches.push({ slideIndex: slide.slideIndex, shapeIndex: null, text: slide.fallbackText });
    }
  }
  return matches;
}
