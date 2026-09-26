import type { ZipContents } from 'pptx-svg';

import {
  getDescendants,
  getSlidePath,
  getSlideRelationshipsPath,
  parseXml,
  resolvePartPath,
} from './ooxmlXml';

const SLIDE_LAYOUT_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';

/** A layout offered by the presentation's existing slide masters. */
export interface SlideLayoutDefinition {
  /** Stable for the lifetime of the loaded package. */
  readonly id: string;
  readonly layoutPath: string;
  readonly name: string;
  readonly type: string | null;
  /** A slide already using this layout; pptx-svg uses it as the insertion source. */
  readonly representativeSlideIndex: number;
}

function displayName(layoutPath: string, layout: Element): string {
  const matchingName = layout.getAttribute('matchingName')?.trim();
  if (matchingName) return matchingName;

  const commonSlideData = getDescendants(layout, 'cSld')[0];
  const layoutName = commonSlideData?.getAttribute('name')?.trim();
  if (layoutName) return layoutName;

  const type = layout.getAttribute('type')?.trim();
  if (type) return type;
  return layoutPath.split('/').at(-1)?.replace(/\.xml$/i, '') ?? 'Layout';
}

/**
 * Finds the layouts actually attached to existing slides. Keeping a source
 * slide for each one lets us create new slides through pptx-svg without
 * synthesizing a partial, lossy layout relationship graph.
 */
export function listSlideLayouts(
  contents: Pick<ZipContents, 'textFiles'>,
  slideCount: number,
): SlideLayoutDefinition[] {
  const layouts = new Map<string, SlideLayoutDefinition>();

  for (let slideIndex = 0; slideIndex < slideCount; slideIndex += 1) {
    const slidePath = getSlidePath(slideIndex);
    const relationshipPath = getSlideRelationshipsPath(slideIndex);
    const relationshipsXml = contents.textFiles.get(relationshipPath);
    if (!relationshipsXml) continue;

    const relationships = parseXml(relationshipsXml, relationshipPath);
    const relationship = getDescendants(relationships, 'Relationship').find((candidate) =>
      candidate.getAttribute('Type') === SLIDE_LAYOUT_RELATIONSHIP_TYPE
      && candidate.getAttribute('TargetMode') !== 'External'
    );
    const target = relationship?.getAttribute('Target')?.trim();
    if (!target) continue;

    const layoutPath = resolvePartPath(slidePath, target);
    if (layouts.has(layoutPath)) continue;
    const layoutXml = contents.textFiles.get(layoutPath);
    if (!layoutXml) continue;

    const layoutDocument = parseXml(layoutXml, layoutPath);
    const layout = getDescendants(layoutDocument, 'sldLayout')[0];
    if (!layout) continue;

    layouts.set(layoutPath, {
      id: layoutPath,
      layoutPath,
      name: displayName(layoutPath, layout),
      type: layout.getAttribute('type')?.trim() || null,
      representativeSlideIndex: slideIndex,
    });
  }

  return [...layouts.values()];
}
