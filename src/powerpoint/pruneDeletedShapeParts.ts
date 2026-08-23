import { buildZip, extractZip, type ZipContents } from 'pptx-svg';
import {
  getDescendants,
  getSlidePath,
  getSlideRelationshipsPath,
  parseXml,
  resolvePartPath,
  serializeXml,
} from './ooxmlXml';

const DRAWING_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const PRUNABLE_PART_PREFIXES = ['ppt/media/', 'ppt/charts/', 'ppt/embeddings/'] as const;

function slideStillReferencesRelationship(slideXml: string, relationshipId: string): boolean {
  const escaped = relationshipId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`=["']${escaped}["']`).test(slideXml);
}

export interface ShapePartPruneResult {
  buffer: ArrayBuffer;
  removedPartPaths: string[];
  removedRelationshipIds: string[];
  removedExternalTargets: string[];
}

/** Collect `r:*` relationship ids from a shape subtree (blip embeds, hyperlinks, charts). */
export function collectShapeRelationshipIds(shape: Element): string[] {
  const ids = new Set<string>();
  const elements = [shape, ...Array.from(shape.getElementsByTagName('*'))];
  for (const element of elements) {
    for (const attribute of Array.from(element.attributes)) {
      if (
        (attribute.namespaceURI === DRAWING_RELATIONSHIP_NAMESPACE || attribute.prefix === 'r')
        && attribute.value
      ) {
        ids.add(attribute.value);
      }
    }
  }
  return [...ids];
}

function relationshipsPathToPartPath(relsPath: string): string {
  return relsPath.replace(/\/_rels\/([^/]+)\.rels$/, '/$1');
}

export function collectReferencedInternalParts(
  zip: ZipContents,
  textOverrides: Map<string, string>,
): Set<string> {
  const referenced = new Set<string>();
  const textFiles = new Map(zip.textFiles);
  for (const [path, contents] of textOverrides) {
    textFiles.set(path, contents);
  }

  for (const [relsPath, xml] of textFiles) {
    if (!relsPath.includes('/_rels/') || !relsPath.endsWith('.rels')) continue;
    const sourcePart = relationshipsPathToPartPath(relsPath);
    const document = parseXml(xml, relsPath);
    for (const relationship of getDescendants(document, 'Relationship')) {
      const target = relationship.getAttribute('Target');
      if (!target || relationship.getAttribute('TargetMode') === 'External') continue;
      referenced.add(resolvePartPath(sourcePart, target));
    }
  }
  return referenced;
}

export function isPrunablePart(partPath: string): boolean {
  return PRUNABLE_PART_PREFIXES.some((prefix) => partPath.startsWith(prefix));
}

function stripContentTypeOverrides(contentTypesXml: string, removedParts: Iterable<string>): string {
  const removed = new Set(
    [...removedParts].map((partPath) => `/${partPath}`),
  );
  if (removed.size === 0) return contentTypesXml;

  const document = parseXml(contentTypesXml, '[Content_Types].xml');
  let changed = false;
  // Snapshot before removeChild — live NodeLists shift under for…of.
  for (const override of Array.from(getDescendants(document, 'Override'))) {
    const partName = override.getAttribute('PartName');
    if (!partName || !removed.has(partName)) continue;
    override.parentNode?.removeChild(override);
    changed = true;
  }
  return changed ? serializeXml(document) : contentTypesXml;
}

/**
 * After shapes are removed from `slideDoc`, drop unused slide relationships and
 * package parts (media/charts/embeddings) that nothing else references.
 */
export async function pruneAfterShapeDeletion(
  packageBuffer: ArrayBuffer,
  slideIndex: number,
  slideDoc: XMLDocument,
  deletedRelationshipIds: readonly string[],
): Promise<ShapePartPruneResult> {
  const zip = await extractZip(packageBuffer);
  const slidePath = getSlidePath(slideIndex);
  const relationshipsPath = getSlideRelationshipsPath(slideIndex);
  const textModifications = new Map<string, string>([
    [slidePath, serializeXml(slideDoc)],
  ]);
  const removals = new Set<string>();
  const removedRelationshipIds: string[] = [];
  const removedExternalTargets: string[] = [];

  const uniqueDeletedIds = [...new Set(deletedRelationshipIds)];
  if (uniqueDeletedIds.length > 0) {
    const relationshipsXml = zip.textFiles.get(relationshipsPath);
    if (relationshipsXml) {
      const remainingSlideXml = textModifications.get(slidePath) ?? '';
      const stillNeeded = new Set(
        uniqueDeletedIds.filter((relationshipId) =>
          slideStillReferencesRelationship(remainingSlideXml, relationshipId)
        ),
      );
      const relationships = parseXml(relationshipsXml, relationshipsPath);
      // Snapshot before removeChild — live NodeLists shift under for…of.
      for (const relationship of Array.from(getDescendants(relationships, 'Relationship'))) {
        const relationshipId = relationship.getAttribute('Id');
        if (!relationshipId || stillNeeded.has(relationshipId)) continue;
        if (!uniqueDeletedIds.includes(relationshipId)) continue;

        if (relationship.getAttribute('TargetMode') === 'External') {
          const target = relationship.getAttribute('Target');
          if (target) {
            removedExternalTargets.push(target);
          }
        }

        relationship.parentNode?.removeChild(relationship);
        removedRelationshipIds.push(relationshipId);
      }
      textModifications.set(relationshipsPath, serializeXml(relationships));
    }
  }

  const referenced = collectReferencedInternalParts(zip, textModifications);
  for (const partPath of [...zip.textFiles.keys(), ...zip.binaryFiles.keys()]) {
    if (!isPrunablePart(partPath)) continue;
    if (referenced.has(partPath)) continue;
    removals.add(partPath);
    // Chart relationship parts die with their chart.
    if (partPath.startsWith('ppt/charts/') && partPath.endsWith('.xml')) {
      const chartRels = partPath.replace(/^(.*\/)([^/]+)$/, '$1_rels/$2.rels');
      if (zip.textFiles.has(chartRels)) removals.add(chartRels);
    }
  }

  const contentTypesXml = zip.textFiles.get('[Content_Types].xml');
  if (contentTypesXml && removals.size > 0) {
    const nextContentTypes = stripContentTypeOverrides(contentTypesXml, removals);
    if (nextContentTypes !== contentTypesXml) {
      textModifications.set('[Content_Types].xml', nextContentTypes);
    }
  }

  const buffer = await buildZip(
    packageBuffer,
    textModifications,
    removals.size > 0 ? removals : undefined,
  );
  return {
    buffer,
    removedPartPaths: [...removals].sort(),
    removedRelationshipIds,
    removedExternalTargets,
  };
}
