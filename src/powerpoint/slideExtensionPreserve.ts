import { buildZip, extractZip } from 'pptx-svg';
import {
  SLIDE_CONTENT_TYPE,
  SLIDE_RELATIONSHIP_TYPE,
  getBlipEmbedId,
  getDescendants,
  getDirectChild,
  getElementChildren,
  parseXml,
  serializeXml,
} from './ooxmlXml';

const KNOWN_OOXML_PREFIXES = new Set(['a', 'c', 'm', 'mc', 'p', 'r']);

export async function normalizeSlideManifest(buffer: ArrayBuffer, slideCount: number): Promise<ArrayBuffer> {
  const zip = await extractZip(buffer);
  const presentationPath = 'ppt/presentation.xml';
  const relationshipsPath = 'ppt/_rels/presentation.xml.rels';
  const contentTypesPath = '[Content_Types].xml';
  const presentation = zip.textFiles.get(presentationPath);
  const relationships = zip.textFiles.get(relationshipsPath);
  const contentTypes = zip.textFiles.get(contentTypesPath);
  if (!presentation || !relationships || !contentTypes) {
    throw new Error('Cannot normalize slide metadata because required OOXML parts are missing.');
  }

  let nextRelId = 1;
  for (const match of relationships.matchAll(/\bId="rId(\d+)"/g)) {
    nextRelId = Math.max(nextRelId, Number(match[1]) + 1);
  }

  const slideIds = Array.from(presentation.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"[^>]*\/?>/g))
    .map((match) => Number(match[1]))
    .filter(Number.isFinite);
  let nextSlideId = Math.max(255, ...slideIds) + 1;
  const normalizedSlideEntries: string[] = [];
  const normalizedRelationships: string[] = [];

  for (let index = 0; index < slideCount; index++) {
    const relationshipId = `rId${nextRelId++}`;
    const slideId = slideIds[index] ?? nextSlideId++;
    normalizedSlideEntries.push(`<p:sldId id="${slideId}" r:id="${relationshipId}"/>`);
    normalizedRelationships.push(
      `<Relationship Id="${relationshipId}" Type="${SLIDE_RELATIONSHIP_TYPE}" Target="slides/slide${index + 1}.xml"/>`
    );
  }

  const updatedPresentation = presentation.replace(
    /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    `<p:sldIdLst>${normalizedSlideEntries.join('')}</p:sldIdLst>`
  );
  const updatedRelationships = relationships
    .replace(
      new RegExp(`<Relationship\\b(?=[^>]*\\bType="${SLIDE_RELATIONSHIP_TYPE}")[^>]*/?>`, 'g'),
      ''
    )
    .replace('</Relationships>', `${normalizedRelationships.join('')}</Relationships>`);
  const slideOverrides = Array.from(
    { length: slideCount },
    (_, index) =>
      `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="${SLIDE_CONTENT_TYPE}"/>`
  ).join('');
  const updatedContentTypes = contentTypes
    .replace(/<Override\b(?=[^>]*\bPartName="\/ppt\/slides\/slide\d+\.xml")[^>]*\/>/g, '')
    .replace('</Types>', `${slideOverrides}</Types>`);

  const removals = new Set<string>();
  for (const path of [...zip.textFiles.keys(), ...zip.binaryFiles.keys()]) {
    const match = path.match(/^ppt\/slides\/(?:_rels\/)?slide(\d+)\.xml(?:\.rels)?$/);
    if (match && Number(match[1]) > slideCount) {
      removals.add(path);
    }
  }

  return buildZip(
    buffer,
    new Map([
      [presentationPath, updatedPresentation],
      [relationshipsPath, updatedRelationships],
      [contentTypesPath, updatedContentTypes]
    ]),
    removals
  );
}

export async function preserveSlideExtensionLists(previousBuffer: ArrayBuffer, exportedBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const [previousZip, exportedZip] = await Promise.all([
    extractZip(previousBuffer),
    extractZip(exportedBuffer)
  ]);
  const modifications = new Map<string, string>();

  for (const [slidePath, exportedXml] of exportedZip.textFiles) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(slidePath)) continue;

    const previousXml = previousZip.textFiles.get(slidePath);
    if (!previousXml) continue;

    const mergedXml = preserveSlideExtensionList(previousXml, exportedXml, slidePath);
    if (mergedXml !== null && mergedXml !== exportedXml) {
      modifications.set(slidePath, mergedXml);
    }
  }

  return modifications.size > 0
    ? buildZip(exportedBuffer, modifications)
    : exportedBuffer;
}

function preserveSlideExtensionList(previousXml: string, exportedXml: string, slidePath: string): string | null {
  const previousDocument = parseXml(previousXml, slidePath);
  const exportedDocument = parseXml(exportedXml, slidePath);
  const previousCommonSlide = getDescendants(previousDocument, 'cSld')[0];
  const exportedCommonSlide = getDescendants(exportedDocument, 'cSld')[0];
  if (!previousCommonSlide || !exportedCommonSlide) return null;

  let changed = false;
  const previousSlideRoot = previousDocument.documentElement;
  const exportedSlideRoot = exportedDocument.documentElement;

  const previousSlideExtensionList = getDirectChild(previousSlideRoot, 'extLst');
  if (previousSlideExtensionList) {
    const exportedSlideExtensionList = getDirectChild(exportedSlideRoot, 'extLst');
    const importedSlideExtensionList = exportedDocument.importNode(previousSlideExtensionList, true);
    if (exportedSlideExtensionList) {
      exportedSlideRoot.replaceChild(importedSlideExtensionList, exportedSlideExtensionList);
    } else {
      exportedSlideRoot.appendChild(importedSlideExtensionList);
    }
    changed = true;
  }

  const previousExtensionList = getDirectChild(previousCommonSlide, 'extLst');
  if (previousExtensionList) {
    const exportedExtensionList = getDirectChild(exportedCommonSlide, 'extLst');
    const importedExtensionList = exportedDocument.importNode(previousExtensionList, true);

    if (exportedExtensionList) {
      exportedCommonSlide.replaceChild(importedExtensionList, exportedExtensionList);
    } else {
      exportedCommonSlide.appendChild(importedExtensionList);
    }
    changed = true;
  }

  if (restoreShapeNonVisualIdentity(previousDocument, exportedDocument)) {
    changed = true;
  }

  if (restoreUnknownNamespaceMarkup(previousDocument, exportedDocument)) {
    changed = true;
  }

  return changed ? serializeXml(exportedDocument) : null;
}

export interface ShapeIdentity {
  cNvPr: Element;
  id: string;
  name: string;
  extensionList: Element | null;
  fingerprint: string;
}

// When the renderer re-serializes a slide whose shapes were mutated, it strips each
// shape's non-visual identity: it resets <p:cNvPr> ids to "0", clears names, and drops
// the per-shape extension list (e.g. the <a16:creationId> that Office writes on every
// shape). Restore that identity from the previous slide so edits do not silently lose
// it (which otherwise fails save validation on virtually every real-world deck).
//
// Each renderer mutation touches a single shape, so when the shape count is unchanged
// the previous and exported shape trees line up by index. When a shape was added or
// removed the counts differ, so unchanged shapes are matched by a geometry + text
// fingerprint instead.
function restoreShapeNonVisualIdentity(previousDocument: XMLDocument, exportedDocument: XMLDocument): boolean {
  const previousShapes = collectShapeIdentities(previousDocument);
  const exportedShapes = collectShapeIdentities(exportedDocument);
  if (previousShapes.length === 0 || exportedShapes.length === 0) return false;

  const pairs: Array<[ShapeIdentity, ShapeIdentity]> = [];
  if (previousShapes.length === exportedShapes.length) {
    exportedShapes.forEach((exported, index) => {
      const previous = previousShapes[index];
      if (previous) {
        pairs.push([previous, exported]);
      }
    });
  } else {
    const remaining = [...previousShapes];
    for (const exported of exportedShapes) {
      const matchIndex = remaining.findIndex(
        (candidate) =>
          candidate.fingerprint === exported.fingerprint
      );
      const previous = matchIndex >= 0 ? remaining[matchIndex] : undefined;
      if (previous) {
        pairs.push([previous, exported]);
        remaining.splice(matchIndex, 1);
      }
    }
  }

  let changed = false;
  for (const [previous, exported] of pairs) {
    if (isAnonymizedShapeId(exported.id) && !isAnonymizedShapeId(previous.id)) {
      exported.cNvPr.setAttribute('id', previous.id);
      if (!exported.name && previous.name) {
        exported.cNvPr.setAttribute('name', previous.name);
      }
      changed = true;
    }
    if (previous.extensionList && !exported.extensionList) {
      exported.cNvPr.appendChild(exportedDocument.importNode(previous.extensionList, true));
      changed = true;
    }
  }

  if (ensureUniqueShapeIds(exportedShapes)) {
    changed = true;
  }

  return changed;
}

function collectShapeIdentities(xmlDocument: XMLDocument): ShapeIdentity[] {
  return getDescendants(xmlDocument, 'cNvPr').map((cNvPr) => {
    const shape = cNvPr.parentNode?.parentNode as Element | undefined;
    return {
      cNvPr,
      id: cNvPr.getAttribute('id') ?? '',
      name: cNvPr.getAttribute('name') ?? '',
      extensionList: getDirectChild(cNvPr, 'extLst'),
      fingerprint: getShapeFingerprint(shape),
    };
  });
}

function getShapeFingerprint(shape: Element | undefined): string {
  if (!shape) return '';
  const transform = getDescendants(shape, 'xfrm')[0];
  const offset = transform ? getDirectChild(transform, 'off') : null;
  const extent = transform ? getDirectChild(transform, 'ext') : null;
  const geometry = [
    offset?.getAttribute('x') ?? '',
    offset?.getAttribute('y') ?? '',
    extent?.getAttribute('cx') ?? '',
    extent?.getAttribute('cy') ?? ''
  ].join(',');
  const text = (shape.textContent ?? '').replace(/\s+/g, ' ').trim().slice(0, 64);
  return `${geometry}|${text}`;
}

function isAnonymizedShapeId(id: string): boolean {
  return id === '' || id === '0';
}

function restoreUnknownNamespaceMarkup(previousDocument: XMLDocument, exportedDocument: XMLDocument): boolean {
  let changed = restoreBlipExtensionMarkup(previousDocument, exportedDocument);

  const pairs = pairShapeRoots(previousDocument, exportedDocument);
  for (const [previousShape, exportedShape] of pairs) {
    if (graftMissingUnknownNamespaceElements(previousShape, exportedShape, exportedDocument)) {
      changed = true;
    }
  }

  // Graft any root-level extensions (e.g. <p:extLst> with unknown elements) that aren't tied to shapes.
  // We skip traversing the shape tree because the loop above already processes them.
  const isSpTree = (element: Element) => element.localName === 'spTree';
  if (previousDocument.documentElement && exportedDocument.documentElement) {
    if (graftMissingUnknownNamespaceElements(previousDocument.documentElement, exportedDocument.documentElement, exportedDocument, isSpTree)) {
      changed = true;
    }
  }

  return changed;
}

function restoreBlipExtensionMarkup(previousDocument: XMLDocument, exportedDocument: XMLDocument): boolean {
  const previousBlips = getDescendants(previousDocument, 'blip');
  const exportedBlips = getDescendants(exportedDocument, 'blip');
  let changed = false;

  for (const previousBlip of previousBlips) {
    const embed = getBlipEmbedId(previousBlip);
    if (!embed) continue;

    const exportedBlip = exportedBlips.find((candidate) => getBlipEmbedId(candidate) === embed);
    if (!exportedBlip) continue;

    const previousExtensionList = getDirectChild(previousBlip, 'extLst');
    if (!previousExtensionList) continue;

    const exportedExtensionList = getDirectChild(exportedBlip, 'extLst');
    if (!exportedExtensionList) {
      const insertBefore = getElementChildren(exportedBlip)[0] ?? null;
      exportedBlip.insertBefore(exportedDocument.importNode(previousExtensionList, true), insertBefore);
      changed = true;
      continue;
    }

    for (const child of getElementChildren(previousExtensionList)) {
      const uri = child.getAttribute('uri');
      const exportedChild = getElementChildren(exportedExtensionList).find(
        (candidate) => candidate.localName === child.localName && candidate.getAttribute('uri') === uri
      );
      if (!exportedChild) {
        exportedExtensionList.appendChild(exportedDocument.importNode(child, true));
        changed = true;
        continue;
      }
      if (graftMissingUnknownNamespaceElements(child, exportedChild, exportedDocument)) {
        changed = true;
      }
    }
  }

  return changed;
}

function pairShapeRoots(previousDocument: XMLDocument, exportedDocument: XMLDocument): Array<[Element, Element]> {
  const previousShapes = collectShapeIdentities(previousDocument);
  const exportedShapes = collectShapeIdentities(exportedDocument);
  if (previousShapes.length === 0 || exportedShapes.length === 0) return [];

  const pairs: Array<[Element, Element]> = [];
  if (previousShapes.length === exportedShapes.length) {
    exportedShapes.forEach((exported, index) => {
      const previous = previousShapes[index];
      if (!previous) return;
      const previousRoot = getShapeRoot(previous);
      const exportedRoot = getShapeRoot(exported);
      if (previous && previousRoot && exportedRoot) {
        pairs.push([previousRoot, exportedRoot]);
      }
    });
    return pairs;
  }

  const remaining = [...previousShapes];
  for (const exported of exportedShapes) {
    const exportedRoot = getShapeRoot(exported);
    if (!exportedRoot) continue;
    const matchIndex = remaining.findIndex(
      (candidate) =>
        candidate.fingerprint === exported.fingerprint
    );
    const previous = matchIndex >= 0 ? remaining[matchIndex] : undefined;
    const previousRoot = previous ? getShapeRoot(previous) : undefined;
    if (previous && previousRoot) {
      pairs.push([previousRoot, exportedRoot]);
      remaining.splice(matchIndex, 1);
    }
  }
  return pairs;
}

function getShapeRoot(identity: ShapeIdentity): Element | null {
  const shape = identity.cNvPr.parentNode?.parentNode;
  return shape && shape.nodeType === 1 ? (shape as Element) : null;
}

function graftMissingUnknownNamespaceElements(
  previousRoot: Element,
  exportedRoot: Element,
  exportedDocument: XMLDocument,
  skipTraversal?: (element: Element) => boolean
): boolean {
  const previousCounts = countUnknownElementNames(previousRoot, skipTraversal);
  const exportedCounts = countUnknownElementNames(exportedRoot, skipTraversal);
  let changed = false;

  for (const [elementName, previousCount] of previousCounts) {
    const exportedCount = exportedCounts.get(elementName) ?? 0;
    const deficit = previousCount - exportedCount;
    if (deficit <= 0) continue;

    const candidates = findUnknownElementsByName(previousRoot, elementName, skipTraversal);
    let grafted = 0;
    for (const candidate of candidates) {
      if (grafted >= deficit) break;
      const parentInPrevious = candidate.parentElement;
      if (!parentInPrevious) continue;
      const parentInExported = findCorrespondingElement(previousRoot, exportedRoot, parentInPrevious);
      if (!parentInExported || elementExistsUnderParent(parentInExported, candidate, skipTraversal)) continue;
      parentInExported.appendChild(exportedDocument.importNode(candidate, true));
      grafted++;
      changed = true;
    }
  }

  return changed;
}

function countUnknownElementNames(root: Element, skipTraversal?: (element: Element) => boolean): Map<string, number> {
  const counts = new Map<string, number>();
  walkElements(root, (element) => {
    const qualifiedName = getQualifiedName(element);
    if (!qualifiedName || !isUnknownNamespaceQName(qualifiedName)) return;
    counts.set(qualifiedName, (counts.get(qualifiedName) ?? 0) + 1);
  }, skipTraversal);
  return counts;
}

function findUnknownElementsByName(root: Element, elementName: string, skipTraversal?: (element: Element) => boolean): Element[] {
  const matches: Element[] = [];
  walkElements(root, (element) => {
    if (getQualifiedName(element) === elementName) {
      matches.push(element);
    }
  }, skipTraversal);
  return matches;
}

function walkElements(root: Element, visit: (element: Element) => void, skipTraversal?: (element: Element) => boolean): void {
  if (skipTraversal?.(root)) return;
  visit(root);
  for (const child of getElementChildren(root)) {
    walkElements(child, visit, skipTraversal);
  }
}

function getQualifiedName(element: Element): string | null {
  const prefix = element.prefix;
  const localName = element.localName;
  if (prefix && localName) {
    return `${prefix}:${localName}`;
  }
  const tag = element.tagName ?? '';
  const match = /^([A-Za-z_][\w.-]*):(.+)$/.exec(tag);
  return match ? `${match[1]}:${match[2]}` : null;
}

function isUnknownNamespaceQName(qualifiedName: string): boolean {
  const prefix = qualifiedName.split(':')[0];
  if (!prefix) return false;
  return !KNOWN_OOXML_PREFIXES.has(prefix);
}

function elementExistsUnderParent(parent: Element, template: Element, skipTraversal?: (element: Element) => boolean): boolean {
  const templateName = getQualifiedName(template);
  if (!templateName) return false;

  const uri = template.getAttribute('uri');
  const embed = template.localName === 'blip' ? getBlipEmbedId(template) : null;
  const val = template.getAttribute('val');

  for (const candidate of findUnknownElementsByName(parent, templateName, skipTraversal)) {
    if (uri && candidate.getAttribute('uri') !== uri) continue;
    if (embed && getBlipEmbedId(candidate) !== embed) continue;
    if (val !== null && candidate.getAttribute('val') !== val) continue;
    if ((candidate.textContent ?? '').trim() === (template.textContent ?? '').trim()) {
      return true;
    }
  }
  return false;
}

interface StructuralPathSegment {
  localName: string;
  indexAmongSameLocalName: number;
  embed?: string | null;
  uri?: string;
}

function findCorrespondingElement(
  previousRoot: Element,
  exportedRoot: Element,
  targetInPrevious: Element
): Element | null {
  if (targetInPrevious === previousRoot) return exportedRoot;
  return resolveStructuralPath(exportedRoot, buildStructuralPath(previousRoot, targetInPrevious));
}

function buildStructuralPath(root: Element, target: Element): StructuralPathSegment[] {
  const segments: StructuralPathSegment[] = [];
  let current: Element | null = target;
  while (current && current !== root) {
    const parent: Element | null = current.parentElement;
    if (!parent) break;
    segments.unshift({
      localName: current.localName,
      indexAmongSameLocalName: indexAmongSameLocalNameSiblings(parent, current),
      embed: current.localName === 'blip' ? getBlipEmbedId(current) : undefined,
      uri: current.localName === 'ext' ? current.getAttribute('uri') ?? undefined : undefined
    });
    current = parent;
  }
  return segments;
}

function indexAmongSameLocalNameSiblings(parent: Element, child: Element): number {
  let index = 0;
  for (const sibling of getElementChildren(parent)) {
    if (sibling.localName !== child.localName) continue;
    if (sibling === child) return index;
    index++;
  }
  return 0;
}

function resolveStructuralPath(root: Element, path: StructuralPathSegment[]): Element | null {
  let current: Element = root;
  for (const segment of path) {
    if (segment.embed && segment.localName === 'blip') {
      const blip = getDescendants(current, 'blip').find((candidate) => getBlipEmbedId(candidate) === segment.embed);
      if (!blip) return null;
      current = blip;
      continue;
    }

    const matches = getElementChildren(current).filter((candidate) => candidate.localName === segment.localName);
    let match = matches[segment.indexAmongSameLocalName];
    if (segment.uri) {
      const byUri = matches.find((candidate) => candidate.getAttribute('uri') === segment.uri);
      match = byUri ?? match;
    }
    if (!match) return null;
    current = match;
  }
  return current;
}

function ensureUniqueShapeIds(shapes: ShapeIdentity[]): boolean {
  const reservedIds = new Set<number>();
  for (const shape of shapes) {
    const numericId = Number(shape.cNvPr.getAttribute('id'));
    if (Number.isInteger(numericId) && numericId > 0) {
      reservedIds.add(numericId);
    }
  }

  const consumedIds = new Set<number>();
  let changed = false;
  let nextId = 1;
  for (const shape of shapes) {
    const numericId = Number(shape.cNvPr.getAttribute('id'));
    if (Number.isInteger(numericId) && numericId > 0 && !consumedIds.has(numericId)) {
      consumedIds.add(numericId);
      continue;
    }
    while (consumedIds.has(nextId) || reservedIds.has(nextId)) nextId++;
    shape.cNvPr.setAttribute('id', String(nextId));
    consumedIds.add(nextId);
    changed = true;
  }
  return changed;
}
