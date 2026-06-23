import { buildZip, extractZip } from 'pptx-svg';
import {
  SLIDE_CONTENT_TYPE,
  SLIDE_RELATIONSHIP_TYPE,
  getDescendants,
  getDirectChild,
  parseXml,
  serializeXml,
} from './ooxmlXml';

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

  return changed ? serializeXml(exportedDocument) : null;
}

export interface ShapeIdentity {
  cNvPr: Element;
  id: string;
  name: string;
  extensionList: Element | null;
  fingerprint: string;
  shapeKind: string;
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
      if (previous && previous.shapeKind === exported.shapeKind) {
        pairs.push([previous, exported]);
      }
    });
  } else {
    const remaining = [...previousShapes];
    for (const exported of exportedShapes) {
      const matchIndex = remaining.findIndex(
        (candidate) =>
          candidate.shapeKind === exported.shapeKind && candidate.fingerprint === exported.fingerprint
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

function collectShapeIdentities(document: XMLDocument): ShapeIdentity[] {
  return getDescendants(document, 'cNvPr').map((cNvPr) => {
    const shape = cNvPr.parentNode?.parentNode as Element | undefined;
    return {
      cNvPr,
      id: cNvPr.getAttribute('id') ?? '',
      name: cNvPr.getAttribute('name') ?? '',
      extensionList: getDirectChild(cNvPr, 'extLst'),
      fingerprint: getShapeFingerprint(shape),
      shapeKind: shape?.localName ?? ''
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
  return `${shape.localName}|${geometry}|${text}`;
}

function isAnonymizedShapeId(id: string): boolean {
  return id === '' || id === '0';
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
