import { buildZip, extractZip, type ZipContents } from 'pptx-svg';
import {
  getDescendants,
  getElementChildren,
  getShapeChildren,
  getShapeElement,
  getShapeTree,
  nextRelationshipId,
  parseXml,
  serializeXml,
} from './powerpoint/ooxmlXml';

const DRAWING_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const DEFAULT_PASTE_OFFSET_EMU = 228600;

export interface SlideObjectClipboard {
  buffer: ArrayBuffer;
  slideIndex: number;
  shapeIndex: number;
  shapeIndexes: readonly number[];
}

export interface PasteSlideObjectResult {
  buffer: ArrayBuffer;
  shapeIndex: number;
}

export interface PasteSlideObjectsResult {
  buffer: ArrayBuffer;
  shapeIndexes: number[];
}

interface PasteContext {
  source: ZipContents;
  destination: ZipContents;
  textModifications: Map<string, string>;
  binaryModifications: Map<string, Uint8Array>;
  contentTypesDocument: XMLDocument;
  clonedParts: Map<string, string>;
}

export function createSlideObjectClipboard(
  buffer: ArrayBuffer,
  slideIndex: number,
  shapeIndex: number
): SlideObjectClipboard {
  return createSlideObjectsClipboard(buffer, slideIndex, [shapeIndex]);
}

/** Snapshot multiple top-level objects in slide order for one atomic paste. */
export function createSlideObjectsClipboard(
  buffer: ArrayBuffer,
  slideIndex: number,
  shapeIndexes: readonly number[]
): SlideObjectClipboard {
  const normalizedShapeIndexes = normalizeShapeIndexes(shapeIndexes);
  return {
    buffer: buffer.slice(0),
    slideIndex,
    shapeIndex: normalizedShapeIndexes[0]!,
    shapeIndexes: normalizedShapeIndexes,
  };
}

export async function pasteSlideObject(
  destinationBuffer: ArrayBuffer,
  clipboard: SlideObjectClipboard,
  destinationSlideIndex: number,
  offsetEmu = DEFAULT_PASTE_OFFSET_EMU
): Promise<PasteSlideObjectResult> {
  const result = await pasteSlideObjects(
    destinationBuffer,
    {
      ...clipboard,
      shapeIndex: clipboard.shapeIndex,
      shapeIndexes: [clipboard.shapeIndex],
    },
    destinationSlideIndex,
    offsetEmu,
  );
  const [shapeIndex] = result.shapeIndexes;
  if (shapeIndex === undefined) {
    throw new Error('Could not paste the selected PowerPoint object.');
  }
  return { buffer: result.buffer, shapeIndex };
}

/** Paste all copied objects into one OOXML/relationship transaction. */
export async function pasteSlideObjects(
  destinationBuffer: ArrayBuffer,
  clipboard: SlideObjectClipboard,
  destinationSlideIndex: number,
  offsetEmu = DEFAULT_PASTE_OFFSET_EMU
): Promise<PasteSlideObjectsResult> {
  const sourceShapeIndexes = normalizeShapeIndexes(clipboard.shapeIndexes);
  const [source, destination] = await Promise.all([
    extractZip(clipboard.buffer),
    extractZip(destinationBuffer)
  ]);
  const sourceSlidePath = getSlidePath(clipboard.slideIndex);
  const destinationSlidePath = getSlidePath(destinationSlideIndex);
  const sourceSlideDocument = parseXml(
    getRequiredTextFile(source, sourceSlidePath),
    sourceSlidePath
  );
  const destinationSlideDocument = parseXml(
    getRequiredTextFile(destination, destinationSlidePath),
    destinationSlidePath
  );
  const contentTypesDocument = parseXml(
    getRequiredTextFile(destination, '[Content_Types].xml'),
    '[Content_Types].xml'
  );
  const context: PasteContext = {
    source,
    destination,
    textModifications: new Map(),
    binaryModifications: new Map(),
    contentTypesDocument,
    clonedParts: new Map()
  };
  const destinationShapeTree = getShapeTree(destinationSlideDocument);
  const shapeIndexes: number[] = [];
  for (const sourceShapeIndex of sourceShapeIndexes) {
    const clonedShape = destinationSlideDocument.importNode(
      getShapeElement(sourceSlideDocument, sourceShapeIndex),
      true
    );

    offsetShape(clonedShape, offsetEmu, offsetEmu);
    assignUniqueNonVisualIds(destinationSlideDocument, clonedShape);
    await copyShapeRelationships(sourceSlidePath, destinationSlidePath, clonedShape, context);
    destinationShapeTree.appendChild(clonedShape);
    shapeIndexes.push(getShapeChildren(destinationShapeTree).length - 1);
  }

  context.textModifications.set(destinationSlidePath, serializeXml(destinationSlideDocument));
  context.textModifications.set('[Content_Types].xml', serializeXml(contentTypesDocument));
  const buffer = await buildZip(
    destinationBuffer,
    context.textModifications,
    undefined,
    context.binaryModifications
  );
  return { buffer, shapeIndexes };
}

function normalizeShapeIndexes(shapeIndexes: readonly number[]): number[] {
  const normalized = [...new Set(shapeIndexes)]
    .filter((shapeIndex) => Number.isInteger(shapeIndex) && shapeIndex >= 0)
    .sort((left, right) => left - right);
  if (normalized.length === 0) {
    throw new Error('Select at least one PowerPoint object to copy.');
  }
  return normalized;
}

function getRequiredTextFile(zip: ZipContents, partPath: string): string {
  const contents = zip.textFiles.get(partPath);
  if (!contents) throw new Error(`Missing PowerPoint XML part: ${partPath}`);
  return contents;
}

function getSlidePath(slideIndex: number): string {
  return `ppt/slides/slide${slideIndex + 1}.xml`;
}

function offsetShape(shape: Element, dxEmu: number, dyEmu: number): void {
  const transform = getDescendants(shape, 'xfrm')[0];
  const offset = getElementChildren(transform).find((element) => element.localName === 'off');
  if (!offset) return;

  offset.setAttribute('x', String((Number(offset.getAttribute('x')) || 0) + dxEmu));
  offset.setAttribute('y', String((Number(offset.getAttribute('y')) || 0) + dyEmu));
}

function assignUniqueNonVisualIds(destinationSlide: XMLDocument, clonedShape: Element): void {
  const usedIds = new Set(
    getDescendants(destinationSlide, 'cNvPr')
      .filter((element) => !clonedShape.contains(element))
      .map((element) => Number(element.getAttribute('id')))
      .filter(Number.isFinite)
  );
  let nextId = Math.max(0, ...usedIds) + 1;

  for (const nonVisualProperties of getDescendants(clonedShape, 'cNvPr')) {
    while (usedIds.has(nextId)) nextId++;
    nonVisualProperties.setAttribute('id', String(nextId));
    const name = nonVisualProperties.getAttribute('name');
    if (name && !name.endsWith(' Copy')) {
      nonVisualProperties.setAttribute('name', `${name} Copy`);
    }
    usedIds.add(nextId++);
  }

  // A pasted/duplicated shape is a new object, so give it fresh Office creation
  // GUIDs. Reusing the source's <a16:creationId> would leave duplicate GUIDs in the
  // deck (these are meant to uniquely identify each shape across edits/collaboration).
  for (const creationId of getDescendants(clonedShape, 'creationId')) {
    creationId.setAttribute('id', generateOoxmlGuid());
  }
}

function generateOoxmlGuid(): string {
  // Prefer the popout-aware Obsidian window, then the main window. Guard with
  // typeof so the headless test bundle (no DOM globals) falls back cleanly.
  const cryptoApi =
    (typeof activeWindow !== 'undefined' ? activeWindow.crypto : undefined)
    ?? (typeof window !== 'undefined' ? window.crypto : undefined);
  const uuid = cryptoApi?.randomUUID?.() ?? fallbackUuid();
  return `{${uuid.toUpperCase()}}`;
}

function fallbackUuid(): string {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (character) => {
    const random = Math.floor(Math.random() * 16);
    const value = character === 'x' ? random : (random & 0x3) | 0x8;
    return value.toString(16);
  });
}

async function copyShapeRelationships(
  sourceSlidePath: string,
  destinationSlidePath: string,
  clonedShape: Element,
  context: PasteContext
): Promise<void> {
  const sourceRelationshipsPath = getRelationshipsPath(sourceSlidePath);
  const sourceRelationshipsXml = context.source.textFiles.get(sourceRelationshipsPath);
  const relationshipAttributes = getRelationshipAttributes(clonedShape);
  if (!sourceRelationshipsXml || relationshipAttributes.length === 0) return;

  const sourceRelationships = parseXml(sourceRelationshipsXml, sourceRelationshipsPath);
  const destinationRelationshipsPath = getRelationshipsPath(destinationSlidePath);
  const destinationRelationships = context.destination.textFiles.has(destinationRelationshipsPath)
    ? parseXml(getRequiredTextFile(context.destination, destinationRelationshipsPath), destinationRelationshipsPath)
    : createRelationshipsDocument();

  for (const attribute of relationshipAttributes) {
    const sourceRelationship = findRelationship(sourceRelationships, attribute.value);
    if (!sourceRelationship) continue;

    const clonedRelationship = destinationRelationships.importNode(sourceRelationship, true);
    const relationshipId = nextRelationshipId(destinationRelationships);
    clonedRelationship.setAttribute('Id', relationshipId);
    attribute.value = relationshipId;

    const target = clonedRelationship.getAttribute('Target');
    if (target && clonedRelationship.getAttribute('TargetMode') !== 'External') {
      const sourceTargetPath = resolvePartPath(sourceSlidePath, target);
      const destinationTargetPath = await ensureRelatedPart(
        sourceTargetPath,
        context,
        isChartRelationship(clonedRelationship)
      );
      clonedRelationship.setAttribute('Target', getRelativePartPath(destinationSlidePath, destinationTargetPath));
    }

    destinationRelationships.documentElement.appendChild(clonedRelationship);
  }

  context.textModifications.set(destinationRelationshipsPath, serializeXml(destinationRelationships));
}

function isRelationshipAttribute(attribute: Attr): boolean {
  return (
    attribute.namespaceURI === DRAWING_RELATIONSHIP_NAMESPACE
    || attribute.prefix === 'r'
  );
}

function getRelationshipAttributes(element: Element): Attr[] {
  const elements = [element, ...Array.from(element.getElementsByTagName('*'))];
  return elements.flatMap((descendant) =>
    Array.from(descendant.attributes).filter(isRelationshipAttribute)
  );
}

function createRelationshipsDocument(): XMLDocument {
  return parseXml(
    `<Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}"/>`,
    '(new relationships part)'
  );
}

function findRelationship(xmlDocument: XMLDocument, relationshipId: string): Element | null {
  return getDescendants(xmlDocument, 'Relationship')
    .find((relationship) => relationship.getAttribute('Id') === relationshipId)
    ?? null;
}

function isChartRelationship(relationship: Element): boolean {
  return relationship.getAttribute('Type')?.endsWith('/chart') ?? false;
}

async function ensureRelatedPart(
  sourcePartPath: string,
  context: PasteContext,
  forceClone: boolean
): Promise<string> {
  const cached = context.clonedParts.get(sourcePartPath);
  if (cached) return cached;

  const sourceText = context.source.textFiles.get(sourcePartPath);
  const sourceBinary = context.source.binaryFiles.get(sourcePartPath);
  if (sourceText === undefined && sourceBinary === undefined) {
    return sourcePartPath;
  }

  if (!forceClone && hasEquivalentDestinationPart(sourcePartPath, sourceText, sourceBinary, context)) {
    return sourcePartPath;
  }

  const destinationPartPath = getAvailablePartPath(sourcePartPath, context);
  context.clonedParts.set(sourcePartPath, destinationPartPath);
  if (sourceText !== undefined) {
    context.textModifications.set(destinationPartPath, sourceText);
  } else if (sourceBinary) {
    context.binaryModifications.set(destinationPartPath, sourceBinary.slice());
  }
  cloneContentTypeRegistration(sourcePartPath, destinationPartPath, context);

  const sourceRelationshipsPath = getRelationshipsPath(sourcePartPath);
  const sourceRelationshipsXml = context.source.textFiles.get(sourceRelationshipsPath);
  if (sourceRelationshipsXml) {
    const relationships = parseXml(sourceRelationshipsXml, sourceRelationshipsPath);
    for (const relationship of getDescendants(relationships, 'Relationship')) {
      const target = relationship.getAttribute('Target');
      if (!target || relationship.getAttribute('TargetMode') === 'External') continue;

      const sourceTargetPath = resolvePartPath(sourcePartPath, target);
      const destinationTargetPath = await ensureRelatedPart(
        sourceTargetPath,
        context,
        shouldCloneDependency(sourcePartPath, sourceTargetPath)
      );
      relationship.setAttribute('Target', getRelativePartPath(destinationPartPath, destinationTargetPath));
    }
    context.textModifications.set(
      getRelationshipsPath(destinationPartPath),
      serializeXml(relationships)
    );
  }

  return destinationPartPath;
}

function shouldCloneDependency(sourcePartPath: string, sourceTargetPath: string): boolean {
  return (
    sourcePartPath.startsWith('ppt/charts/')
    && (
      sourceTargetPath.startsWith('ppt/charts/')
      || sourceTargetPath.startsWith('ppt/embeddings/')
    )
  );
}

function hasEquivalentDestinationPart(
  partPath: string,
  sourceText: string | undefined,
  sourceBinary: Uint8Array | undefined,
  context: PasteContext
): boolean {
  if (sourceText !== undefined) {
    return context.destination.textFiles.get(partPath) === sourceText;
  }

  const destinationBinary = context.destination.binaryFiles.get(partPath);
  return Boolean(sourceBinary && destinationBinary && sameBytes(sourceBinary, destinationBinary));
}

function getAvailablePartPath(sourcePartPath: string, context: PasteContext): string {
  const standardNumberedPart = sourcePartPath.match(/^(.*?)(\d+)(\.[^./]+)$/);
  if (standardNumberedPart) {
    const prefix = standardNumberedPart[1];
    const suffix = standardNumberedPart[3];
    if (prefix && suffix) {
      let nextNumber = 1;
      while (partPathExists(`${prefix}${nextNumber}${suffix}`, context)) nextNumber++;
      return `${prefix}${nextNumber}${suffix}`;
    }
  }

  const extensionIndex = sourcePartPath.lastIndexOf('.');
  const prefix = extensionIndex === -1 ? sourcePartPath : sourcePartPath.slice(0, extensionIndex);
  const suffix = extensionIndex === -1 ? '' : sourcePartPath.slice(extensionIndex);
  let copyNumber = 1;
  while (partPathExists(`${prefix}-copy-${copyNumber}${suffix}`, context)) copyNumber++;
  return `${prefix}-copy-${copyNumber}${suffix}`;
}

function partPathExists(partPath: string, context: PasteContext): boolean {
  return (
    context.destination.textFiles.has(partPath)
    || context.destination.binaryFiles.has(partPath)
    || context.textModifications.has(partPath)
    || context.binaryModifications.has(partPath)
  );
}

function cloneContentTypeRegistration(
  sourcePartPath: string,
  destinationPartPath: string,
  context: PasteContext
): void {
  const sourceContentTypes = context.source.textFiles.get('[Content_Types].xml');
  if (!sourceContentTypes) return;

  const sourceDocument = parseXml(sourceContentTypes, '[Content_Types].xml');
  const sourceOverride = getDescendants(sourceDocument, 'Override')
    .find((override) => override.getAttribute('PartName') === `/${sourcePartPath}`);
  if (sourceOverride) {
    const destinationPartName = `/${destinationPartPath}`;
    const alreadyRegistered = getDescendants(context.contentTypesDocument, 'Override')
      .some((override) => override.getAttribute('PartName') === destinationPartName);
    if (!alreadyRegistered) {
      const override = context.contentTypesDocument.createElementNS(CONTENT_TYPES_NAMESPACE, 'Override');
      override.setAttribute('PartName', destinationPartName);
      override.setAttribute('ContentType', sourceOverride.getAttribute('ContentType') ?? '');
      context.contentTypesDocument.documentElement.appendChild(override);
    }
  }

  const extension = getPartExtension(destinationPartPath);
  if (!extension) return;
  const alreadyRegistered = getDescendants(context.contentTypesDocument, 'Default')
    .some((entry) => entry.getAttribute('Extension')?.toLowerCase() === extension);
  if (alreadyRegistered) return;

  const sourceDefault = getDescendants(sourceDocument, 'Default')
    .find((entry) => entry.getAttribute('Extension')?.toLowerCase() === extension);
  if (sourceDefault) {
    const entry = context.contentTypesDocument.createElementNS(CONTENT_TYPES_NAMESPACE, 'Default');
    entry.setAttribute('Extension', sourceDefault.getAttribute('Extension') ?? extension);
    entry.setAttribute('ContentType', sourceDefault.getAttribute('ContentType') ?? '');
    context.contentTypesDocument.documentElement.appendChild(entry);
  }
}

function getPartExtension(partPath: string): string {
  return partPath.match(/\.([^./]+)$/)?.[1]?.toLowerCase() ?? '';
}

function getRelationshipsPath(partPath: string): string {
  const slashIndex = partPath.lastIndexOf('/');
  const directory = slashIndex === -1 ? '' : `${partPath.slice(0, slashIndex + 1)}`;
  const fileName = slashIndex === -1 ? partPath : partPath.slice(slashIndex + 1);
  return `${directory}_rels/${fileName}.rels`;
}

function resolvePartPath(sourcePartPath: string, target: string): string {
  const parts = sourcePartPath.split('/');
  parts.pop();

  for (const targetPart of target.replace(/\\/g, '/').split('/')) {
    if (!targetPart || targetPart === '.') continue;
    if (targetPart === '..') {
      parts.pop();
    } else {
      parts.push(targetPart);
    }
  }

  return parts.join('/');
}

function getRelativePartPath(sourcePartPath: string, targetPartPath: string): string {
  const sourceParts = sourcePartPath.split('/');
  sourceParts.pop();
  const targetParts = targetPartPath.split('/');

  while (sourceParts[0] && sourceParts[0] === targetParts[0]) {
    sourceParts.shift();
    targetParts.shift();
  }

  return [...sourceParts.map(() => '..'), ...targetParts].join('/');
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  return left.every((byte, index) => byte === right[index]);
}
