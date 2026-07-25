import { buildZip, extractZip, type ZipContents } from 'pptx-svg';
import {
  DRAWINGML_NAMESPACE,
  PACKAGE_RELATIONSHIP_NAMESPACE,
  RELATIONSHIP_NAMESPACE,
  createRelationshipsDocument,
  getDescendants,
  getDirectChild,
  getElementChildren,
  getSlidePath,
  getSlideRelationshipsPath,
  nextRelationshipId,
  parseXml,
  resolvePartPath,
  serializeXml,
} from './ooxmlXml';

const PRESENTATION_NAMESPACE = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const CONTENT_TYPES_NAMESPACE = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NOTES_SLIDE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';
const NOTES_MASTER_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster';
const SLIDE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const THEME_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const NOTES_SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
const NOTES_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const EDITABLE_MAIN_CONTENT_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
]);

export interface SlideNotesReadResult {
  /** Plain text, with OOXML paragraphs and line breaks represented as `\n`. */
  text: string;
  /** The connected notes part, or null when this slide has no speaker notes. */
  notesSlidePath: string | null;
}

export interface SlideNotesWriteResult {
  buffer: ArrayBuffer;
  notesSlidePath: string;
  createdNotesSlide: boolean;
  createdNotesMaster: boolean;
}

/**
 * Read the editable body placeholder from one slide's notes part. This never
 * uses the renderer, because the package is the only lossless source for
 * notes text and its unmodeled OOXML.
 */
export async function readSlideNotesText(
  buffer: ArrayBuffer,
  slideIndex: number,
): Promise<SlideNotesReadResult> {
  assertSlideIndex(slideIndex);
  const zip = await extractZip(buffer);
  const notesSlidePath = resolveNotesSlidePath(zip, slideIndex);
  if (!notesSlidePath) return { text: '', notesSlidePath: null };

  const notesXml = getRequiredTextFile(zip, notesSlidePath);
  const notesDocument = parseXml(notesXml, notesSlidePath);
  const bodyShape = findNotesBodyShape(notesDocument);
  return {
    text: bodyShape ? readBodyPlaceholderText(bodyShape) : '',
    notesSlidePath,
  };
}

/**
 * Update one slide's speaker-notes body without routing the package through
 * the renderer. Existing notes parts, relationships, placeholders, fields,
 * and extension markup are left alone except for the editable body text. When
 * a slide has no notes yet, this creates the same notes-slide/master package
 * graph PowerPoint expects.
 */
export async function writeSlideNotesText(
  buffer: ArrayBuffer,
  slideIndex: number,
  text: string,
): Promise<SlideNotesWriteResult> {
  assertSlideIndex(slideIndex);
  const zip = await extractZip(buffer);
  assertEditableNotesPackage(zip);
  const slidePath = getSlidePath(slideIndex);
  getRequiredTextFile(zip, slidePath);

  const modifications = new Map<string, string>();
  let notesSlidePath = resolveNotesSlidePath(zip, slideIndex);
  let createdNotesSlide = false;
  let createdNotesMaster = false;

  if (notesSlidePath) {
    const notesDocument = parseXml(getRequiredTextFile(zip, notesSlidePath), notesSlidePath);
    const bodyShape = findNotesBodyShape(notesDocument) ?? appendNotesBodyShape(notesDocument);
    writeBodyPlaceholderText(bodyShape, text);
    modifications.set(notesSlidePath, serializeXml(notesDocument));
  } else {
    const notesMaster = ensureNotesMaster(zip, modifications);
    createdNotesMaster = notesMaster.created;
    notesSlidePath = nextNumberedPartPath('ppt/notesSlides/notesSlide', '.xml', zip, modifications);

    const notesDocument = createNotesSlideDocument(text);
    const notesRelationships = createRelationshipsDocument();
    appendRelationship(
      notesRelationships,
      NOTES_MASTER_RELATIONSHIP_TYPE,
      relativePartPath(notesSlidePath, notesMaster.path),
    );
    appendRelationship(
      notesRelationships,
      SLIDE_RELATIONSHIP_TYPE,
      relativePartPath(notesSlidePath, slidePath),
    );

    const slideRelationshipsPath = getSlideRelationshipsPath(slideIndex);
    const slideRelationshipsXml = zip.textFiles.get(slideRelationshipsPath);
    const slideRelationships = slideRelationshipsXml
      ? parseXml(slideRelationshipsXml, slideRelationshipsPath)
      : createRelationshipsDocument();
    appendRelationship(
      slideRelationships,
      NOTES_SLIDE_RELATIONSHIP_TYPE,
      relativePartPath(slidePath, notesSlidePath),
    );

    const contentTypesPath = '[Content_Types].xml';
    const contentTypes = parseXml(
      modifications.get(contentTypesPath) ?? getRequiredTextFile(zip, contentTypesPath),
      contentTypesPath,
    );
    ensureContentTypeOverride(contentTypes, notesSlidePath, NOTES_SLIDE_CONTENT_TYPE);

    modifications.set(notesSlidePath, serializeXml(notesDocument));
    modifications.set(getRelationshipsPath(notesSlidePath), serializeXml(notesRelationships));
    modifications.set(slideRelationshipsPath, serializeXml(slideRelationships));
    modifications.set(contentTypesPath, serializeXml(contentTypes));
    createdNotesSlide = true;
  }

  return {
    buffer: await buildZip(buffer, modifications),
    notesSlidePath,
    createdNotesSlide,
    createdNotesMaster,
  };
}

function resolveNotesSlidePath(zip: ZipContents, slideIndex: number): string | null {
  const slidePath = getSlidePath(slideIndex);
  const relationshipsPath = getSlideRelationshipsPath(slideIndex);
  const relationshipsXml = zip.textFiles.get(relationshipsPath);
  if (!relationshipsXml) return null;

  const relationships = parseXml(relationshipsXml, relationshipsPath);
  const notesRelationship = getRelationshipsByType(relationships, NOTES_SLIDE_RELATIONSHIP_TYPE)[0];
  if (!notesRelationship) return null;
  if (notesRelationship.getAttribute('TargetMode') === 'External') {
    throw new Error(`Slide ${slideIndex + 1} has an external notes relationship, which is not editable.`);
  }

  const target = notesRelationship.getAttribute('Target');
  if (!target) {
    throw new Error(`Slide ${slideIndex + 1} has a notes relationship with no target.`);
  }
  const notesSlidePath = resolvePartPath(slidePath, target);
  if (!zip.textFiles.has(notesSlidePath)) {
    throw new Error(`Slide ${slideIndex + 1} notes part is missing: ${notesSlidePath}.`);
  }
  return notesSlidePath;
}

function ensureNotesMaster(
  zip: ZipContents,
  modifications: Map<string, string>,
): { path: string; created: boolean } {
  const existingPath = [...zip.textFiles.keys()]
    .find((path) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path));
  if (existingPath) return { path: existingPath, created: false };

  const presentationPath = 'ppt/presentation.xml';
  const presentationRelationshipsPath = 'ppt/_rels/presentation.xml.rels';
  const contentTypesPath = '[Content_Types].xml';
  const presentation = parseXml(getRequiredTextFile(zip, presentationPath), presentationPath);
  const presentationRelationships = parseXml(
    getRequiredTextFile(zip, presentationRelationshipsPath),
    presentationRelationshipsPath,
  );
  const contentTypes = parseXml(getRequiredTextFile(zip, contentTypesPath), contentTypesPath);
  const path = nextNumberedPartPath('ppt/notesMasters/notesMaster', '.xml', zip, modifications);
  const relationshipId = appendRelationship(
    presentationRelationships,
    NOTES_MASTER_RELATIONSHIP_TYPE,
    relativePartPath(presentationPath, path),
  );
  appendNotesMasterId(presentation, relationshipId);
  ensureContentTypeOverride(contentTypes, path, NOTES_MASTER_CONTENT_TYPE);

  const themePath = [...zip.textFiles.keys()]
    .find((candidate) => /^ppt\/theme\/[^/]+\.xml$/.test(candidate));
  const notesMasterRelationships = createRelationshipsDocument();
  if (themePath) {
    appendRelationship(
      notesMasterRelationships,
      THEME_RELATIONSHIP_TYPE,
      relativePartPath(path, themePath),
    );
  }

  modifications.set(path, createNotesMasterXml());
  modifications.set(getRelationshipsPath(path), serializeXml(notesMasterRelationships));
  modifications.set(presentationPath, serializeXml(presentation));
  modifications.set(presentationRelationshipsPath, serializeXml(presentationRelationships));
  modifications.set(contentTypesPath, serializeXml(contentTypes));
  return { path, created: true };
}

function appendNotesMasterId(presentation: XMLDocument, relationshipId: string): void {
  const root = presentation.documentElement;
  const existingList = getDirectChild(root, 'notesMasterIdLst');
  const list = existingList ?? presentation.createElementNS(PRESENTATION_NAMESPACE, 'p:notesMasterIdLst');
  const masterId = presentation.createElementNS(PRESENTATION_NAMESPACE, 'p:notesMasterId');
  masterId.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:id', relationshipId);
  list.appendChild(masterId);
  if (existingList) return;

  const firstFollowingList = getElementChildren(root).find((child) =>
    new Set(['handoutMasterIdLst', 'sldIdLst', 'sldSz', 'notesSz']).has(child.localName),
  );
  root.insertBefore(list, firstFollowingList ?? null);
}

function createNotesMasterXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="${DRAWINGML_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}">
  <p:cSld name=""><p:spTree>${groupShapeTreeXml()}</p:spTree></p:cSld>
  <p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>
</p:notesMaster>`;
}

function createNotesSlideDocument(text: string): XMLDocument {
  const document = parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${DRAWINGML_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}">
  <p:cSld name=""><p:spTree>${groupShapeTreeXml()}${notesBodyShapeXml(2)}</p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>`,
    '(new notes slide)',
  );
  const bodyShape = findNotesBodyShape(document);
  if (!bodyShape) throw new Error('Could not create a speaker-notes body placeholder.');
  writeBodyPlaceholderText(bodyShape, text);
  return document;
}

function groupShapeTreeXml(): string {
  return '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    + '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
}

function notesBodyShapeXml(id: number): string {
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="${id}" name="Text Placeholder ${id - 1}"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="685800" y="685800"/><a:ext cx="5486400" cy="4114800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>`;
}

function findNotesBodyShape(document: XMLDocument): Element | null {
  return getDescendants(document, 'sp').find((shape) =>
    getDescendants(shape, 'ph').some((placeholder) =>
      placeholder.namespaceURI === PRESENTATION_NAMESPACE && placeholder.getAttribute('type') === 'body',
    ),
  ) ?? null;
}

function appendNotesBodyShape(document: XMLDocument): Element {
  const shapeTree = getDescendants(document, 'spTree')[0];
  if (!shapeTree) throw new Error('Could not find the notes slide shape tree.');
  const maxId = Math.max(
    1,
    ...getDescendants(document, 'cNvPr')
      .map((properties) => Number(properties.getAttribute('id')))
      .filter((id) => Number.isFinite(id)),
  );
  const template = parseXml(
    `<p:notes xmlns:a="${DRAWINGML_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}">${notesBodyShapeXml(maxId + 1)}</p:notes>`,
    '(new notes body placeholder)',
  );
  const shape = getDescendants(template, 'sp')[0];
  if (!shape) throw new Error('Could not create the notes body placeholder.');
  const imported = document.importNode(shape, true);
  shapeTree.appendChild(imported);
  return imported;
}

function readBodyPlaceholderText(shape: Element): string {
  const textBody = getDescendants(shape, 'txBody')
    .find((element) => element.namespaceURI === PRESENTATION_NAMESPACE);
  if (!textBody) return '';

  return getElementChildren(textBody)
    .filter((child) => child.namespaceURI === DRAWINGML_NAMESPACE && child.localName === 'p')
    .map((paragraph) => readParagraphText(paragraph))
    .join('\n');
}

function readParagraphText(paragraph: Element): string {
  let result = '';
  const visit = (element: Element): void => {
    if (element.namespaceURI === DRAWINGML_NAMESPACE && element.localName === 't') {
      result += element.textContent ?? '';
      return;
    }
    if (element.namespaceURI === DRAWINGML_NAMESPACE && element.localName === 'br') {
      result += '\n';
      return;
    }
    for (const child of getElementChildren(element)) visit(child);
  };
  for (const child of getElementChildren(paragraph)) visit(child);
  return result;
}

function writeBodyPlaceholderText(shape: Element, text: string): void {
  const textBody = getDescendants(shape, 'txBody')
    .find((element) => element.namespaceURI === PRESENTATION_NAMESPACE);
  if (!textBody) throw new Error('The speaker-notes body placeholder has no text body.');

  const paragraphs = getElementChildren(textBody)
    .filter((child) => child.namespaceURI === DRAWINGML_NAMESPACE && child.localName === 'p');
  const template = paragraphs[0] ?? createEmptyParagraph(textBody.ownerDocument);
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  const replacements = lines.map((line) => createNotesParagraph(template, line));
  const insertionPoint = paragraphs[0] ?? getDirectChild(textBody, 'extLst');
  for (const replacement of replacements) textBody.insertBefore(replacement, insertionPoint ?? null);
  for (const paragraph of paragraphs) textBody.removeChild(paragraph);
}

function createEmptyParagraph(document: Document): Element {
  return document.createElementNS(DRAWINGML_NAMESPACE, 'a:p');
}

function createNotesParagraph(template: Element, text: string): Element {
  const paragraph = template.cloneNode(true) as Element;
  const templateRun = getElementChildren(paragraph)
    .find((child) => child.namespaceURI === DRAWINGML_NAMESPACE && child.localName === 'r') ?? null;
  const endParagraphProperties = getElementChildren(paragraph)
    .find((child) => child.namespaceURI === DRAWINGML_NAMESPACE && child.localName === 'endParaRPr') ?? null;

  for (const child of getElementChildren(paragraph)) {
    if (child.namespaceURI === DRAWINGML_NAMESPACE && (child.localName === 'pPr' || child === endParagraphProperties)) {
      continue;
    }
    paragraph.removeChild(child);
  }

  const run = templateRun
    ? templateRun.cloneNode(true) as Element
    : paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:r');
  const textElement = getDescendants(run, 't')
    .find((element) => element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:t');
  if (!textElement.parentNode) run.appendChild(textElement);
  textElement.textContent = text;
  if (/^\s|\s$/u.test(text)) {
    textElement.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  } else {
    textElement.removeAttributeNS(XML_NAMESPACE, 'space');
  }
  paragraph.insertBefore(run, endParagraphProperties);
  return paragraph;
}

function appendRelationship(relationships: XMLDocument, type: string, target: string): string {
  const relationship = relationships.createElementNS(PACKAGE_RELATIONSHIP_NAMESPACE, 'Relationship');
  const id = nextRelationshipId(relationships);
  relationship.setAttribute('Id', id);
  relationship.setAttribute('Type', type);
  relationship.setAttribute('Target', target);
  relationships.documentElement.appendChild(relationship);
  return id;
}

function ensureContentTypeOverride(document: XMLDocument, partPath: string, contentType: string): void {
  const partName = `/${partPath}`;
  const exists = getDescendants(document, 'Override')
    .some((override) => override.getAttribute('PartName') === partName);
  if (exists) return;

  const override = document.createElementNS(CONTENT_TYPES_NAMESPACE, 'Override');
  override.setAttribute('PartName', partName);
  override.setAttribute('ContentType', contentType);
  document.documentElement.appendChild(override);
}

function getRelationshipsByType(document: XMLDocument, type: string): Element[] {
  return getDescendants(document, 'Relationship')
    .filter((relationship) => relationship.getAttribute('Type') === type);
}

function getRelationshipsPath(partPath: string): string {
  const lastSlash = partPath.lastIndexOf('/');
  const directory = partPath.slice(0, lastSlash);
  const name = partPath.slice(lastSlash + 1);
  return `${directory}/_rels/${name}.rels`;
}

function relativePartPath(sourcePath: string, targetPath: string): string {
  const sourceDirectory = sourcePath.split('/').slice(0, -1);
  const target = targetPath.split('/');
  let commonLength = 0;
  while (
    commonLength < sourceDirectory.length
    && commonLength < target.length
    && sourceDirectory[commonLength] === target[commonLength]
  ) {
    commonLength++;
  }
  return [
    ...Array.from({ length: sourceDirectory.length - commonLength }, () => '..'),
    ...target.slice(commonLength),
  ].join('/');
}

function nextNumberedPartPath(
  prefix: string,
  suffix: string,
  zip: ZipContents,
  modifications: Map<string, string>,
): string {
  let index = 1;
  while (
    zip.textFiles.has(`${prefix}${index}${suffix}`)
    || zip.binaryFiles.has(`${prefix}${index}${suffix}`)
    || modifications.has(`${prefix}${index}${suffix}`)
  ) {
    index++;
  }
  return `${prefix}${index}${suffix}`;
}

function getRequiredTextFile(zip: ZipContents, path: string): string {
  const contents = zip.textFiles.get(path);
  if (contents === undefined) throw new Error(`Missing OOXML text part: ${path}.`);
  return contents;
}

function assertSlideIndex(slideIndex: number): void {
  if (!Number.isInteger(slideIndex) || slideIndex < 0) {
    throw new RangeError(`Invalid slide index: ${slideIndex}.`);
  }
}

function assertEditableNotesPackage(zip: ZipContents): void {
  const contentTypesPath = '[Content_Types].xml';
  const contentTypes = parseXml(getRequiredTextFile(zip, contentTypesPath), contentTypesPath);
  const presentationOverride = getDescendants(contentTypes, 'Override')
    .find((override) => override.getAttribute('PartName') === '/ppt/presentation.xml');
  const contentType = presentationOverride?.getAttribute('ContentType');
  if (!contentType || !EDITABLE_MAIN_CONTENT_TYPES.has(contentType)) {
    throw new Error('Speaker notes editing supports modern PPTX and POTX packages only.');
  }
}
