// OOXML package constants and pure XML/part-path helpers extracted from
// PresentationEngine.ts. Shared by the engine, chart formatting, and slide insertions.

export const SLIDE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
export const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
export const DRAWINGML_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PACKAGE_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
export const IMAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
export const SHAPE_ELEMENT_NAMES = new Set(['cxnSp', 'graphicFrame', 'grpSp', 'pic', 'sp']);

export function normalizeLabelText(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

export function parseXml(contents: string, partPath: string): XMLDocument {
  const doc = new DOMParser().parseFromString(contents, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`Could not parse PowerPoint XML part: ${partPath}`);
  }
  return doc;
}

export function serializeXml(doc: XMLDocument): string {
  return new XMLSerializer().serializeToString(doc);
}

export function getDescendants(element: Element | XMLDocument, localName: string): Element[] {
  return Array.from(element.getElementsByTagNameNS('*', localName));
}

export function getElementChildren(element: Element | undefined): Element[] {
  return Array.from(element?.childNodes ?? [])
    .filter((node): node is Element => node.nodeType === 1);
}

export function resolvePartPath(sourcePath: string, target: string): string {
  const parts = sourcePath.split('/');
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

export function getPartExtension(partPath: string): string {
  return partPath.match(/\.([^./]+)$/)?.[1]?.toLowerCase() ?? '';
}

export function imageExtensionForMime(mimeType: string): string {
  switch (mimeType.toLowerCase()) {
    case 'image/jpeg':
      return 'jpeg';
    case 'image/gif':
      return 'gif';
    case 'image/webp':
      return 'webp';
    case 'image/bmp':
      return 'bmp';
    case 'image/png':
    default:
      return 'png';
  }
}

export function contentTypeForImageExtension(extension: string): string {
  switch (extension.toLowerCase()) {
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'webp':
      return 'image/webp';
    case 'bmp':
      return 'image/bmp';
    case 'png':
    default:
      return 'image/png';
  }
}

export function createRelationshipsDocument(): XMLDocument {
  return parseXml(
    `<Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}"/>`,
    '(new relationships part)'
  );
}

export function nextRelationshipId(relationships: XMLDocument): string {
  const used = new Set(
    getDescendants(relationships, 'Relationship')
      .map((relationship) => relationship.getAttribute('Id'))
      .filter((id): id is string => Boolean(id))
  );
  let next = 1;
  while (used.has(`rId${next}`)) next++;
  return `rId${next}`;
}

export function nextImageMediaPath(
  textFiles: Map<string, string>,
  binaryFiles: Map<string, Uint8Array>,
  extension: string
): string {
  let maxIndex = 0;
  const pattern = /^ppt\/media\/image(\d+)\.[^./]+$/;
  for (const key of [...textFiles.keys(), ...binaryFiles.keys()]) {
    const match = key.match(pattern);
    if (match) maxIndex = Math.max(maxIndex, Number(match[1]));
  }
  return `ppt/media/image${maxIndex + 1}.${extension}`;
}

export function ensureDefaultContentType(
  contentTypesDoc: XMLDocument,
  extension: string,
  contentType: string
): void {
  const normalized = extension.toLowerCase();
  const exists = getDescendants(contentTypesDoc, 'Default')
    .some((entry) => entry.getAttribute('Extension')?.toLowerCase() === normalized);
  if (exists) return;

  const namespace = contentTypesDoc.documentElement.namespaceURI;
  const entry = contentTypesDoc.createElementNS(namespace, 'Default');
  entry.setAttribute('Extension', normalized);
  entry.setAttribute('ContentType', contentType);
  contentTypesDoc.documentElement.appendChild(entry);
}

export function getBlipEmbedId(blip: Element): string | null {
  return blip.getAttributeNS(RELATIONSHIP_NAMESPACE, 'embed') || blip.getAttribute('r:embed');
}

export function setBlipEmbedId(blip: Element, relationshipId: string): void {
  const existing = blip.getAttributeNodeNS(RELATIONSHIP_NAMESPACE, 'embed');
  if (existing) {
    existing.value = relationshipId;
  } else {
    blip.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:embed', relationshipId);
  }
}

// Convert an inset crop percentage (0-100) to the OOXML <a:srcRect> unit of
// 1000ths of a percent (0-100000), clamped to the valid range.
export function cropPercentToPermille(percent: number): number {
  if (!Number.isFinite(percent)) return 0;
  return Math.max(0, Math.min(100000, Math.round(percent * 1000)));
}

export function getSlidePath(slideIndex: number): string {
  return `ppt/slides/slide${slideIndex + 1}.xml`;
}

export function getSlideRelationshipsPath(slideIndex: number): string {
  return `ppt/slides/_rels/slide${slideIndex + 1}.xml.rels`;
}

export function hasAncestor(element: Element, localNames: Set<string>): boolean {
  let current = element.parentElement;
  while (current) {
    if (localNames.has(current.localName)) return true;
    current = current.parentElement;
  }
  return false;
}

export function getDirectChild(element: Element, localName: string): Element | null {
  return getElementChildren(element)
    .find((child) => child.localName === localName) ?? null;
}
