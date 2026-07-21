import { DRAWINGML_NAMESPACE, getElementChildren } from './ooxmlXml';

export type ParagraphListStyle = 'none' | 'bullet' | 'number';

const DEFAULT_LIST_MARGIN_LEFT_EMU = 285750;
const DEFAULT_LIST_HANGING_INDENT_EMU = -285750;
const DEFAULT_BULLET_FONT = 'Arial';

/** Read the explicit DrawingML list marker on one paragraph, if it has one. */
export function getDrawingParagraphListStyle(paragraph: Element): ParagraphListStyle | null {
  const properties = getElementChildren(paragraph).find(
    (element) => element.localName === 'pPr' && element.namespaceURI === DRAWINGML_NAMESPACE,
  );
  if (!properties) return null;

  if (getElementChildren(properties).some((element) => element.localName === 'buChar')) return 'bullet';
  if (getElementChildren(properties).some((element) => element.localName === 'buAutoNum')) return 'number';
  if (getElementChildren(properties).some((element) => element.localName === 'buNone')) return 'none';
  return null;
}

function ensureParagraphProperties(paragraph: Element): Element {
  let properties = getElementChildren(paragraph).find(
    (element) => element.localName === 'pPr' && element.namespaceURI === DRAWINGML_NAMESPACE,
  );
  if (!properties) {
    properties = paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:pPr');
    paragraph.insertBefore(properties, paragraph.firstChild);
  }
  return properties;
}

function clearListMarkers(properties: Element): void {
  for (const child of getElementChildren(properties)) {
    if (child.localName === 'buChar' || child.localName === 'buAutoNum' || child.localName === 'buNone') {
      properties.removeChild(child);
    }
  }
}

function insertParagraphPropertyChild(properties: Element, child: Element): void {
  const tail = getElementChildren(properties).find((element) =>
    element.localName === 'tabLst' || element.localName === 'defRPr' || element.localName === 'extLst',
  );
  properties.insertBefore(child, tail ?? null);
}

function ensureDefaultListIndent(properties: Element): void {
  if (!properties.hasAttribute('marL')) {
    properties.setAttribute('marL', String(DEFAULT_LIST_MARGIN_LEFT_EMU));
  }
  if (!properties.hasAttribute('indent')) {
    properties.setAttribute('indent', String(DEFAULT_LIST_HANGING_INDENT_EMU));
  }
}

function clearDefaultListIndent(properties: Element): void {
  if (properties.getAttribute('marL') === String(DEFAULT_LIST_MARGIN_LEFT_EMU)) {
    properties.removeAttribute('marL');
  }
  if (properties.getAttribute('indent') === String(DEFAULT_LIST_HANGING_INDENT_EMU)) {
    properties.removeAttribute('indent');
  }
}

function ensureBulletFont(properties: Element): void {
  const hasBulletFont = getElementChildren(properties).some(
    (child) => child.localName === 'buFont' || child.localName === 'buFontTx',
  );
  if (hasBulletFont) return;

  const font = properties.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:buFont');
  font.setAttribute('typeface', DEFAULT_BULLET_FONT);
  insertParagraphPropertyChild(properties, font);
}

function clearDefaultBulletFont(properties: Element): void {
  for (const child of getElementChildren(properties)) {
    if (child.localName === 'buFont' && child.getAttribute('typeface') === DEFAULT_BULLET_FONT) {
      properties.removeChild(child);
    }
  }
}

/** Apply a native PowerPoint list marker to a DrawingML paragraph. */
export function applyDrawingParagraphListStyle(paragraph: Element, style: ParagraphListStyle): void {
  const properties = ensureParagraphProperties(paragraph);
  clearListMarkers(properties);

  if (style === 'none') {
    clearDefaultListIndent(properties);
    clearDefaultBulletFont(properties);
    const marker = paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:buNone');
    insertParagraphPropertyChild(properties, marker);
    return;
  }

  ensureDefaultListIndent(properties);

  if (style === 'bullet') {
    ensureBulletFont(properties);
    const marker = paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:buChar');
    marker.setAttribute('char', '•');
    insertParagraphPropertyChild(properties, marker);
    return;
  }

  const marker = paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:buAutoNum');
  marker.setAttribute('type', 'arabicPeriod');
  insertParagraphPropertyChild(properties, marker);
}
