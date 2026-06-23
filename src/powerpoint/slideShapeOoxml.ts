import type { ShapeTransform } from 'pptx-svg';
import {
  DRAWINGML_NAMESPACE,
  SHAPE_ELEMENT_NAMES,
  getDescendants,
  getElementChildren,
} from './ooxmlXml';

export interface ShapeBox {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

function intAttr(value: string | null): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function getShapeElement(slideDoc: XMLDocument, shapeIndex: number): Element {
  const shapeTree = getDescendants(slideDoc, 'spTree')[0];
  const shape = getElementChildren(shapeTree)
    .filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName))[shapeIndex];
  if (!shape) {
    throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
  }
  return shape;
}

export function getSpTreeShapes(shapeTree: Element): Element[] {
  return getElementChildren(shapeTree).filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName));
}

export function getShapeBox(shape: Element): ShapeBox | null {
  const offset = getDescendants(shape, 'off')[0];
  const extent = getDescendants(shape, 'ext')[0];
  if (!offset || !extent) return null;
  return {
    x: intAttr(offset.getAttribute('x')),
    y: intAttr(offset.getAttribute('y')),
    cx: intAttr(extent.getAttribute('cx')),
    cy: intAttr(extent.getAttribute('cy')),
  };
}

export function getShapeTreeElement(slideDoc: XMLDocument): Element {
  const shapeTree = getDescendants(slideDoc, 'spTree')[0];
  if (!shapeTree) {
    throw new Error('Could not find the slide shape tree.');
  }
  return shapeTree;
}

/** Resolve a pptx-svg renderer composite shape index to its OOXML element. */
export function getShapeElementByRendererIndex(slideDoc: XMLDocument, shapeIndex: number): Element {
  const shapes = getSpTreeShapes(getShapeTreeElement(slideDoc));
  if (shapeIndex < 1000) {
    const shape = shapes[shapeIndex];
    if (!shape) {
      throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
    }
    return shape;
  }

  const groupIndex = Math.floor(shapeIndex / 1000);
  const childIndex = shapeIndex % 1000;
  const group = shapes[groupIndex];
  if (!group || group.localName !== 'grpSp') {
    throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
  }

  const children = getElementChildren(group).filter((element) =>
    SHAPE_ELEMENT_NAMES.has(element.localName)
  );
  const child = children[childIndex];
  if (!child) {
    throw new Error(`Could not find slide object ${shapeIndex + 1}.`);
  }
  return child;
}

export function applyTransformToShape(shape: Element, transform: ShapeTransform): boolean {
  const xfrm = getDescendants(shape, 'xfrm')[0];
  if (!xfrm) return false;

  let offset = getElementChildren(xfrm).find(
    (element) => element.localName === 'off' && element.namespaceURI === DRAWINGML_NAMESPACE
  );
  let extent = getElementChildren(xfrm).find(
    (element) => element.localName === 'ext' && element.namespaceURI === DRAWINGML_NAMESPACE
  );
  if (!offset) {
    offset = shape.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:off');
    xfrm.insertBefore(offset, xfrm.firstChild);
  }
  if (!extent) {
    extent = shape.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:ext');
    xfrm.appendChild(extent);
  }

  offset.setAttribute('x', String(Math.round(transform.x)));
  offset.setAttribute('y', String(Math.round(transform.y)));
  extent.setAttribute('cx', String(Math.max(1, Math.round(transform.cx))));
  extent.setAttribute('cy', String(Math.max(1, Math.round(transform.cy))));
  xfrm.setAttribute('rot', String(Math.round(transform.rot)));
  return true;
}

export function nextShapeId(slideDoc: XMLDocument): number {
  let maxId = 1;
  for (const cNvPr of getDescendants(slideDoc, 'cNvPr')) {
    const id = Number(cNvPr.getAttribute('id'));
    if (Number.isFinite(id) && id > maxId) maxId = id;
  }
  return maxId + 1;
}

export function qualifyName(reference: Element, localName: string): string {
  return reference.prefix ? `${reference.prefix}:${localName}` : localName;
}

export function adjacentUnselectedShape(
  element: Element,
  selected: Set<Element>,
  direction: 1 | -1
): Element | null {
  let current = direction === 1 ? element.nextElementSibling : element.previousElementSibling;
  while (current) {
    if (SHAPE_ELEMENT_NAMES.has(current.localName) && !selected.has(current)) return current;
    current = direction === 1 ? current.nextElementSibling : current.previousElementSibling;
  }
  return null;
}
