import { isSVGGElement } from '../domGuards';

const FLIP_WRAPPER_CLASS = 'native-powerpoint-flip-wrapper';

interface SvgFactoryWindow {
  createSvg(tagName: 'g'): SVGGElement;
}

function intDataAttr(element: Element, name: string): number {
  const raw = element.getAttribute(name);
  if (!raw) return 0;
  const value = Number(raw);
  return Number.isFinite(value) ? value : 0;
}

/** Lines/connectors bake flip into endpoint geometry; a wrapper would double-flip. */
export function shapeFlipRenderedInGeometry(shape: Element): boolean {
  const geom = shape.getAttribute('data-ooxml-geom') ?? '';
  return geom === 'line' || /connector/i.test(geom);
}

export function flipTransformForShape(shape: Element, scale: number): string | null {
  const flipH = shape.getAttribute('data-ooxml-flip-h') === '1';
  const flipV = shape.getAttribute('data-ooxml-flip-v') === '1';
  if (!flipH && !flipV) return null;

  const x = intDataAttr(shape, 'data-ooxml-x');
  const y = intDataAttr(shape, 'data-ooxml-y');
  const cx = intDataAttr(shape, 'data-ooxml-cx');
  const cy = intDataAttr(shape, 'data-ooxml-cy');
  const centerX = (x + cx / 2) / scale;
  const centerY = (y + cy / 2) / scale;
  const sx = flipH ? -1 : 1;
  const sy = flipV ? -1 : 1;
  return `translate(${centerX},${centerY}) scale(${sx},${sy}) translate(${-centerX},${-centerY})`;
}

function unwrapFlipWrapper(shape: SVGGElement, wrapper: Element): void {
  while (wrapper.firstChild) {
    shape.insertBefore(wrapper.firstChild, wrapper);
  }
  wrapper.remove();
}

function syncFlipWrapper(shape: SVGGElement, scale: number): void {
  const existing = shape.querySelector(`:scope > g.${FLIP_WRAPPER_CLASS}`);
  const transform = flipTransformForShape(shape, scale);

  if (!transform || shapeFlipRenderedInGeometry(shape)) {
    if (existing) unwrapFlipWrapper(shape, existing);
    return;
  }

  let wrapper = existing as SVGGElement | null;
  if (!wrapper) {
    // Obsidian enhances `Document#createSvg` for HTML helpers. On the live SVG
    // document that helper can try to append a second root node, so use the
    // document window's unattached SVG factory instead.
    const svgWindow = shape.ownerDocument.win as unknown as SvgFactoryWindow;
    wrapper = svgWindow.createSvg('g');
    wrapper.classList.add(FLIP_WRAPPER_CLASS);
    while (shape.firstChild) {
      wrapper.appendChild(shape.firstChild);
    }
    shape.appendChild(wrapper);
  }
  wrapper.setAttribute('transform', transform);
}

/**
 * pptx-svg records flipH/flipV on shape groups but does not mirror pictures,
 * presets, or text in the SVG output. Mirror those shapes here so flip edits
 * are visible while OOXML round-trip stays authoritative.
 */
export function applyShapeFlipTransforms(svg: SVGSVGElement): void {
  const scale = Number(svg.getAttribute('data-ooxml-scale'));
  if (!Number.isFinite(scale) || scale <= 0) return;

  svg.querySelectorAll('g[data-ooxml-shape-idx]').forEach((shape) => {
    if (isSVGGElement(shape)) {
      syncFlipWrapper(shape, scale);
    }
  });
}
