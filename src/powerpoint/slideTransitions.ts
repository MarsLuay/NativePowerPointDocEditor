import {
  getDirectChild,
  getElementChildren,
  parseXml,
  serializeXml,
} from './ooxmlXml';

// Office 2010 added a precise millisecond duration (p14:dur) to the
// ISO/IEC 29500 transition definition. Keep the ISO speed attribute as a
// fallback for older PowerPoint clients.
const PRESENTATIONML_NAMESPACE = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const POWERPOINT_2010_NAMESPACE = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const MARKUP_COMPATIBILITY_NAMESPACE = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const MAX_POWERPOINT_TIME_MS = 2_147_483_647;

const SUPPORTED_EFFECTS = new Set(['cut', 'fade', 'push', 'wipe', 'split']);
const STANDARD_EFFECTS = new Set([
  'blinds', 'checker', 'circle', 'comb', 'cover', 'cut', 'diamond', 'dissolve', 'fade',
  'newsflash', 'plus', 'pull', 'push', 'random', 'randomBar', 'split', 'strips', 'wedge',
  'wheel', 'wipe', 'zoom',
]);

export type SlideTransitionEffect = 'none' | 'cut' | 'fade' | 'push' | 'wipe' | 'split';
export type SlideTransitionSpeed = 'fast' | 'medium' | 'slow';
export type SlideTransitionDirection = 'left' | 'right' | 'up' | 'down';
export type SlideTransitionSplitDirection = 'in' | 'out';
export type SlideTransitionSplitOrientation = 'horizontal' | 'vertical';

export interface NoSlideTransition {
  kind: 'none';
}

// This is the write API. `durationMs` writes p14:dur, while `speed` keeps an
// older Office fallback. Callers do not need to set a speed unless they want
// a particular fallback value.
export interface SlideTransitionSettings {
  kind: Exclude<SlideTransitionEffect, 'none'>;
  direction?: SlideTransitionDirection;
  splitDirection?: SlideTransitionSplitDirection;
  splitOrientation?: SlideTransitionSplitOrientation;
  throughBlack?: boolean;
  durationMs?: number | null;
  advanceAfterMs?: number | null;
  advanceOnClick?: boolean;
  speed?: SlideTransitionSpeed;
}

export type SlideTransitionUpdate = NoSlideTransition | SlideTransitionSettings;

export interface ReadSlideTransition extends Omit<SlideTransitionSettings, 'durationMs' | 'advanceAfterMs' | 'advanceOnClick' | 'speed'> {
  durationMs: number | null;
  advanceAfterMs: number | null;
  advanceOnClick: boolean;
  speed: SlideTransitionSpeed | null;
}

export interface UnknownSlideTransition {
  kind: 'unknown';
  effectName: string;
  durationMs: number | null;
  advanceAfterMs: number | null;
  advanceOnClick: boolean;
  speed: SlideTransitionSpeed | null;
}

export type SlideTransition = NoSlideTransition | ReadSlideTransition | UnknownSlideTransition;

interface TransitionLocation {
  transition: Element;
  container: Element;
}

/** Read the transition that is applied when this slide becomes visible. */
export function readSlideTransition(slideXml: string, partPath = '(slide)'): SlideTransition {
  const document = parseSlideXml(slideXml, partPath);
  const location = findTransition(document.documentElement);
  if (!location) return { kind: 'none' };

  const transition = location.transition;
  const effect = getElementChildren(transition)
    .find((child) => child.localName !== 'sndAc' && child.localName !== 'extLst');
  const timing = readTiming(transition);
  if (!effect) {
    return { kind: 'unknown', effectName: '', ...timing };
  }
  if (!SUPPORTED_EFFECTS.has(effect.localName)) {
    return { kind: 'unknown', effectName: effect.localName, ...timing };
  }

  const kind = effect.localName as Exclude<SlideTransitionEffect, 'none'>;
  const result: ReadSlideTransition = { kind, ...timing };
  if (kind === 'push' || kind === 'wipe') {
    result.direction = sideDirectionFromOoxml(effect.getAttribute('dir'));
  }
  if (kind === 'split') {
    result.splitDirection = splitDirectionFromOoxml(effect.getAttribute('dir'));
    result.splitOrientation = splitOrientationFromOoxml(effect.getAttribute('orient'));
  }
  if (kind === 'cut' || kind === 'fade') {
    result.throughBlack = readBoolean(effect.getAttribute('thruBlk'), false);
  }
  return result;
}

/**
 * Replace only this slide's transition OOXML. Slide timing and slide-level
 * extension markup are left untouched. `kind: 'none'` removes the transition
 * element, matching PowerPoint's "None" transition choice.
 */
export function writeSlideTransition(
  slideXml: string,
  update: SlideTransitionUpdate,
  partPath = '(slide)'
): string {
  const document = parseSlideXml(slideXml, partPath);
  const slide = document.documentElement;
  const existing = findTransition(slide);

  if (update.kind === 'none') {
    if (existing) slide.removeChild(existing.container);
    return serializeXml(document);
  }

  validateTransitionSettings(update);
  const transition = createTransition(document, slide, update, existing?.transition);
  if (existing) {
    if (existing.container === existing.transition) {
      slide.replaceChild(transition, existing.transition);
    } else {
      // An mc:AlternateContent wrapper is transition-specific here. Replacing
      // it avoids leaving a conflicting transition in its fallback branch.
      slide.replaceChild(transition, existing.container);
    }
  } else {
    slide.insertBefore(transition, findTransitionInsertionPoint(slide));
  }

  return serializeXml(document);
}

function parseSlideXml(slideXml: string, partPath: string): XMLDocument {
  const document = parseXml(slideXml, partPath);
  if (document.documentElement.localName !== 'sld') {
    throw new Error(`Expected a PresentationML slide part: ${partPath}`);
  }
  return document;
}

function findTransition(slide: Element): TransitionLocation | null {
  const direct = getDirectChild(slide, 'transition');
  if (direct) return { transition: direct, container: direct };

  // PowerPoint can wrap a p14:dur transition in mc:AlternateContent. Treat
  // that wrapper as one atomic transition so writes cannot leave two effects.
  for (const child of getElementChildren(slide)) {
    if (child.namespaceURI !== MARKUP_COMPATIBILITY_NAMESPACE || child.localName !== 'AlternateContent') continue;
    const transition = getElementChildren(child)
      .flatMap((branch) => getElementChildren(branch))
      .find((candidate) => candidate.localName === 'transition' && candidate.namespaceURI === slide.namespaceURI);
    if (transition) return { transition, container: child };
  }
  return null;
}

function readTiming(transition: Element): Pick<ReadSlideTransition, 'durationMs' | 'advanceAfterMs' | 'advanceOnClick' | 'speed'> {
  const speed = transition.getAttribute('spd');
  return {
    durationMs: parsePowerPointTime(transition.getAttributeNS(POWERPOINT_2010_NAMESPACE, 'dur')),
    advanceAfterMs: parsePowerPointTime(transition.getAttribute('advTm')),
    advanceOnClick: readBoolean(transition.getAttribute('advClick'), true),
    speed: speed === 'fast' || speed === 'med' || speed === 'slow'
      ? speed === 'med' ? 'medium' : speed
      : null,
  };
}

function parsePowerPointTime(value: string | null): number | null {
  if (value === null || !/^\d+$/.test(value)) return null;
  const time = Number(value);
  return Number.isSafeInteger(time) && time <= MAX_POWERPOINT_TIME_MS ? time : null;
}

function readBoolean(value: string | null, fallback: boolean): boolean {
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  return fallback;
}

function sideDirectionFromOoxml(value: string | null): SlideTransitionDirection | undefined {
  switch (value) {
    case 'l': return 'left';
    case 'r': return 'right';
    case 'u': return 'up';
    case 'd': return 'down';
    default: return undefined;
  }
}

function splitDirectionFromOoxml(value: string | null): SlideTransitionSplitDirection | undefined {
  if (value === 'in' || value === 'out') return value;
  return undefined;
}

function splitOrientationFromOoxml(value: string | null): SlideTransitionSplitOrientation | undefined {
  if (value === 'horz') return 'horizontal';
  if (value === 'vert') return 'vertical';
  return undefined;
}

function validateTransitionSettings(update: SlideTransitionSettings): void {
  if (!SUPPORTED_EFFECTS.has(update.kind)) {
    throw new Error(`Unsupported slide transition: ${update.kind}`);
  }
  if ((update.kind === 'push' || update.kind === 'wipe') && !update.direction) {
    throw new Error(`${update.kind} transitions require a direction.`);
  }
  if (update.direction && !['left', 'right', 'up', 'down'].includes(update.direction)) {
    throw new Error(`Unsupported slide transition direction: ${update.direction}`);
  }
  if (update.kind !== 'push' && update.kind !== 'wipe' && update.direction !== undefined) {
    throw new Error(`${update.kind} transitions do not support a side direction.`);
  }
  if (update.kind !== 'split' && (update.splitDirection !== undefined || update.splitOrientation !== undefined)) {
    throw new Error(`${update.kind} transitions do not support split options.`);
  }
  if (update.splitDirection && update.splitDirection !== 'in' && update.splitDirection !== 'out') {
    throw new Error(`Unsupported split direction: ${update.splitDirection}`);
  }
  if (update.splitOrientation && update.splitOrientation !== 'horizontal' && update.splitOrientation !== 'vertical') {
    throw new Error(`Unsupported split orientation: ${update.splitOrientation}`);
  }
  if (update.kind !== 'cut' && update.kind !== 'fade' && update.throughBlack !== undefined) {
    throw new Error(`${update.kind} transitions do not support throughBlack.`);
  }
  if (update.throughBlack !== undefined && typeof update.throughBlack !== 'boolean') {
    throw new Error('throughBlack must be a boolean.');
  }
  if (update.advanceOnClick !== undefined && typeof update.advanceOnClick !== 'boolean') {
    throw new Error('advanceOnClick must be a boolean.');
  }
  validateTime('durationMs', update.durationMs);
  validateTime('advanceAfterMs', update.advanceAfterMs);
  if (update.speed !== undefined && !['fast', 'medium', 'slow'].includes(update.speed)) {
    throw new Error(`Unsupported slide transition speed: ${update.speed}`);
  }
}

function validateTime(name: string, value: number | null | undefined): void {
  if (value === undefined || value === null) return;
  if (!Number.isSafeInteger(value) || value < 0 || value > MAX_POWERPOINT_TIME_MS) {
    throw new Error(`${name} must be an integer from 0 to ${MAX_POWERPOINT_TIME_MS}.`);
  }
}

function createTransition(
  document: XMLDocument,
  slide: Element,
  update: SlideTransitionSettings,
  existing?: Element
): Element {
  const transition = createPresentationElement(document, slide, 'transition');
  const speed = update.speed ?? speedForDuration(update.durationMs);
  transition.setAttribute('spd', speed === 'medium' ? 'med' : speed);
  transition.setAttribute('advClick', update.advanceOnClick === false ? '0' : '1');
  if (update.advanceAfterMs !== undefined && update.advanceAfterMs !== null) {
    transition.setAttribute('advTm', String(update.advanceAfterMs));
  }
  if (update.durationMs !== undefined && update.durationMs !== null) {
    const durationPrefix = namespacePrefix(slide, POWERPOINT_2010_NAMESPACE) ?? 'p14';
    if (!namespacePrefix(slide, POWERPOINT_2010_NAMESPACE)) {
      transition.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${durationPrefix}`, POWERPOINT_2010_NAMESPACE);
    }
    transition.setAttributeNS(POWERPOINT_2010_NAMESPACE, `${durationPrefix}:dur`, String(update.durationMs));
  }

  const effect = createPresentationElement(document, slide, update.kind);
  if (update.kind === 'push' || update.kind === 'wipe') {
    effect.setAttribute('dir', sideDirectionToOoxml(update.direction!));
  } else if (update.kind === 'split') {
    effect.setAttribute('dir', update.splitDirection === 'in' ? 'in' : 'out');
    effect.setAttribute('orient', update.splitOrientation === 'vertical' ? 'vert' : 'horz');
  } else if (update.throughBlack !== undefined) {
    effect.setAttribute('thruBlk', update.throughBlack ? '1' : '0');
  }
  transition.appendChild(effect);
  // Sound and extension settings belong to the transition, but are not part
  // of selecting an effect/direction. PowerPoint keeps them when the effect
  // changes, so preserve them verbatim and in their schema-defined position.
  for (const child of getElementChildren(existing)) {
    if (child.localName !== 'sndAc' && child.localName !== 'extLst') continue;
    transition.appendChild(document.importNode(child, true));
  }
  return transition;
}

function speedForDuration(durationMs: number | null | undefined): SlideTransitionSpeed {
  if (durationMs === undefined || durationMs === null || durationMs <= 500) return 'fast';
  if (durationMs <= 1_000) return 'medium';
  return 'slow';
}

function sideDirectionToOoxml(direction: SlideTransitionDirection): 'l' | 'r' | 'u' | 'd' {
  switch (direction) {
    case 'left': return 'l';
    case 'right': return 'r';
    case 'up': return 'u';
    case 'down': return 'd';
  }
}

function createPresentationElement(document: XMLDocument, slide: Element, localName: string): Element {
  const namespace = slide.namespaceURI || PRESENTATIONML_NAMESPACE;
  const prefix = slide.prefix || '';
  return document.createElementNS(namespace, prefix ? `${prefix}:${localName}` : localName);
}

function findTransitionInsertionPoint(slide: Element): Element | null {
  return getElementChildren(slide)
    .find((child) => child.localName === 'timing' || child.localName === 'extLst') ?? null;
}

function namespacePrefix(element: Element, namespace: string): string | null {
  let current: Element | null = element;
  while (current) {
    for (const attribute of Array.from(current.attributes)) {
      if (attribute.namespaceURI !== XMLNS_NAMESPACE || attribute.value !== namespace) continue;
      if (attribute.localName !== 'xmlns') return attribute.localName;
    }
    current = current.parentElement;
  }
  return null;
}
