import type { RunStyleChange } from '../PresentationEngine';
import {
  DRAWINGML_NAMESPACE,
  getDescendants,
  getElementChildren,
} from './ooxmlXml';
import {
  applyDrawingParagraphListStyle,
  type ParagraphListStyle,
} from './paragraphListStyle';

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

/**
 * Zero-width space stored in otherwise-empty `<a:t>` nodes.
 *
 * pptx-svg drops empty runs at parse time (`text === ""`) and then skips
 * paragraphs with `runs.length === 0` during SVG emit. Enter at the end of a
 * line creates a native empty sibling `<a:p>`; without this anchor the new
 * paragraph never appears in SVG, so the inline editor cannot rebind to it.
 */
export const EMPTY_PARAGRAPH_RENDER_ANCHOR = '\u200B';

export function isEmptyParagraphRenderAnchorText(text: string): boolean {
  if (text.length === 0) return true;
  for (let index = 0; index < text.length; index++) {
    if (text.charAt(index) !== EMPTY_PARAGRAPH_RENDER_ANCHOR) return false;
  }
  return true;
}

/** Strip render anchors for editor display / logical emptiness checks. */
export function stripEmptyParagraphRenderAnchors(text: string): string {
  return text.split(EMPTY_PARAGRAPH_RENDER_ANCHOR).join('');
}

/** Persist empty paragraph text in a form the renderer will still emit. */
export function toStoredParagraphRunText(text: string): string {
  return text.length === 0 ? EMPTY_PARAGRAPH_RENDER_ANCHOR : text;
}

export interface DrawingParagraphText {
  text: string;
  listStyle: ParagraphListStyle;
  /** When set, forces run bold on/off. When omitted, paragraph 0 keeps template bold; later paragraphs default to not bold. */
  bold?: boolean;
}

export function setDrawingText(container: Element, text: string): void {
  const textElements = getDescendants(container, 't')
    .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  const firstText = textElements[0];
  if (!firstText) {
    const firstParagraph = getDrawingParagraphs(container)[0];
    if (!firstParagraph) {
      throw new Error('This PowerPoint label has no editable text paragraph.');
    }

    const activeDocument = firstParagraph.ownerDocument;
    const run = activeDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:r');
    const textElement = activeDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:t');
    textElement.textContent = text;
    run.appendChild(textElement);
    const endParagraphProperties = getElementChildren(firstParagraph)
      .find((element) => element.localName === 'endParaRPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
    firstParagraph.insertBefore(run, endParagraphProperties ?? null);
    return;
  }

  firstText.textContent = text;
  for (const element of textElements.slice(1)) {
    element.textContent = '';
  }
}

export function getDrawingParagraphs(container: Element): Element[] {
  const textBody = getDescendants(container, 'txBody')
    .find((element) => element.namespaceURI === DRAWINGML_NAMESPACE || element.namespaceURI === container.namespaceURI);
  const scope = textBody ?? container;
  return getElementChildren(scope)
    .filter((element) => element.localName === 'p' && element.namespaceURI === DRAWINGML_NAMESPACE);
}

/**
 * Remove the direct predecessor of a paragraph only when it is a genuinely
 * empty DrawingML paragraph. Empty runs are valid placeholders, but a soft
 * break, field, or non-empty run carries visible content and must be retained.
 */
export function hasEmptyDrawingParagraphBefore(container: Element, paragraphIndex: number): boolean {
  const paragraphs = getDrawingParagraphs(container);
  const paragraph = paragraphs[paragraphIndex];
  const previous = paragraphIndex > 0 ? paragraphs[paragraphIndex - 1] : null;
  if (!paragraph || !previous || previous.parentNode !== paragraph.parentNode) return false;

  const contentChildren = getElementChildren(previous).filter((child) =>
    child.namespaceURI === DRAWINGML_NAMESPACE
      && child.localName !== 'pPr'
      && child.localName !== 'endParaRPr'
  );
  if (contentChildren.some((child) => child.localName !== 'r')) return false;
  if (getDrawingRuns(previous).some((run) => !isEmptyParagraphRenderAnchorText(getDrawingRunText(run)))) {
    return false;
  }

  return true;
}

export function removeEmptyDrawingParagraphBefore(container: Element, paragraphIndex: number): boolean {
  if (!hasEmptyDrawingParagraphBefore(container, paragraphIndex)) return false;
  const previous = getDrawingParagraphs(container)[paragraphIndex - 1];
  if (!previous?.parentNode) return false;

  previous.parentNode.removeChild(previous);
  return true;
}

/**
 * Merge a paragraph into its direct predecessor for Backspace at its start.
 * The predecessor keeps its paragraph-level formatting; every visible child of
 * the current paragraph moves before the predecessor's end-paragraph style so
 * run formatting, fields, and soft breaks survive the join.
 */
export function mergeDrawingParagraphWithPrevious(
  container: Element,
  paragraphIndex: number,
): { merged: boolean; caretOffset: number } {
  const paragraphs = getDrawingParagraphs(container);
  const previous = paragraphIndex > 0 ? paragraphs[paragraphIndex - 1] : null;
  const current = paragraphs[paragraphIndex];
  if (!previous || !current || previous.parentNode !== current.parentNode) {
    return { merged: false, caretOffset: 0 };
  }

  const caretOffset = getDrawingRuns(previous)
    .reduce((offset, run) => offset + getDrawingRunText(run).length, 0);
  const insertionPoint = getElementChildren(previous).find((child) => (
    child.localName === 'endParaRPr' && child.namespaceURI === DRAWINGML_NAMESPACE
  ));
  for (const child of getElementChildren(current)) {
    if (
      child.namespaceURI !== DRAWINGML_NAMESPACE
      || child.localName === 'pPr'
      || child.localName === 'endParaRPr'
    ) {
      continue;
    }
    previous.insertBefore(child, insertionPoint ?? null);
  }
  current.parentNode?.removeChild(current);
  return { merged: true, caretOffset };
}

/** Remove soft breaks whose layout is confined to a single DrawingML paragraph. */
export function removeDrawingParagraphSoftBreaks(paragraph: Element): number {
  let removed = 0;
  for (const child of getElementChildren(paragraph)) {
    if (child.localName === 'br' && child.namespaceURI === DRAWINGML_NAMESPACE) {
      paragraph.removeChild(child);
      removed++;
    }
  }
  return removed;
}

/**
 * Replace a text body's direct paragraphs while retaining the first paragraph's
 * run and end-paragraph styling. Each entry becomes a real `<a:p>` so list
 * markers belong to PowerPoint paragraphs rather than literal glyph text.
 *
 * Heading-first templates often have bold on paragraph 0. Later paragraphs
 * default to not bold unless `bold` is set, so body copy does not inherit the
 * heading weight.
 */
export function replaceDrawingParagraphs(container: Element, paragraphs: readonly DrawingParagraphText[]): void {
  if (paragraphs.length === 0) {
    throw new Error('A PowerPoint text body needs at least one paragraph.');
  }
  if (paragraphs.some(({ text }) => /[\r\n]/.test(text))) {
    throw new Error('Paragraph text cannot contain line breaks; use separate paragraph entries instead.');
  }

  const existing = getDrawingParagraphs(container);
  const template = existing[0];
  if (!template) {
    throw new Error('This PowerPoint label has no editable text paragraph.');
  }
  const parent = template.parentNode;
  if (!parent) {
    throw new Error('Could not find the PowerPoint text body.');
  }

  const replacements = paragraphs.map(({ text, listStyle, bold }, index) => {
    const paragraph = template.cloneNode(true) as Element;
    setDrawingParagraphContents(paragraph, text);
    applyDrawingParagraphListStyle(paragraph, listStyle);
    const resolvedBold = bold ?? (index === 0 ? undefined : false);
    if (resolvedBold !== undefined) {
      applyDrawingParagraphRunBold(paragraph, resolvedBold);
    }
    return paragraph;
  });

  for (const paragraph of replacements) {
    parent.insertBefore(paragraph, template);
  }
  for (const paragraph of existing) {
    parent.removeChild(paragraph);
  }
}

function applyDrawingParagraphRunBold(paragraph: Element, bold: boolean): void {
  const run = getDrawingRuns(paragraph)[0];
  if (!run) return;
  const doc = paragraph.ownerDocument;
  let runProperties = getElementChildren(run).find(
    (element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE,
  );
  if (!runProperties) {
    runProperties = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:rPr');
    run.insertBefore(runProperties, run.firstChild);
  }
  applyRunPropertyChange(runProperties, doc, { bold });
}

function setDrawingParagraphContents(paragraph: Element, text: string): void {
  const templateRun = getDrawingRuns(paragraph)[0] ?? null;
  const endParagraphProperties = getElementChildren(paragraph).find(
    (element) => element.localName === 'endParaRPr' && element.namespaceURI === DRAWINGML_NAMESPACE,
  ) ?? null;
  for (const child of getElementChildren(paragraph)) {
    if (child.localName === 'pPr' || child === endParagraphProperties) continue;
    paragraph.removeChild(child);
  }

  const run = templateRun
    ? cloneDrawingRun(templateRun, paragraph.ownerDocument)
    : paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:r');
  let textElement = getElementChildren(run)
    .find((element) => element.localName === 't' && element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? getDescendants(run, 't').find((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  if (!textElement) {
    textElement = paragraph.ownerDocument.createElementNS(DRAWINGML_NAMESPACE, 'a:t');
    run.appendChild(textElement);
  }
  textElement.textContent = text;
  paragraph.insertBefore(run, endParagraphProperties);
}

function clearDrawingParagraphContent(paragraph: Element): Element | null {
  let templateRun: Element | null = null;
  for (const child of getElementChildren(paragraph)) {
    if (child.localName === 'pPr' && child.namespaceURI === DRAWINGML_NAMESPACE) continue;
    if (child.localName === 'r' && child.namespaceURI === DRAWINGML_NAMESPACE && !templateRun) {
      templateRun = child;
    }
    paragraph.removeChild(child);
  }
  return templateRun;
}

function appendDrawingParagraphRun(paragraph: Element, templateRun: Element | null, text: string): void {
  const doc = paragraph.ownerDocument;
  const run = templateRun ? cloneDrawingRun(templateRun, doc) : doc.createElementNS(DRAWINGML_NAMESPACE, 'a:r');
  setDrawingRunText(run, text);
  paragraph.appendChild(run);
}

export function setDrawingParagraphText(container: Element, paragraphIndex: number, text: string): void {
  const paragraphs = getDrawingParagraphs(container);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new Error('Could not find the selected text paragraph.');
  }

  if (!text.includes('\n')) {
    removeDrawingParagraphSoftBreaks(paragraph);
    let runs = getDrawingRuns(paragraph);
    if (runs.length === 0) {
      // Empty template paragraphs can contain only <a:endParaRPr>. Keep that
      // paragraph and its list properties, but create the run needed to write
      // the requested text instead of rejecting an otherwise valid edit.
      ensureDrawingParagraphRun(paragraph, null);
      runs = getDrawingRuns(paragraph);
    }

    // The inline SVG preview keeps text distributed over its existing runs so
    // each segment retains its own font and direct formatting. Commit must use
    // that same ownership model: writing the complete paragraph into run zero
    // made every edited multi-font paragraph adopt the first run's font.
    const nextRunTexts = redistributeDrawingRunText(
      runs.map((run) => getDrawingRunText(run)),
      text,
    );
    runs.forEach((run, runIndex) => {
      const nextText = nextRunTexts[runIndex] ?? '';
      setDrawingRunText(run, text.length === 0 && runIndex === 0
        ? toStoredParagraphRunText('')
        : nextText);
    });
    return;
  }

  const doc = paragraph.ownerDocument;
  const templateRun = clearDrawingParagraphContent(paragraph);
  const lines = text.split('\n');
  lines.forEach((line, lineIndex) => {
    if (lineIndex > 0) {
      paragraph.appendChild(doc.createElementNS(DRAWINGML_NAMESPACE, 'a:br'));
    }
    appendDrawingParagraphRun(paragraph, templateRun, toStoredParagraphRunText(line));
  });
}

/**
 * Redistribute edited paragraph text over its original OOXML runs without
 * changing run order or style ownership. Insertions at a boundary belong to
 * the following run (except at the end, where they belong to the final run),
 * matching the live SVG preview and PowerPoint's expected typing behavior.
 */
function redistributeDrawingRunText(previousRunTexts: readonly string[], nextText: string): string[] {
  if (previousRunTexts.length === 0) return [];

  const previousText = previousRunTexts.join('');
  if (previousText === nextText) return [...previousRunTexts];

  let commonPrefixLength = 0;
  const sharedLength = Math.min(previousText.length, nextText.length);
  while (
    commonPrefixLength < sharedLength
    && previousText.charAt(commonPrefixLength) === nextText.charAt(commonPrefixLength)
  ) {
    commonPrefixLength++;
  }

  let commonSuffixLength = 0;
  while (
    commonSuffixLength < sharedLength - commonPrefixLength
    && previousText.charAt(previousText.length - 1 - commonSuffixLength)
      === nextText.charAt(nextText.length - 1 - commonSuffixLength)
  ) {
    commonSuffixLength++;
  }

  const previousChangeStart = commonPrefixLength;
  const previousChangeEnd = previousText.length - commonSuffixLength;
  const nextChangeEnd = nextText.length - commonSuffixLength;
  const delta = nextText.length - previousText.length;
  const boundaries = [0];
  for (const runText of previousRunTexts) {
    boundaries.push((boundaries[boundaries.length - 1] ?? 0) + runText.length);
  }

  const ownerIndex = previousRunTexts.findIndex(
    (_, index) => (boundaries[index + 1] ?? 0) > previousChangeStart,
  );
  const replacementOwnerIndex = ownerIndex === -1 ? previousRunTexts.length - 1 : ownerIndex;
  const isInsertion = previousChangeStart === previousChangeEnd;
  const mapBoundary = (boundary: number, boundaryIndex: number): number => {
    if (isInsertion) {
      if (boundary < previousChangeStart) return boundary;
      if (boundary > previousChangeStart) return boundary + delta;
      return boundaryIndex > replacementOwnerIndex ? nextChangeEnd : boundary;
    }
    if (boundary <= previousChangeStart) return boundary;
    if (boundary >= previousChangeEnd) return boundary + delta;
    return boundaryIndex <= replacementOwnerIndex ? previousChangeStart : nextChangeEnd;
  };

  const nextBoundaries = boundaries.map(mapBoundary);
  return previousRunTexts.map((_, index) =>
    nextText.slice(nextBoundaries[index] ?? 0, nextBoundaries[index + 1] ?? nextText.length),
  );
}

export interface DrawingParagraphFontSummary {
  runCount: number;
  nonEmptyRunCount: number;
  fontFamilies: string[];
  fontSizesPt: number[];
}

/** Compact, content-free font diagnostics for an edited DrawingML paragraph. */
export function getDrawingParagraphFontSummary(paragraph: Element | null | undefined): DrawingParagraphFontSummary | null {
  if (!paragraph) return null;

  const runs = getDrawingRuns(paragraph);
  const fontFamilies = new Set<string>();
  const fontSizesPt = new Set<number>();
  for (const run of runs) {
    const rPr = getElementChildren(run).find((child) => (
      child.namespaceURI === DRAWINGML_NAMESPACE && child.localName === 'rPr'
    ));
    const latin = rPr
      ? getElementChildren(rPr).find((child) => (
        child.namespaceURI === DRAWINGML_NAMESPACE && child.localName === 'latin'
      ))
      : null;
    const typeface = latin?.getAttribute('typeface');
    if (typeface) fontFamilies.add(typeface);
    const size = Number(rPr?.getAttribute('sz'));
    if (Number.isFinite(size) && size > 0) fontSizesPt.add(size / 100);
  }

  return {
    runCount: runs.length,
    nonEmptyRunCount: runs.filter((run) => !isEmptyParagraphRenderAnchorText(getDrawingRunText(run))).length,
    fontFamilies: [...fontFamilies].slice(0, 6),
    fontSizesPt: [...fontSizesPt].sort((left, right) => left - right).slice(0, 6),
  };
}

function findEndParagraphProperties(paragraph: Element): Element | null {
  return getElementChildren(paragraph).find(
    (element) => element.localName === 'endParaRPr' && element.namespaceURI === DRAWINGML_NAMESPACE,
  ) ?? null;
}

function insertDrawingRunBeforeEndParagraphProperties(paragraph: Element, run: Element): void {
  paragraph.insertBefore(run, findEndParagraphProperties(paragraph));
}

function findNearestDrawingParagraphRun(paragraph: Element): Element | null {
  const parent = paragraph.parentNode;
  if (!parent || parent.nodeType !== 1) return null;
  const paragraphs = getElementChildren(parent as Element).filter(
    (child) => child.localName === 'p' && child.namespaceURI === DRAWINGML_NAMESPACE,
  );
  const paragraphIndex = paragraphs.indexOf(paragraph);
  if (paragraphIndex === -1) return null;

  for (let index = paragraphIndex - 1; index >= 0; index--) {
    const candidate = paragraphs[index];
    const run = candidate ? getDrawingRuns(candidate)[0] : null;
    if (run) return run;
  }
  for (let index = paragraphIndex + 1; index < paragraphs.length; index++) {
    const candidate = paragraphs[index];
    const run = candidate ? getDrawingRuns(candidate)[0] : null;
    if (run) return run;
  }
  return null;
}

function ensureDrawingParagraphRun(paragraph: Element, templateRun: Element | null): void {
  if (getDrawingRuns(paragraph).length > 0) return;
  const doc = paragraph.ownerDocument;
  const styleTemplate = templateRun ?? findNearestDrawingParagraphRun(paragraph);
  const run = styleTemplate ? cloneDrawingRun(styleTemplate, doc) : doc.createElementNS(DRAWINGML_NAMESPACE, 'a:r');
  let textElement = getElementChildren(run)
    .find((element) => element.localName === 't' && element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? getDescendants(run, 't').find((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  if (!textElement) {
    textElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:t');
    run.appendChild(textElement);
  }
  textElement.textContent = EMPTY_PARAGRAPH_RENDER_ANCHOR;
  insertDrawingRunBeforeEndParagraphProperties(paragraph, run);
}

/**
 * Split one real DrawingML paragraph at its run-text offset. The sibling keeps
 * the source paragraph's pPr (including native list style), while the suffix
 * retains each run's direct formatting. This is the normal PowerPoint Enter
 * behavior; soft breaks belong to `setDrawingParagraphText` instead.
 */
export function splitDrawingParagraphAtOffset(
  container: Element,
  paragraphIndex: number,
  offset: number,
): number {
  const paragraphs = getDrawingParagraphs(container);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new Error('Could not find the selected text paragraph.');
  }
  const parent = paragraph.parentNode;
  if (!parent) {
    throw new Error('Could not find the PowerPoint text body.');
  }

  const doc = paragraph.ownerDocument;
  const templateRun = getDrawingRuns(paragraph)[0] ?? null;
  const segments = getDrawingRunSegments(paragraph);
  const total = segments.at(-1)?.end ?? 0;
  const splitOffset = Math.max(0, Math.min(total, offset));
  splitParagraphAtOffset(paragraph, splitOffset, doc);

  const splitSegments = getDrawingRunSegments(paragraph);
  const firstSuffixRunIndex = splitSegments.findIndex((segment) => segment.start >= splitOffset);
  const suffixRuns = firstSuffixRunIndex === -1
    ? []
    : getDrawingRuns(paragraph).slice(firstSuffixRunIndex);

  const suffixParagraph = paragraph.cloneNode(true) as Element;
  for (const child of getElementChildren(suffixParagraph)) {
    if (
      child.localName === 'pPr'
      || (child.localName === 'endParaRPr' && child.namespaceURI === DRAWINGML_NAMESPACE)
    ) {
      continue;
    }
    suffixParagraph.removeChild(child);
  }

  for (const run of suffixRuns) {
    insertDrawingRunBeforeEndParagraphProperties(suffixParagraph, cloneDrawingRun(run, doc));
    paragraph.removeChild(run);
  }

  ensureDrawingParagraphRun(paragraph, templateRun);
  ensureDrawingParagraphRun(suffixParagraph, templateRun);
  parent.insertBefore(suffixParagraph, paragraph.nextSibling);
  return paragraphIndex + 1;
}

/** Outcome of isolating selected text as a native PowerPoint list paragraph. */
export interface DrawingParagraphListStyleRangeResult {
  changed: boolean;
  sourceParagraphIndex: number;
  selectedParagraphIndex: number;
  selectedRange: DrawingTextRange;
  beforeParagraphCount: number;
  afterParagraphCount: number;
  splitPrefix: boolean;
  splitSuffix: boolean;
}

/**
 * Make one non-empty run-text range into its own DrawingML paragraph, then
 * apply the requested native list marker to that paragraph only. Splitting the
 * suffix before the prefix keeps the source paragraph index stable and retains
 * each cloned run's direct formatting on all three resulting paragraphs.
 */
export function applyDrawingParagraphListStyleToRange(
  container: Element,
  paragraphIndex: number,
  startOffset: number,
  endOffset: number,
  style: ParagraphListStyle,
): DrawingParagraphListStyleRangeResult {
  const paragraphs = getDrawingParagraphs(container);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new Error('Could not find the selected text paragraph.');
  }

  const total = paragraphTextLength(paragraph);
  const start = Math.max(0, Math.min(total, Math.floor(startOffset)));
  const end = Math.max(start, Math.min(total, Math.floor(endOffset)));
  if (end <= start) {
    throw new Error('Select text before applying a list style.');
  }

  // `splitDrawingParagraphAtOffset` preserves run formatting, but a soft break
  // or field cannot yet be partitioned at a run-only character offset. Reject
  // that ambiguous structural edit rather than moving it to the wrong list
  // paragraph or silently deleting it.
  const needsInteriorSplit = start > 0 || end < total;
  const hasNonRunContent = getElementChildren(paragraph).some((child) => (
    child.namespaceURI === DRAWINGML_NAMESPACE
      && child.localName !== 'pPr'
      && child.localName !== 'endParaRPr'
      && child.localName !== 'r'
  ));
  if (needsInteriorSplit && hasNonRunContent) {
    throw new Error('Cannot apply a list style to part of a paragraph containing a soft break or field.');
  }

  const beforeParagraphCount = paragraphs.length;
  const splitSuffix = end < total;
  if (splitSuffix) {
    splitDrawingParagraphAtOffset(container, paragraphIndex, end);
  }

  const splitPrefix = start > 0;
  const selectedParagraphIndex = splitPrefix
    ? splitDrawingParagraphAtOffset(container, paragraphIndex, start)
    : paragraphIndex;
  const selectedParagraph = getDrawingParagraphs(container)[selectedParagraphIndex];
  if (!selectedParagraph) {
    throw new Error('Could not isolate the selected text paragraph.');
  }

  applyDrawingParagraphListStyle(selectedParagraph, style);
  return {
    changed: true,
    sourceParagraphIndex: paragraphIndex,
    selectedParagraphIndex,
    selectedRange: {
      paragraphIndex: selectedParagraphIndex,
      start: 0,
      end: end - start,
    },
    beforeParagraphCount,
    afterParagraphCount: getDrawingParagraphs(container).length,
    splitPrefix,
    splitSuffix,
  };
}

export function setDrawingTextRun(container: Element, paragraphIndex: number, runIndex: number, text: string): void {
  const paragraphs = getDrawingParagraphs(container);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new Error('Could not find the selected text paragraph.');
  }

  const runs = getElementChildren(paragraph)
    .filter((element) => element.localName === 'r' && element.namespaceURI === DRAWINGML_NAMESPACE);
  const run = runs[runIndex];
  if (!run) {
    throw new Error('Could not find the selected text run.');
  }

  const textElement = getElementChildren(run)
    .find((element) => element.localName === 't' && element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? getDescendants(run, 't').find((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  if (!textElement) {
    throw new Error('Could not find the selected text node.');
  }

  textElement.textContent = text;
}

export function getDrawingRunText(run: Element): string {
  const textElement = getElementChildren(run)
    .find((element) => element.localName === 't' && element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? getDescendants(run, 't').find((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  return textElement?.textContent ?? '';
}

function setDrawingRunText(run: Element, text: string): void {
  const textElement = getElementChildren(run)
    .find((element) => element.localName === 't' && element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? getDescendants(run, 't').find((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  if (!textElement) {
    throw new Error('Could not find the selected text node.');
  }
  textElement.textContent = text;
}

function cloneDrawingRun(run: Element, _doc: XMLDocument): Element {
  return run.cloneNode(true) as Element;
}

interface DrawingRunSegment {
  run: Element;
  runIndex: number;
  start: number;
  end: number;
  text: string;
}

/** A character range inside one DrawingML paragraph. */
export interface DrawingTextRange {
  paragraphIndex: number;
  start: number;
  end: number;
}

/** Outcome of deleting one or more text ranges from a shape. */
export interface DrawingTextRangeDeletionResult {
  changed: boolean;
  paragraphIndex: number;
  caretOffset: number;
  deletedRangeCount: number;
  removedParagraphCount: number;
  mergedParagraphs: boolean;
}

function getDrawingRunSegments(paragraph: Element): DrawingRunSegment[] {
  const segments: DrawingRunSegment[] = [];
  let offset = 0;
  getDrawingRuns(paragraph).forEach((run, runIndex) => {
    const text = getDrawingRunText(run);
    segments.push({ run, runIndex, start: offset, end: offset + text.length, text });
    offset += text.length;
  });
  return segments;
}

function paragraphTextLength(paragraph: Element): number {
  return getDrawingRunSegments(paragraph).at(-1)?.end ?? 0;
}

function mergeDrawingRanges(ranges: readonly DrawingTextRange[]): DrawingTextRange[] {
  const merged: DrawingTextRange[] = [];
  for (const range of [...ranges].sort((left, right) => (
    left.paragraphIndex - right.paragraphIndex || left.start - right.start || left.end - right.end
  ))) {
    const previous = merged.at(-1);
    if (previous && previous.paragraphIndex === range.paragraphIndex && range.start <= previous.end) {
      previous.end = Math.max(previous.end, range.end);
    } else {
      merged.push({ ...range });
    }
  }
  return merged;
}

function deleteDrawingParagraphTextRange(paragraph: Element, start: number, end: number): boolean {
  let changed = false;
  for (const segment of getDrawingRunSegments(paragraph)) {
    const localStart = Math.max(0, Math.min(segment.text.length, start - segment.start));
    const localEnd = Math.max(0, Math.min(segment.text.length, end - segment.start));
    if (localEnd <= localStart) continue;
    const nextText = segment.text.slice(0, localStart) + segment.text.slice(localEnd);
    setDrawingRunText(segment.run, nextText);
    changed = true;
  }
  return changed;
}

/**
 * Convert an imported literal bullet prefix into native list formatting without
 * leaving a second visible marker in the first run. Leading indentation stays
 * intact and direct run formatting is preserved.
 */
export function stripDrawingParagraphLeadingBullet(paragraph: Element): boolean {
  const text = getDrawingRuns(paragraph).map((run) => getDrawingRunText(run)).join('');
  const match = text.match(/^([\t \u00A0]*)[•◦▪‣⁃][\t \u00A0]+/);
  if (!match) return false;
  const leadingIndentLength = match[1]?.length ?? 0;
  return deleteDrawingParagraphTextRange(paragraph, leadingIndentLength, match[0].length);
}

function retainDrawingParagraphPrefix(paragraph: Element, offset: number): void {
  for (const segment of getDrawingRunSegments(paragraph)) {
    const retainedLength = Math.max(0, Math.min(segment.text.length, offset - segment.start));
    setDrawingRunText(segment.run, segment.text.slice(0, retainedLength));
  }
}

function cloneDrawingParagraphSuffix(paragraph: Element, offset: number): Element[] {
  const clones: Element[] = [];
  for (const segment of getDrawingRunSegments(paragraph)) {
    const localStart = Math.max(0, Math.min(segment.text.length, offset - segment.start));
    if (localStart >= segment.text.length) continue;
    const clone = segment.run.cloneNode(true) as Element;
    setDrawingRunText(clone, segment.text.slice(localStart));
    clones.push(clone);
  }
  return clones;
}

/**
 * Delete selected text from one shape in a single OOXML mutation. A continuous
 * selection across paragraphs joins the remaining prefix and suffix in the
 * first paragraph while retaining the original runs and their formatting.
 */
export function deleteDrawingTextRanges(
  container: Element,
  requestedRanges: readonly DrawingTextRange[],
): DrawingTextRangeDeletionResult {
  const paragraphs = getDrawingParagraphs(container);
  const ranges = mergeDrawingRanges(requestedRanges.flatMap((range) => {
    if (!Number.isFinite(range.paragraphIndex) || !Number.isFinite(range.start) || !Number.isFinite(range.end)) {
      return [];
    }
    const paragraph = paragraphs[range.paragraphIndex];
    if (!paragraph) return [];
    const length = paragraphTextLength(paragraph);
    const start = Math.max(0, Math.min(length, Math.floor(range.start)));
    const end = Math.max(start, Math.min(length, Math.floor(range.end)));
    return end > start ? [{ paragraphIndex: range.paragraphIndex, start, end }] : [];
  }));
  const first = ranges[0];
  if (!first) {
    return {
      changed: false,
      paragraphIndex: 0,
      caretOffset: 0,
      deletedRangeCount: 0,
      removedParagraphCount: 0,
      mergedParagraphs: false,
    };
  }
  const last = ranges.at(-1) ?? first;
  const spansParagraphs = first.paragraphIndex !== last.paragraphIndex;
  const continuousParagraphSpan = spansParagraphs
    && ranges.length === last.paragraphIndex - first.paragraphIndex + 1
    && ranges.every((range, index) => range.paragraphIndex === first.paragraphIndex + index)
    && first.end === paragraphTextLength(paragraphs[first.paragraphIndex]!)
    && last.start === 0
    && ranges.slice(1, -1).every((range) => {
      const length = paragraphTextLength(paragraphs[range.paragraphIndex]!);
      return range.start === 0 && range.end === length;
    });

  if (continuousParagraphSpan) {
    const firstParagraph = paragraphs[first.paragraphIndex]!;
    const lastParagraph = paragraphs[last.paragraphIndex]!;
    const suffixRuns = cloneDrawingParagraphSuffix(lastParagraph, last.end);
    retainDrawingParagraphPrefix(firstParagraph, first.start);
    const endProperties = getElementChildren(firstParagraph).find((child) => (
      child.namespaceURI === DRAWINGML_NAMESPACE && child.localName === 'endParaRPr'
    ));
    for (const run of suffixRuns) {
      firstParagraph.insertBefore(run, endProperties ?? null);
    }
    for (let index = last.paragraphIndex; index > first.paragraphIndex; index--) {
      const paragraph = paragraphs[index];
      paragraph?.parentNode?.removeChild(paragraph);
    }
    return {
      changed: true,
      paragraphIndex: first.paragraphIndex,
      caretOffset: first.start,
      deletedRangeCount: ranges.length,
      removedParagraphCount: last.paragraphIndex - first.paragraphIndex,
      mergedParagraphs: true,
    };
  }

  let changed = false;
  for (const range of [...ranges].reverse()) {
    const paragraph = paragraphs[range.paragraphIndex];
    if (paragraph) changed = deleteDrawingParagraphTextRange(paragraph, range.start, range.end) || changed;
  }
  return {
    changed,
    paragraphIndex: first.paragraphIndex,
    caretOffset: first.start,
    deletedRangeCount: ranges.length,
    removedParagraphCount: 0,
    mergedParagraphs: false,
  };
}

function splitDrawingRunAt(
  paragraph: Element,
  runIndex: number,
  localOffset: number,
  doc: XMLDocument
): void {
  const runs = getDrawingRuns(paragraph);
  const run = runs[runIndex];
  if (!run) return;

  const text = getDrawingRunText(run);
  if (localOffset <= 0 || localOffset >= text.length) return;

  setDrawingRunText(run, text.slice(0, localOffset));
  const afterRun = cloneDrawingRun(run, doc);
  setDrawingRunText(afterRun, text.slice(localOffset));

  const next = run.nextSibling;
  if (next) {
    paragraph.insertBefore(afterRun, next);
  } else {
    paragraph.appendChild(afterRun);
  }
}

function splitParagraphAtOffset(paragraph: Element, offset: number, doc: XMLDocument): void {
  if (offset <= 0) return;

  const segments = getDrawingRunSegments(paragraph);
  const total = segments.at(-1)?.end ?? 0;
  if (offset >= total) return;

  for (const segment of segments) {
    if (offset > segment.start && offset < segment.end) {
      splitDrawingRunAt(paragraph, segment.runIndex, offset - segment.start, doc);
      return;
    }
  }
}

/**
 * Canonical string for a run's properties, used to decide whether two runs can
 * be coalesced. A run with no `<a:rPr>` and a run with an empty `<a:rPr/>` are
 * intentionally distinct keys so the comparison only ever merges byte-identical
 * properties (false negatives are harmless; false positives would lose styling).
 */
function runPropertiesKey(run: Element): string {
  const rPr = getElementChildren(run)
    .find((element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
  return rPr ? new XMLSerializer().serializeToString(rPr) : '';
}

function findRunTextElement(run: Element): Element | null {
  return getElementChildren(run)
    .find((element) => element.localName === 't' && element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? getDescendants(run, 't').find((element) => element.namespaceURI === DRAWINGML_NAMESPACE)
    ?? null;
}

/**
 * Merge DOM-adjacent `<a:r>` runs that carry identical run properties into one
 * run. Range styling splits runs at every selection boundary, so repeated edits
 * fragment a paragraph into many runs whose `rPr` are identical (the 33-run
 * problem). Coalescing them keeps the run list minimal, which shrinks every
 * subsequent offset mapping and keeps the saved OOXML clean.
 *
 * Only consecutive element-child runs are merged, so a line break (`<a:br>`),
 * field (`<a:fld>`), or any other element between two runs blocks the merge and
 * preserves document structure. The merge target keeps its own `<a:t>` element
 * (and attributes); `xml:space="preserve"` is added when the combined text gains
 * leading/trailing whitespace so spaces are never dropped on save.
 */
function mergeAdjacentRuns(paragraph: Element): boolean {
  let merged = false;
  let previousRun: Element | null = null;
  let previousKey = '';

  for (const child of getElementChildren(paragraph)) {
    const isRun = child.localName === 'r' && child.namespaceURI === DRAWINGML_NAMESPACE;
    if (!isRun) {
      previousRun = null;
      previousKey = '';
      continue;
    }

    const key = runPropertiesKey(child);
    if (previousRun && key === previousKey) {
      const targetText = findRunTextElement(previousRun);
      if (targetText) {
        const combined = getDrawingRunText(previousRun) + getDrawingRunText(child);
        const otherText = findRunTextElement(child);
        const preserve = combined !== combined.trim()
          || targetText.getAttribute('xml:space') === 'preserve'
          || otherText?.getAttribute('xml:space') === 'preserve';
        targetText.textContent = combined;
        if (preserve) {
          targetText.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
        }
        paragraph.removeChild(child);
        merged = true;
        // Keep `previousRun`/`previousKey` so a third identical run also folds in.
        continue;
      }
    }

    previousRun = child;
    previousKey = key;
  }

  return merged;
}

export function applyRunStyleToParagraphRange(
  paragraph: Element,
  doc: XMLDocument,
  startOffset: number,
  endOffset: number,
  change: RunStyleChange
): boolean {
  let start = Math.max(0, startOffset);
  let end = Math.max(0, endOffset);
  if (start > end) {
    [start, end] = [end, start];
  }

  if (start === end) {
    const segments = getDrawingRunSegments(paragraph);
    const total = segments.at(-1)?.end ?? 0;
    const position = Math.min(start, total);
    const segment = segments.find((candidate) => position >= candidate.start && position <= candidate.end)
      ?? segments.at(-1);
    if (!segment || segment.text.length === 0) return false;
    const changed = applyRunPropertyChange(getRunProperties(segment.run, doc), doc, change);
    if (changed) mergeAdjacentRuns(paragraph);
    return changed;
  }

  splitParagraphAtOffset(paragraph, start, doc);
  splitParagraphAtOffset(paragraph, end, doc);

  let changed = false;
  for (const segment of getDrawingRunSegments(paragraph)) {
    if (segment.end <= start || segment.start >= end) continue;
    if (segment.text.length === 0) continue;
    changed = applyRunPropertyChange(getRunProperties(segment.run, doc), doc, change) || changed;
  }
  if (changed) {
    // Collapse the boundary splits (and any prior fragmentation) back down so
    // repeated range edits cannot grow the run count without bound.
    mergeAdjacentRuns(paragraph);
  }
  return changed;
}

export function isParagraphRangeStyled(
  paragraph: Element,
  startOffset: number,
  endOffset: number,
  flag: 'bold' | 'italic' | 'underline'
): boolean {
  let start = Math.max(0, startOffset);
  let end = Math.max(0, endOffset);
  if (start > end) {
    [start, end] = [end, start];
  }
  if (start === end) return false;

  let matched = false;
  for (const segment of getDrawingRunSegments(paragraph)) {
    if (segment.end <= start || segment.start >= end) continue;
    if (segment.text.length === 0) continue;
    matched = true;

    const runProperties = getElementChildren(segment.run)
      .find((element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
    if (!runProperties) return false;

    if (flag === 'bold') {
      const bold = runProperties.getAttribute('b');
      if (bold !== '1' && bold !== 'true') return false;
    } else if (flag === 'italic') {
      const italic = runProperties.getAttribute('i');
      if (italic !== '1' && italic !== 'true') return false;
    } else {
      const underline = runProperties.getAttribute('u');
      if (!underline || underline === 'none') return false;
    }
  }
  return matched;
}

/**
 * Locate every non-overlapping occurrence of `query` in `source` and report the
 * `[start, end)` character ranges. Honors the case-insensitivity default by
 * lower-casing both sides; like the rest of this module it assumes lower-casing
 * preserves length (true for the scripts these decks use).
 */
export function findTextMatches(
  source: string,
  query: string,
  matchCase: boolean
): Array<{ start: number; end: number }> {
  const matches: Array<{ start: number; end: number }> = [];
  if (!query) {
    return matches;
  }

  const haystack = matchCase ? source : source.toLocaleLowerCase();
  const needle = matchCase ? query : query.toLocaleLowerCase();
  let index = 0;

  while (index <= source.length) {
    const found = haystack.indexOf(needle, index);
    if (found === -1) {
      break;
    }
    matches.push({ start: found, end: found + needle.length });
    index = found + needle.length;
  }

  return matches;
}

/**
 * Replace every occurrence of `query` within a single paragraph, allowing a
 * match to span multiple DrawingML runs (a:t). Runs are concatenated in
 * document order (the same view the find feature searches over the shape), the
 * matches are located in that combined string, and then each affected run is
 * rewritten: the replacement is anchored in the run where the match starts
 * (preserving that run's formatting) and the matched characters are removed from
 * the runs the match covers. Runs untouched by any match keep their exact text.
 * Returns how many matches were replaced.
 */
export function replaceTextInParagraph(
  paragraph: Element,
  query: string,
  replacement: string,
  matchCase: boolean
): number {
  const segments = getDrawingRunSegments(paragraph);
  if (segments.length === 0) {
    return 0;
  }

  const combined = segments.map((segment) => segment.text).join('');
  if (!combined) {
    return 0;
  }

  const matches = findTextMatches(combined, query, matchCase);
  if (matches.length === 0) {
    return 0;
  }

  const segmentIndexAt = (position: number): number => {
    for (let index = 0; index < segments.length; index++) {
      const segment = segments[index];
      if (segment && segment.start <= position && position < segment.end) {
        return index;
      }
    }
    return segments.length - 1;
  };

  const newTexts: string[] = segments.map(() => '');
  const appendTo = (index: number, value: string): void => {
    newTexts[index] = (newTexts[index] ?? '') + value;
  };

  let matchIndex = 0;
  for (let position = 0; position < combined.length; ) {
    const nextMatch = matches[matchIndex];
    if (nextMatch && position === nextMatch.start) {
      appendTo(segmentIndexAt(position), replacement);
      position = nextMatch.end;
      matchIndex += 1;
      continue;
    }
    appendTo(segmentIndexAt(position), combined.charAt(position));
    position += 1;
  }

  segments.forEach((segment, index) => {
    const updated = newTexts[index] ?? '';
    if (updated !== segment.text) {
      setDrawingRunText(segment.run, updated);
    }
  });

  return matches.length;
}

export interface DrawingRunPosition {
  paragraphIndex: number;
  runIndex: number;
  run: Element;
}

// Canonical child order of CT_TextCharacterProperties (a:rPr). New children must be
// inserted at the right position or PowerPoint rejects the slide on save.
export const RUN_PROPERTY_CHILD_ORDER = [
  'ln',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'effectLst',
  'effectDag',
  'highlight',
  'uLnTx',
  'uLn',
  'uFillTx',
  'uFill',
  'latin',
  'ea',
  'cs',
  'sym',
  'hlinkClick',
  'hlinkMouseOver',
  'rtl',
  'extLst'
];

export function getDrawingRuns(paragraph: Element): Element[] {
  return getElementChildren(paragraph)
    .filter((element) => element.localName === 'r' && element.namespaceURI === DRAWINGML_NAMESPACE);
}

export function getShapeRunPositions(shape: Element): DrawingRunPosition[] {
  const positions: DrawingRunPosition[] = [];
  getDrawingParagraphs(shape).forEach((paragraph, paragraphIndex) => {
    getDrawingRuns(paragraph).forEach((run, runIndex) => {
      positions.push({ paragraphIndex, runIndex, run });
    });
  });
  return positions;
}

/**
 * Text bodies without an explicit AutoFit child are rendered as no-auto-fit by
 * pptx-svg, unlike PowerPoint's shrink-on-overflow editing behavior. Preserve
 * explicit `noAutofit`, `normAutofit`, and `spAutoFit` choices; otherwise make
 * the safe default explicit before text mutations can overflow the frame.
 */
export function ensureDefaultShrinkAutofit(shape: Element, doc: XMLDocument): boolean {
  const bodyProps = getDescendants(shape, 'bodyPr')
    .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);

  let changed = false;
  for (const bodyPr of bodyProps) {
    const hasExplicitAutofit = getElementChildren(bodyPr).some((element) => (
      element.namespaceURI === DRAWINGML_NAMESPACE
        && (element.localName === 'noAutofit'
          || element.localName === 'normAutofit'
          || element.localName === 'spAutoFit')
    ));
    if (hasExplicitAutofit) continue;

    const normAutofit = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:normAutofit');
    const insertionPoint = getElementChildren(bodyPr).find((element) => (
      element.namespaceURI === DRAWINGML_NAMESPACE
        && (element.localName === 'scene3d'
          || element.localName === 'flatTx'
          || element.localName === 'extLst')
    ));
    bodyPr.insertBefore(normAutofit, insertionPoint ?? null);
    changed = true;
  }
  return changed;
}

/**
 * Disable shrink-to-fit ("normAutofit") on every text body in the shape. The
 * SVG renderer recomputes normAutofit dynamically and ignores the stored
 * fontScale, so when a user explicitly sets a font size the only way to honor
 * it is to turn the shrinking autofit off. normAutofit is replaced with
 * noAutofit; spAutoFit (resize shape to text) already honors the size and is
 * left intact.
 */
export function disableShrinkAutofit(shape: Element, doc: XMLDocument): boolean {
  const bodyProps = getDescendants(shape, 'bodyPr')
    .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);

  let changed = false;
  for (const bodyPr of bodyProps) {
    const normAutofit = getElementChildren(bodyPr)
      .find((element) => element.localName === 'normAutofit' && element.namespaceURI === DRAWINGML_NAMESPACE);
    if (!normAutofit) continue;

    const noAutofit = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:noAutofit');
    bodyPr.replaceChild(noAutofit, normAutofit);
    changed = true;
  }
  return changed;
}

export function getRunProperties(run: Element, doc: XMLDocument): Element {
  const existing = getElementChildren(run)
    .find((element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
  if (existing) {
    return existing;
  }

  const rPr = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:rPr');
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}

function insertRunPropertyChild(rPr: Element, child: Element): void {
  const order = RUN_PROPERTY_CHILD_ORDER.indexOf(child.localName);
  const reference = getElementChildren(rPr).find((existing) => {
    const existingOrder = RUN_PROPERTY_CHILD_ORDER.indexOf(existing.localName);
    return existingOrder !== -1 && existingOrder > order;
  }) ?? null;
  rPr.insertBefore(child, reference);
}

function setRunHighlight(rPr: Element, doc: XMLDocument, highlight: string | null): boolean {
  const highlights = getElementChildren(rPr)
    .filter((element) => element.localName === 'highlight' && element.namespaceURI === DRAWINGML_NAMESPACE);

  if (highlight === null) {
    if (highlights.length === 0) return false;
    highlights.forEach((element) => rPr.removeChild(element));
    return true;
  }

  const normalizedHighlight = normalizeHexColor(highlight);
  if (
    highlights.length === 1
    && getElementChildren(highlights[0])
      .filter((element) => element.localName === 'srgbClr' && element.namespaceURI === DRAWINGML_NAMESPACE)
      .some((element) => element.getAttribute('val') === normalizedHighlight)
  ) {
    return false;
  }

  highlights.forEach((element) => rPr.removeChild(element));
  const highlightElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:highlight');
  const colorElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:srgbClr');
  colorElement.setAttribute('val', normalizedHighlight);
  highlightElement.appendChild(colorElement);
  insertRunPropertyChild(rPr, highlightElement);
  return true;
}

export function normalizeHexColor(hex: string): string {
  return hex.replace(/^#/, '').toUpperCase();
}

function setRunLatinFont(rPr: Element, doc: XMLDocument, fontFamily: string): boolean {
  const latinFonts = getElementChildren(rPr)
    .filter((element) => element.localName === 'latin' && element.namespaceURI === DRAWINGML_NAMESPACE);
  if (latinFonts.length === 1 && latinFonts[0]!.getAttribute('typeface') === fontFamily) {
    return false;
  }

  latinFonts.forEach((element) => rPr.removeChild(element));
  const latin = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:latin');
  latin.setAttribute('typeface', fontFamily);
  insertRunPropertyChild(rPr, latin);
  return true;
}

function setRunSolidFill(rPr: Element, doc: XMLDocument, color: string | null): boolean {
  const fills = getElementChildren(rPr)
    .filter((element) => element.localName === 'solidFill' && element.namespaceURI === DRAWINGML_NAMESPACE);

  if (color === null) {
    if (fills.length === 0) return false;
    fills.forEach((element) => rPr.removeChild(element));
    return true;
  }

  const normalizedColor = normalizeHexColor(color);
  if (
    fills.length === 1
    && getElementChildren(fills[0])
      .filter((element) => element.localName === 'srgbClr' && element.namespaceURI === DRAWINGML_NAMESPACE)
      .some((element) => element.getAttribute('val') === normalizedColor)
  ) {
    return false;
  }

  fills.forEach((element) => rPr.removeChild(element));
  const solidFill = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:solidFill');
  const colorElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:srgbClr');
  colorElement.setAttribute('val', normalizedColor);
  solidFill.appendChild(colorElement);
  insertRunPropertyChild(rPr, solidFill);
  return true;
}

// Apply every requested run-level property directly to an <a:rPr>. The WASM
// renderer does not preserve <a:highlight> when it re-serializes a slide, so
// every run-style edit (not just highlight) is performed via OOXML to keep the
// highlight from being clobbered by a later renderer mutation on the same run.
export function applyRunPropertyChange(rPr: Element, doc: XMLDocument, change: RunStyleChange): boolean {
  let changed = false;
  if (change.bold !== undefined) {
    const value = change.bold ? '1' : '0';
    if (rPr.getAttribute('b') !== value) {
      rPr.setAttribute('b', value);
      changed = true;
    }
  }
  if (change.italic !== undefined) {
    const value = change.italic ? '1' : '0';
    if (rPr.getAttribute('i') !== value) {
      rPr.setAttribute('i', value);
      changed = true;
    }
  }
  if (change.underline !== undefined) {
    const value = change.underline ? 'sng' : 'none';
    if (rPr.getAttribute('u') !== value) {
      rPr.setAttribute('u', value);
      changed = true;
    }
  }
  if (change.fontSizePt !== undefined) {
    const value = String(Math.round(change.fontSizePt * 100));
    if (rPr.getAttribute('sz') !== value) {
      rPr.setAttribute('sz', value);
      changed = true;
    }
  }
  if (change.fontFamily !== undefined && change.fontFamily !== '') {
    changed = setRunLatinFont(rPr, doc, change.fontFamily) || changed;
  }
  if (change.color !== undefined) {
    changed = setRunSolidFill(rPr, doc, change.color) || changed;
  }
  if (change.highlight !== undefined) {
    changed = setRunHighlight(rPr, doc, change.highlight) || changed;
  }
  return changed;
}

export function getParagraphProperties(paragraph: Element, doc: XMLDocument): Element {
  const existing = getElementChildren(paragraph)
    .find((element) => element.localName === 'pPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
  if (existing) {
    return existing;
  }

  const pPr = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:pPr');
  paragraph.insertBefore(pPr, paragraph.firstChild);
  return pPr;
}

export function resolvePptxRunAlignment(value: string | null): 'ctr' | 'just' | 'l' | 'r' | null {
  if (value === 'l' || value === 'ctr' || value === 'r' || value === 'just') {
    return value;
  }
  return null;
}
