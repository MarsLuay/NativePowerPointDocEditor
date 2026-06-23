import type { RunStyleChange } from '../PresentationEngine';
import {
  DRAWINGML_NAMESPACE,
  getDescendants,
  getElementChildren,
} from './ooxmlXml';

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

export function setDrawingText(container: Element, text: string): void {
  const textElements = getDescendants(container, 't')
    .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
  const firstText = textElements[0];
  if (!firstText) {
    throw new Error('This PowerPoint label has no editable text node.');
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
    const runs = getDrawingRuns(paragraph);
    if (runs.length === 0) {
      throw new Error('Could not find the selected text paragraph runs.');
    }

    runs.forEach((run, runIndex) => {
      setDrawingRunText(run, runIndex === 0 ? text : '');
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
    appendDrawingParagraphRun(paragraph, templateRun, line);
  });
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
    applyRunPropertyChange(getRunProperties(segment.run, doc), doc, change);
    mergeAdjacentRuns(paragraph);
    return true;
  }

  splitParagraphAtOffset(paragraph, start, doc);
  splitParagraphAtOffset(paragraph, end, doc);

  let changed = false;
  for (const segment of getDrawingRunSegments(paragraph)) {
    if (segment.end <= start || segment.start >= end) continue;
    if (segment.text.length === 0) continue;
    applyRunPropertyChange(getRunProperties(segment.run, doc), doc, change);
    changed = true;
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

function setRunHighlight(rPr: Element, doc: XMLDocument, highlight: string | null): void {
  getElementChildren(rPr)
    .filter((element) => element.localName === 'highlight' && element.namespaceURI === DRAWINGML_NAMESPACE)
    .forEach((element) => rPr.removeChild(element));

  if (highlight === null) {
    return;
  }

  const highlightElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:highlight');
  const colorElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:srgbClr');
  colorElement.setAttribute('val', normalizeHexColor(highlight));
  highlightElement.appendChild(colorElement);
  insertRunPropertyChild(rPr, highlightElement);
}

export function normalizeHexColor(hex: string): string {
  return hex.replace(/^#/, '').toUpperCase();
}

function setRunLatinFont(rPr: Element, doc: XMLDocument, fontFamily: string): void {
  getElementChildren(rPr)
    .filter((element) => element.localName === 'latin' && element.namespaceURI === DRAWINGML_NAMESPACE)
    .forEach((element) => rPr.removeChild(element));

  const latin = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:latin');
  latin.setAttribute('typeface', fontFamily);
  insertRunPropertyChild(rPr, latin);
}

function setRunSolidFill(rPr: Element, doc: XMLDocument, color: string | null): void {
  getElementChildren(rPr)
    .filter((element) => element.localName === 'solidFill' && element.namespaceURI === DRAWINGML_NAMESPACE)
    .forEach((element) => rPr.removeChild(element));

  if (color === null) {
    return;
  }

  const solidFill = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:solidFill');
  const colorElement = doc.createElementNS(DRAWINGML_NAMESPACE, 'a:srgbClr');
  colorElement.setAttribute('val', normalizeHexColor(color));
  solidFill.appendChild(colorElement);
  insertRunPropertyChild(rPr, solidFill);
}

// Apply every requested run-level property directly to an <a:rPr>. The WASM
// renderer does not preserve <a:highlight> when it re-serializes a slide, so
// every run-style edit (not just highlight) is performed via OOXML to keep the
// highlight from being clobbered by a later renderer mutation on the same run.
export function applyRunPropertyChange(rPr: Element, doc: XMLDocument, change: RunStyleChange): void {
  if (change.bold !== undefined) {
    rPr.setAttribute('b', change.bold ? '1' : '0');
  }
  if (change.italic !== undefined) {
    rPr.setAttribute('i', change.italic ? '1' : '0');
  }
  if (change.underline !== undefined) {
    rPr.setAttribute('u', change.underline ? 'sng' : 'none');
  }
  if (change.fontSizePt !== undefined) {
    rPr.setAttribute('sz', String(Math.round(change.fontSizePt * 100)));
  }
  if (change.fontFamily !== undefined && change.fontFamily !== '') {
    setRunLatinFont(rPr, doc, change.fontFamily);
  }
  if (change.color !== undefined) {
    setRunSolidFill(rPr, doc, change.color);
  }
  if (change.highlight !== undefined) {
    setRunHighlight(rPr, doc, change.highlight);
  }
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
