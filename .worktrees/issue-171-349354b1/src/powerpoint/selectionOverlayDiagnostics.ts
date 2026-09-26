export interface SelectionOverlayRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface SelectionOverlayContributorCandidate {
  tagName: string;
  identity: string;
  clientRect: SelectionOverlayRect;
  localRect: SelectionOverlayRect;
  transform: string | null;
  visibility: string | null;
  display: string | null;
  uiKind: 'caret' | 'selection' | 'editor' | 'overlay' | 'other';
}

export interface SelectionOverlayBoundsAnomaly {
  leftDelta: number;
  topDelta: number;
  widthDelta: number;
  heightDelta: number;
  widthRatio: number;
  heightRatio: number;
  materiallyDifferent: boolean;
}

export interface SelectionOverlayContributorSummary {
  descendantCount: number;
  outlierCount: number;
  contributors: Array<SelectionOverlayContributorCandidate & {
    extendsBeyondFrame: boolean;
    impact: number;
  }>;
}

const DEFAULT_ANOMALY_DELTA_PX = 8;
const DEFAULT_ANOMALY_RATIO = 1.25;
const DEFAULT_MAX_CONTRIBUTORS = 6;

function finite(value: number): number {
  return Number.isFinite(value) ? value : 0;
}

function safeRatio(first: number, second: number): number {
  const numerator = Math.max(Math.abs(first), 0.001);
  const denominator = Math.max(Math.abs(second), 0.001);
  return Math.max(numerator / denominator, denominator / numerator);
}

function right(rect: SelectionOverlayRect): number {
  return rect.left + rect.width;
}

function bottom(rect: SelectionOverlayRect): number {
  return rect.top + rect.height;
}

function outsideDistance(rect: SelectionOverlayRect, frame: SelectionOverlayRect): number {
  return Math.max(
    frame.left - rect.left,
    frame.top - rect.top,
    right(rect) - right(frame),
    bottom(rect) - bottom(frame),
    0,
  );
}

function area(rect: SelectionOverlayRect): number {
  return Math.max(0, rect.width) * Math.max(0, rect.height);
}

export function getSelectionOverlayBoundsAnomaly(
  groupBox: SelectionOverlayRect | null,
  frameBox: SelectionOverlayRect | null,
  deltaThreshold = DEFAULT_ANOMALY_DELTA_PX,
  ratioThreshold = DEFAULT_ANOMALY_RATIO,
): SelectionOverlayBoundsAnomaly | null {
  if (!groupBox || !frameBox) {
    return null;
  }

  const result = {
    leftDelta: finite(groupBox.left - frameBox.left),
    topDelta: finite(groupBox.top - frameBox.top),
    widthDelta: finite(groupBox.width - frameBox.width),
    heightDelta: finite(groupBox.height - frameBox.height),
    widthRatio: safeRatio(groupBox.width, frameBox.width),
    heightRatio: safeRatio(groupBox.height, frameBox.height),
    materiallyDifferent: false,
  };
  result.materiallyDifferent = Math.max(
    Math.abs(result.leftDelta),
    Math.abs(result.topDelta),
    Math.abs(result.widthDelta),
    Math.abs(result.heightDelta),
  ) >= deltaThreshold
    && (result.widthRatio >= ratioThreshold || result.heightRatio >= ratioThreshold);
  return result;
}

export function summarizeSelectionOverlayContributors(
  candidates: readonly SelectionOverlayContributorCandidate[],
  frameBox: SelectionOverlayRect,
  maxContributors = DEFAULT_MAX_CONTRIBUTORS,
): SelectionOverlayContributorSummary {
  const annotated = candidates.map((candidate) => {
    const distance = outsideDistance(candidate.localRect, frameBox);
    const candidateArea = area(candidate.localRect);
    const frameArea = Math.max(area(frameBox), 1);
    return {
      ...candidate,
      extendsBeyondFrame: distance > 0,
      impact: Math.round((distance * Math.max(candidateArea / frameArea, 0.1)) * 100) / 100,
    };
  });
  const outlierCount = annotated.filter((candidate) => candidate.extendsBeyondFrame).length;
  const outliers = annotated
    .filter((candidate) => candidate.extendsBeyondFrame)
    .sort((first, second) => second.impact - first.impact)
    .slice(0, Math.max(0, maxContributors));

  return {
    descendantCount: candidates.length,
    outlierCount,
    contributors: outliers,
  };
}
