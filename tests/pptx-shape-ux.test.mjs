import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { test } from "node:test";
import { loadNativePowerPointViewModule } from "./helpers/load-plugin-modules.mjs";

function createView(NativePowerPointView) {
  return new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
}

test("resize previews resize text frames while compensating text glyph scale", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const appliedTransforms = [];
  const previewElement = {
    setAttribute: (name, value) => appliedTransforms.push([name, value]),
    querySelector: () => null,
  };
  const textPreviewElement = {
    querySelector: () => ({}),
  };
  const startTransform = { x: 10, y: 20, cx: 100, cy: 50, rot: 0 };

  assert.equal(view.shouldFreezeTextDuringResize("resize", textPreviewElement), true);
  assert.equal(view.shouldFreezeTextDuringResize("move", textPreviewElement), false);

  view.engine = { getSlideScale: () => 1 };
  view.svgEl = {};
  view.dragState = {
    mode: "resize",
    handle: "w",
    startTransform,
    previewElement,
    previewOriginalTransform: null,
  };

  // Non-text shapes retain a live SVG preview. Pulling the west handle from
  // x=10 to x=60 keeps the east edge fixed at x=110.
  view.updateShapeTransformPreview({ x: 60, y: 20, cx: 50, cy: 50, rot: 0 });
  assert.deepEqual(appliedTransforms, [
    ["transform", "translate(110 45) scale(0.5 1) translate(-110 -45)"],
  ]);

  let outlineUpdates = 0;
  appliedTransforms.length = 0;
  view.engine = { getSlideScale: () => 1, pxToEmu: (value) => value };
  view.dragState = {
    mode: "resize",
    handle: "se",
    pointerId: 7,
    startPoint: { x: 0, y: 0 },
    startClientX: 0,
    startClientY: 0,
    startBox: { left: 0, top: 0, width: 100, height: 50 },
    startTransform,
    latestTransform: startTransform,
    previewElement,
    previewOriginalTransform: null,
    freezeShapeDuringResize: true,
  };
  view.getSvgPoint = () => ({ x: 25, y: 15 });
  view.getSelectedShapeElement = () => previewElement;
  view.isPictureShape = () => false;
  view.shapeHasRotation = () => false;
  view.snapController = {
    computeSnap: () => ({ dx: 0, dy: 0, guideX: null, guideY: null }),
    updateSnapGuides: () => {},
  };
  view.applySelectionOverlayBox = () => { outlineUpdates += 1; };

  view.handleDragMove({ pointerId: 7, clientX: 25, clientY: 15 });
  assert.equal(outlineUpdates, 1, "outline and resize handles must continue to track the drag");
  assert.equal(appliedTransforms.length, 1, "the text frame must track the live resize outline");
  assert.match(
    view.getTextResizeCompensationTransform(1.25, 1.3, 10, 20, 25, 15) ?? "",
    /^matrix\(0\.8 0 0 0\.769/,
    "the nested text transform cancels the frame scale instead of stretching glyphs",
  );
});

test("resize snapping locks only the dragged edges and preserves the opposite anchor", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  let probe;
  view.engine = { pxToEmu: (value) => value };
  view.snapController = {
    computeSnap: (box) => {
      probe = box;
      return { dx: 5, dy: -3, guideX: 15, guideY: 17 };
    },
  };

  const snapped = view.getResizeSnap(
    { x: 10, y: 20, cx: 100, cy: 50, rot: 0 },
    "nw",
    new Set([7]),
  );
  assert.deepEqual(probe, { x: 10, y: 20, cx: 0, cy: 0 });
  assert.deepEqual(snapped, {
    transform: { x: 15, y: 17, cx: 95, cy: 53, rot: 0 },
    guideX: 15,
    guideY: 17,
  });
  assert.equal(
    snapped.transform.x + snapped.transform.cx,
    110,
    "west-edge snap must keep the east anchor fixed",
  );
  assert.equal(
    snapped.transform.y + snapped.transform.cy,
    70,
    "north-edge snap must keep the south anchor fixed",
  );
});

test("resize snapping does not shrink a frame below its minimum size", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  view.engine = { pxToEmu: (value) => value };
  view.snapController = {
    computeSnap: () => ({ dx: -4, dy: 0, guideX: 8, guideY: null }),
  };

  const snapped = view.getResizeSnap(
    { x: 0, y: 0, cx: 12, cy: 40, rot: 0 },
    "e",
    new Set(),
  );
  assert.deepEqual(snapped, {
    transform: { x: 0, y: 0, cx: 12, cy: 40, rot: 0 },
    guideX: null,
    guideY: null,
  });
});

test("multi-selection transforms scale and rotate every selected shape around one group frame", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const first = { x: 0, y: 0, cx: 50, cy: 50, rot: 0 };
  const second = { x: 150, y: 0, cx: 50, cy: 50, rot: 0 };
  const start = new Map([[1, first], [2, second]]);
  const startBounds = view.getGroupTransformBounds(start.values());
  assert.deepEqual(startBounds, { x: 0, y: 0, cx: 200, cy: 50, rot: 0 });

  view.engine = {
    pxToEmu: (value) => value,
    ooxmlToDegrees: (value) => value,
    degreesToOoxml: (value) => value,
    getSlideScale: () => 1,
  };
  const resizedBounds = view.getGroupResizeBounds({
    handle: "w",
    start,
    startBounds,
  }, 100, 0);
  assert.deepEqual(resizedBounds, { x: 100, y: 0, cx: 100, cy: 50, rot: 0 });
  const resized = view.scaleGroupTransforms(start, startBounds, resizedBounds, 0.5, 1);
  assert.deepEqual([...resized.entries()], [
    [1, { x: 100, y: 0, cx: 25, cy: 50, rot: 0 }],
    [2, { x: 175, y: 0, cx: 25, cy: 50, rot: 0 }],
  ]);

  const crossingDrag = { handle: "w", start, startBounds };
  const crossedBounds = view.getGroupResizeBounds(crossingDrag, 250, 0);
  assert.deepEqual(crossedBounds, { x: 200, y: 0, cx: 50, cy: 50, rot: 0 });
  assert.equal(crossingDrag.crossedHorizontal, true);
  assert.deepEqual(
    [...view.scaleGroupTransforms(start, startBounds, crossedBounds, 0.25, 1, true, false).entries()],
    [
      [1, { x: 237.5, y: 0, cx: 12.5, cy: 50, rot: 0 }],
      [2, { x: 200, y: 0, cx: 12.5, cy: 50, rot: 0 }],
    ],
    "crossing west past east mirrors the selected objects across the fixed edge",
  );
  view.svgEl = {};
  const picturePreviewTransform = view.getGroupResizePreviewTransform(crossingDrag, crossedBounds);
  assert.match(
    picturePreviewTransform ?? "",
    /scale\(-0\.25 1\)/,
    "pictures use the same inverted affine transform as the crossed group outline",
  );

  view.groupDrag = {
    mode: "rotate",
    pointerId: 1,
    startPoint: { x: 100, y: 25 },
    startClientX: 100,
    startClientY: 25,
    startBox: { left: 0, top: 0, width: 200, height: 50 },
    startBounds,
    latestBounds: startBounds,
    centerClientX: 100,
    centerClientY: 25,
    startAngle: 0,
    start,
    latest: new Map(start),
    moved: true,
  };
  view.updateGroupRotateDrag({ clientX: 100, clientY: 125, shiftKey: false });
  assert.deepEqual([...view.groupDrag.latest.entries()], [
    [1, { x: 75, y: -75, cx: 50, cy: 50, rot: 90 }],
    [2, { x: 75, y: 75, cx: 50, cy: 50, rot: 90 }],
  ]);
  assert.equal(view.groupDrag.rotationSnapTarget, 90);

  const rotateAt = (degrees) => ({
    clientX: 100 + Math.cos((degrees * Math.PI) / 180) * 100,
    clientY: 25 + Math.sin((degrees * Math.PI) / 180) * 100,
    shiftKey: false,
  });
  view.updateGroupRotateDrag(rotateAt(170));
  view.updateGroupRotateDrag(rotateAt(-170));
  assert.equal(
    Math.round(view.groupDrag.accumulatedRotationDegrees),
    190,
    "crossing the atan2 seam must continue a clockwise turn instead of reversing it",
  );
});

test("group resize preview transforms every object and gives text a reflow path", async () => {
  const source = await readFile(resolve("src/powerpoint/ui/NativePowerPointView.ts"), "utf8");
  const start = source.indexOf("private applyGroupResizePreview(");
  const end = source.indexOf("private restoreGroupShapePreviews", start);

  assert.ok(start >= 0 && end > start);
  const preview = source.slice(start, end);
  assert.doesNotMatch(preview, /this\.isPictureShape\(shape\)/,
    "non-picture text boxes and shapes must move with a group resize preview");
  assert.match(preview, /shape\.setAttribute\('transform'/);
  assert.match(preview, /this\.applyTextResizePreview\(/);

  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  assert.match(
    view.getTextResizeCompensationTransform(-0.25, 1, 200, 25, 237.5, 0) ?? "",
    /^matrix\(-4 0 0 1 /,
    "crossed group resize must cancel its mirror for readable text glyphs",
  );
});

test("group resize keeps later text frames moving when one preview reflow fails", async () => {
  const source = await readFile(resolve("src/powerpoint/ui/NativePowerPointView.ts"), "utf8");
  const groupPreviewStart = source.indexOf("private applyGroupResizePreview(");
  const groupPreviewEnd = source.indexOf("private restoreGroupShapePreviews", groupPreviewStart);
  const textPreviewStart = source.indexOf("private applyTextResizePreview(");
  const textPreviewEnd = source.indexOf("private reflowShapeTextResizePreview", textPreviewStart);

  assert.match(source.slice(groupPreviewStart, groupPreviewEnd), /textTransformCount \+= 1/);
  assert.match(source.slice(groupPreviewStart, groupPreviewEnd), /previewTextReflowError/);
  assert.match(source.slice(textPreviewStart, textPreviewEnd), /onReflowError\?\./);
});

test("text resize previews preserve glyph size while moving each text anchor relatively", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);

  assert.deepEqual(
    view.getRelativeTextPreviewTranslation({ x: 110, y: 60 }, 2, 0.5, 0, 0),
    { x: 110, y: -30 },
    "the text anchor must follow its original relative point as the group expands/squeezes",
  );
  assert.deepEqual(
    view.getRelativeTextPreviewTranslation({ x: 110, y: 60 }, -0.25, 1, 200, 25),
    { x: 112.5, y: 0 },
    "an inverted resize moves readable text to the mirrored relative position",
  );
});

test("resize handles can cross their opposite edge without producing negative OOXML extents", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  view.engine = { pxToEmu: (value) => value };
  view.dragState = {
    mode: "resize",
    handle: "w",
    pointerId: 4,
    startPoint: { x: 0, y: 0 },
    startClientX: 0,
    startClientY: 0,
    startBox: { left: 10, top: 20, width: 100, height: 50 },
    startTransform: { x: 10, y: 20, cx: 100, cy: 50, rot: 0 },
    latestTransform: { x: 10, y: 20, cx: 100, cy: 50, rot: 0 },
  };

  const crossed = view.getResizeTransform(150, 0);
  assert.deepEqual(crossed, {
    transform: { x: 110, y: 20, cx: 50, cy: 50, rot: 0 },
    crossedHorizontal: true,
    crossedVertical: false,
  });
  assert.deepEqual(
    view.computeDragOverlayBox({ clientX: 150, clientY: 0 }),
    { left: 110, top: 20, width: 50, height: 50 },
    "the outline follows the handle beyond the former opposite edge",
  );
});

test("crossed resize styles move the active grip to the opposite outline edge", async () => {
  const css = await readFile(resolve("styles.css"), "utf8");
  assert.match(css, /native-powerpoint-resize-crossed-horizontal[\s\S]*?native-powerpoint-resize-w/);
  assert.match(css, /native-powerpoint-resize-crossed-vertical[\s\S]*?native-powerpoint-resize-n/);
});

test("multi-selection uses one interactive outline with the single-selection controls", async () => {
  const source = await readFile(resolve("src/powerpoint/ui/NativePowerPointView.ts"), "utf8");
  const overlayStart = source.indexOf("private updateSelectionOverlay(): void");
  const overlayEnd = source.indexOf("private startCurrentSelectionDrag", overlayStart);
  const multiBoxStart = source.indexOf("private updateMultiSelectionBoxes(): void");
  const multiBoxEnd = source.indexOf("private removeMultiSelectionBoxes", multiBoxStart);

  assert.ok(overlayStart >= 0 && overlayEnd > overlayStart);
  assert.match(source.slice(overlayStart, overlayEnd), /const isMultiSelection = this\.selectedShapeIndices\.size > 1;/);
  assert.match(source.slice(overlayStart, overlayEnd), /this\.applyMultiSelectionOverlayLayout\(\)/);
  assert.match(source.slice(overlayStart, overlayEnd), /native-powerpoint-move-border-seg-/);
  assert.match(source.slice(overlayStart, overlayEnd), /for \(const segment of \['a', 'b'\]/);
  assert.match(source.slice(overlayStart, overlayEnd), /native-powerpoint-selection-move-surface/);
  assert.match(source, /this\.startGroupDrag\(event, mode, handle\);/);
  assert.ok(multiBoxStart >= 0 && multiBoxEnd > multiBoxStart);
  assert.match(
    source.slice(multiBoxStart, multiBoxEnd),
    /native-powerpoint-multi-selection-box/,
    "each selected object must get its own dashed outline while multi-selecting",
  );
});

test("multi-select union outline passes through additive clicks until modifier release", async () => {
  const source = await readFile(resolve("src/powerpoint/ui/NativePowerPointView.ts"), "utf8");
  const css = await readFile(resolve("styles.css"), "utf8");

  assert.match(source, /syncSelectionChromeHitThrough/);
  assert.match(source, /handleAdditiveSelectThroughSelectionChrome/);
  assert.match(source, /native-powerpoint-additive-hit-through/);
  assert.match(source, /isAdditiveSelectionModifier\(event\)/);
  assert.match(
    css,
    /\.native-powerpoint-selection-box\.native-powerpoint-additive-hit-through\.native-powerpoint-multi-selection-outline \.native-powerpoint-selection-move-surface/,
  );
  assert.match(
    css,
    /\.native-powerpoint-selection-box\.native-powerpoint-additive-hit-through \.native-powerpoint-move-border/,
  );
});

test("outline edge strips leave a center gap so mid-edge resize dots stay reachable", async () => {
  const css = await readFile(resolve("styles.css"), "utf8");
  assert.match(css, /native-powerpoint-move-border-seg-a/);
  assert.match(css, /native-powerpoint-move-border-seg-b/);
  assert.match(css, /calc\(50% \+ 16px\)/);
  assert.match(css, /\.native-powerpoint-resize-handle::before/);
  assert.match(css, /\.native-powerpoint-resize-handle\s*\{[^}]*z-index:\s*8/s);
  assert.match(css, /native-powerpoint-selection-move-surface/);
  assert.match(
    css,
    /\.native-powerpoint-selection-box\.native-powerpoint-multi-selection-outline \.native-powerpoint-selection-move-surface\s*\{[^}]*pointer-events:\s*auto/s,
  );
  assert.match(css, /\.native-powerpoint-move-border-n,\s*\.native-powerpoint-move-border-s\s*\{[^}]*height:\s*6px/s);
  assert.match(css, /\.native-powerpoint-move-border\s*\{[^}]*cursor:\s*move/s);
});

test("multi-selection edge strips move the group instead of stretching it", async () => {
  const source = await readFile(resolve("src/powerpoint/ui/NativePowerPointView.ts"), "utf8");
  const overlayStart = source.indexOf("private updateSelectionOverlay(): void");
  const overlayEnd = source.indexOf("private startCurrentSelectionDrag", overlayStart);
  const overlay = source.slice(overlayStart, overlayEnd);

  assert.match(overlay, /native-powerpoint-move-border/);
  assert.match(overlay, /source: 'outline-edge'/);
  assert.match(overlay, /startCurrentSelectionDrag\(event, 'move'\)/);
  assert.doesNotMatch(
    overlay,
    /move-border[\s\S]*?startCurrentSelectionDrag\(event, 'resize', side\)/,
    "outline edges must move; only resize-handle dots stretch",
  );
});

test("south group resize grows height from the fixed north edge", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  view.engine = { pxToEmu: (value) => value };
  const top = { x: 0, y: 0, cx: 50, cy: 40, rot: 0 };
  const bottom = { x: 0, y: 100, cx: 50, cy: 40, rot: 0 };
  const start = new Map([[1, top], [2, bottom]]);
  const startBounds = view.getGroupTransformBounds(start.values());
  assert.deepEqual(startBounds, { x: 0, y: 0, cx: 50, cy: 140, rot: 0 });

  const resizedBounds = view.getGroupResizeBounds({
    handle: "s",
    start,
    startBounds,
  }, 0, 70);
  assert.deepEqual(resizedBounds, { x: 0, y: 0, cx: 50, cy: 210, rot: 0 });
  const scaleY = resizedBounds.cy / startBounds.cy;
  assert.deepEqual(
    [...view.scaleGroupTransforms(start, startBounds, resizedBounds, 1, scaleY).entries()],
    [
      [1, { x: 0, y: 0, cx: 50, cy: 60, rot: 0 }],
      [2, { x: 0, y: 150, cx: 50, cy: 60, rot: 0 }],
    ],
    "bottom row must move down and grow when the south handle stretches the group",
  );
});

test("pure vertical group resize skips text inverse compensation", async () => {
  const source = await readFile(resolve("src/powerpoint/ui/NativePowerPointView.ts"), "utf8");
  const start = source.indexOf("private applyGroupResizePreview(");
  const end = source.indexOf("private restoreGroupShapePreviews", start);
  assert.ok(start >= 0 && end > start);
  const body = source.slice(start, end);
  assert.match(body, /Pure vertical group stretch/);
  assert.match(body, /Math\.abs\(Math\.abs\(scaleX\) - 1\) <= 0\.001/);
  assert.match(body, /Math\.abs\(Math\.abs\(scaleY\) - 1\) > 0\.001/);
});

test("text preview anchors normalize nested body-local coordinates into slide space", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const shape = { parentElement: null };
  const inner = {
    getAttribute: (name) => (name === "transform" ? "translate(0 100)" : null),
    parentElement: shape,
  };
  const text = {
    parentElement: inner,
    getAttribute: () => null,
    querySelector: () => ({
      getAttribute: (name) => (name === "y" ? "25" : name === "x" ? "0" : null),
    }),
  };

  const anchor = view.resolveTextPreviewAnchorInSlideSpace(
    shape,
    text,
    { x: 0, y: 100, cx: 50, cy: 40, rot: 0 },
    1,
  );
  assert.deepEqual(
    anchor,
    { x: 0, y: 125 },
    "body-local tspan y under nested translate must become slide-space for group resize",
  );

  assert.deepEqual(
    view.getRelativeTextPreviewTranslation(anchor, 1, 1.5, 0, 0),
    { x: 0, y: 62.5 },
    "lower-frame text translation must use slide-space distance from group north",
  );
});

test("multi-select SVG pointer near a resize dot redirects to stretch, not move", async () => {
  const source = await readFile(resolve("src/powerpoint/ui/NativePowerPointView.ts"), "utf8");
  assert.match(source, /findSelectionResizeHandleAtPoint\(/);
  assert.match(source, /PowerPoint shape pointer redirected to resize chrome/);
  assert.match(source, /PowerPoint text pointer redirected to resize chrome/);
  assert.match(source, /this\.startCurrentSelectionDrag\(event, 'resize', resizeHandle\)/);
  assert.match(source, /source: 'move-surface-grip'/);
  assert.match(
    await readFile(resolve("styles.css"), "utf8"),
    /\.native-powerpoint-selection-box\.native-powerpoint-multi-selection-outline \.native-powerpoint-selection-move-surface\s*\{[^}]*bottom:\s*12px/s,
  );
});

test("marquee selection previews a passive union outline before pointer-up", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const created = [];
  const previewedIndices = [];
  const previewClasses = [];

  view.canvasPane = {
    scrollLeft: 0,
    scrollTop: 0,
    getBoundingClientRect: () => ({ left: 0, top: 0 }),
    createDiv: ({ cls }) => {
      const element = {
        cls,
        setCssProps: (props) => created.push([cls, props]),
        remove() {},
      };
      return element;
    },
  };
  view.marquee = {
    pointerId: 3,
    startClientX: 10,
    startClientY: 20,
    additive: true,
    base: [1],
    moved: false,
  };
  view.collectShapesInClientRect = () => [2, 5];
  view.previewSelectionClasses = (indices) => previewClasses.push([...indices]);
  view.getSelectionBoxForShapeIndices = (indices) => {
    previewedIndices.push([...indices]);
    return { left: 11, top: 22, width: 33, height: 44 };
  };

  view.updateMarquee({ pointerId: 3, clientX: 50, clientY: 60 });

  assert.deepEqual(previewClasses, [[1, 2, 5]]);
  assert.deepEqual(previewedIndices, [[1, 2, 5]]);
  assert.deepEqual(created, [
    ["native-powerpoint-marquee-box", { left: "10px", top: "20px", width: "40px", height: "40px" }],
    ["native-powerpoint-marquee-selection-preview", { left: "11px", top: "22px", width: "33px", height: "44px" }],
  ]);
});

test("rotate previews turn the rendered shape group with its selection outline", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const appliedTransforms = [];
  const overlayTransforms = [];
  const previewElement = {
    setAttribute: (name, value) => appliedTransforms.push([name, value]),
  };
  const startTransform = { x: 10, y: 20, cx: 100, cy: 50, rot: 0 };

  view.engine = {
    getSlideScale: () => 1,
    ooxmlToDegrees: (value) => value,
    degreesToOoxml: (value) => value,
  };
  view.svgEl = {};
  view.selectionOverlay = {
    setCssProps: (props) => overlayTransforms.push(props),
  };
  view.dragState = {
    mode: "rotate",
    startTransform,
    latestTransform: startTransform,
    previewElement,
    previewOriginalTransform: null,
    centerClientX: 60,
    centerClientY: 45,
    startAngle: 0,
  };

  const angle = (88 * Math.PI) / 180;
  view.updateRotateDrag({
    clientX: 60 + Math.cos(angle) * 100,
    clientY: 45 + Math.sin(angle) * 100,
    shiftKey: false,
  });

  assert.deepEqual(appliedTransforms, [
    ["transform", "rotate(90 60 45)"],
  ]);
  assert.deepEqual(overlayTransforms, [{ transform: "rotate(90deg)" }]);
  assert.equal(view.dragState.rotationSnapTarget, 90);
  assert.equal(view.getCardinalRotationSnap(84), null, "free rotation must continue outside the snap range");

  const rotateAt = (degrees) => ({
    clientX: 60 + Math.cos((degrees * Math.PI) / 180) * 100,
    clientY: 45 + Math.sin((degrees * Math.PI) / 180) * 100,
    shiftKey: false,
  });
  view.updateRotateDrag(rotateAt(170));
  view.updateRotateDrag(rotateAt(-170));
  assert.equal(
    Math.round(view.dragState.accumulatedRotationDegrees),
    190,
    "the single-object outline must also remain continuous through a full turn",
  );
  assert.deepEqual(overlayTransforms.at(-1), { transform: "rotate(190deg)" });
});

test("selected PowerPoint shapes use an outline without a glow", async () => {
  const css = await readFile(resolve("styles.css"), "utf8");
  const selectedShapeRule = css.match(/\.native-powerpoint-slide-svg \.native-powerpoint-shape-selected\s*\{([\s\S]*?)\}/);

  assert.ok(selectedShapeRule, "the selected-shape rule must remain explicit for the selection state");
  assert.match(selectedShapeRule[1], /filter:\s*none/);
  assert.doesNotMatch(selectedShapeRule[1], /drop-shadow/);
});

test("inline text edits restore a canvas position changed by native caret reveal", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const restored = [];
  view.canvasPane = { scrollLeft: 12, scrollTop: 34 };
  view.restoreCanvasScrollSoon = (position) => restored.push(position);

  view.canvasPane.scrollLeft = 2;
  view.canvasPane.scrollTop = 3;
  view.preserveCanvasScrollAfterInlineTextEdit({ left: 12, top: 34 }, "deleteContentBackward");

  assert.deepEqual(restored, [{ left: 12, top: 34 }]);
});

test("PowerPoint canvas disables browser scroll anchoring during text preview updates", async () => {
  const css = await readFile(resolve("styles.css"), "utf8");
  const canvasPaneRule = css.match(/\.native-powerpoint-canvas-pane\s*\{([\s\S]*?)\}/);

  assert.ok(canvasPaneRule, "the canvas pane must remain an explicit scroll owner");
  assert.match(canvasPaneRule[1], /overflow-anchor:\s*none/);
  assert.match(
    canvasPaneRule[1],
    /scrollbar-gutter:\s*stable/,
    "stable gutter prevents fitScale oscillation when overflow bars appear",
  );
});
test("shape fill menu uses the shared color popover and excludes unsupported objects", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const items = [];
  const menu = {
    addSeparator() { return menu; },
    addItem(build) {
      const item = {
        disabled: false,
        icon: null,
        title: null,
        action: null,
        setDisabled(value) { this.disabled = value; return this; },
        setIcon(value) { this.icon = value; return this; },
        setTitle(value) { this.title = value; return this; },
        onClick(value) { this.action = value; return this; },
      };
      build(item);
      items.push(item);
      return menu;
    },
  };
  const opened = [];

  view.t = (key) => key === "powerpoint:contextMenu.shapeFillColor" ? "Shape fill color…" : key;
  view.engine = { canSetShapeFillColor: () => true };
  view.openShapeFillColorPicker = (shapeIndex) => opened.push(shapeIndex);
  view.addShapeFillColorMenuItem(menu, 4, true);

  assert.equal(items.length, 1);
  assert.equal(items[0].title, "Shape fill color…");
  assert.equal(items[0].icon, "palette");
  assert.equal(items[0].disabled, false);
  items[0].action();
  assert.deepEqual(opened, [4]);

  view.engine = { canSetShapeFillColor: () => false };
  view.addShapeFillColorMenuItem(menu, 5, true);
  assert.equal(items.length, 1, "pictures, charts, tables, groups, and connectors should not offer fill color");
});
