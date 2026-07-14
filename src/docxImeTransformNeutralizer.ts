import { isHTMLElement } from './domGuards';
import { DOCX_CARET_SELECTOR, DOCX_EDITOR_PAGES_SELECTOR, DOCX_HIDDEN_PROSEMIRROR_SELECTOR } from './docxEditorChromeMarkers';
import type { RenderedDomContext } from '@eigenpal/docx-editor-core/plugin-api';
import type { EditorView } from 'prosemirror-view';

const IME_NEUTRALIZED_DATASET_KEY = 'nativePowerPointDocEditorImeNeutralized';
const IME_ANCHORED_DATASET_KEY = 'nativePowerPointDocEditorImeAnchored';
const IME_ORIGINAL_LEFT_DATASET_KEY = 'nativePowerPointDocEditorImeOriginalLeft';
const IME_ORIGINAL_TOP_DATASET_KEY = 'nativePowerPointDocEditorImeOriginalTop';
const IME_ORIGINAL_POSITION_DATASET_KEY = 'nativePowerPointDocEditorImeOriginalPosition';

export interface ParsedEditorZoomTransform {
	translateXPx: number;
	translateYPx: number;
	scale: number;
}

interface ViewportCaretRect {
	left: number;
	top: number;
	bottom: number;
	height: number;
}

interface CaretRectLike {
	left: number;
	top: number;
	bottom: number;
	height?: number;
}

interface HiddenImeAnchorPosition {
	leftPx: number;
	topPx: number;
}

function isFiniteNumber(value: number): boolean {
	return Number.isFinite(value);
}

function parseCssNumberList(value: string): number[] {
	return value
		.split(',')
		.map((part) => Number.parseFloat(part.trim()))
		.filter(Number.isFinite);
}

function parseCssPx(value: string): number | null {
	const parsed = Number.parseFloat(value);
	return Number.isFinite(parsed) ? parsed : null;
}

function setDynamicCssProps(element: HTMLElement, props: Record<string, string>): void {
	if (typeof element.setCssProps === 'function') {
		element.setCssProps(props);
		return;
	}

	Object.assign(element.style, props);
}

export function parseEditorZoomTransform(transform: string): ParsedEditorZoomTransform {
	let translateXPx = 0;
	let translateYPx = 0;
	let scale = 1;
	const trimmed = transform.trim();

	if (!trimmed || trimmed === 'none') {
		return { translateXPx, translateYPx, scale };
	}

	const matrix3dMatch = trimmed.match(/matrix3d\(\s*([^)]+)\)/i);
	if (matrix3dMatch?.[1]) {
		const values = parseCssNumberList(matrix3dMatch[1]);
		if (values.length === 16) {
			translateXPx = values[12] ?? 0;
			translateYPx = values[13] ?? 0;
			scale = Math.hypot(values[0] ?? 1, values[1] ?? 0);
			return { translateXPx, translateYPx, scale };
		}
	}

	const matrixMatch = trimmed.match(/matrix\(\s*([^)]+)\)/i);
	if (matrixMatch?.[1]) {
		const values = parseCssNumberList(matrixMatch[1]);
		if (values.length === 6) {
			translateXPx = values[4] ?? 0;
			translateYPx = values[5] ?? 0;
			scale = Math.hypot(values[0] ?? 1, values[1] ?? 0);
			return { translateXPx, translateYPx, scale };
		}
	}

	const translateXMatch = trimmed.match(/translateX\(\s*(-?[\d.]+)px\s*\)/i);
	if (translateXMatch?.[1]) {
		translateXPx = Number.parseFloat(translateXMatch[1]);
	} else {
		const translateMatch = trimmed.match(/translate(?:3d)?\(\s*(-?[\d.]+)px(?:\s*,\s*(-?[\d.]+)px)?/i);
		if (translateMatch?.[1]) {
			translateXPx = Number.parseFloat(translateMatch[1]);
		}
		if (translateMatch?.[2]) {
			translateYPx = Number.parseFloat(translateMatch[2]);
		}
	}

	const translateYMatch = trimmed.match(/translateY\(\s*(-?[\d.]+)px\s*\)/i);
	if (translateYMatch?.[1]) {
		translateYPx = Number.parseFloat(translateYMatch[1]);
	}

	const scaleMatch = trimmed.match(/scale(?:3d)?\(\s*([\d.]+)(?:\s*,\s*[\d.]+)?/i);
	if (scaleMatch?.[1]) {
		scale = Number.parseFloat(scaleMatch[1]);
	}

	return { translateXPx, translateYPx, scale };
}

export function editorZoomTransformNeedsNeutralization(transform: string): boolean {
	const { translateXPx, translateYPx, scale } = parseEditorZoomTransform(transform);
	return translateXPx !== 0 || translateYPx !== 0 || Math.abs(scale - 1) > 0.0001;
}

export function findDocxEditorZoomWrapper(editorRoot: HTMLElement): HTMLElement | null {
	const pages = editorRoot.querySelector(DOCX_EDITOR_PAGES_SELECTOR);
	const parent = pages?.parentElement;
	if (!parent || !editorRoot.contains(parent)) {
		return null;
	}

	return parent;
}

export function countTransformAncestors(element: Element | null, stopAt: Element | null = null): number {
	let count = 0;
	let current: Element | null = element;

	while (current && current !== stopAt) {
		const transform = isHTMLElement(current) ? current.style.transform : '';
		if (editorZoomTransformNeedsNeutralization(transform)) {
			count += 1;
		}
		current = current.parentElement;
	}

	return count;
}

export function neutralizeDocxEditorZoomWrapper(wrapper: HTMLElement): boolean {
	const transform = wrapper.style.transform;
	if (!editorZoomTransformNeedsNeutralization(transform)) {
		return false;
	}

	const { translateXPx, translateYPx, scale } = parseEditorZoomTransform(transform);

	if (Math.abs(scale - 1) > 0.0001) {
		wrapper.style.setProperty('zoom', String(scale));
	} else {
		wrapper.style.removeProperty('zoom');
	}

	if (translateXPx !== 0) {
		wrapper.style.setProperty('margin-left', `${translateXPx}px`);
	} else {
		wrapper.style.removeProperty('margin-left');
	}

	if (translateYPx !== 0) {
		wrapper.style.setProperty('margin-top', `${translateYPx}px`);
	} else {
		wrapper.style.removeProperty('margin-top');
	}

	wrapper.style.removeProperty('transform');
	wrapper.style.removeProperty('transform-origin');
	wrapper.dataset[IME_NEUTRALIZED_DATASET_KEY] = 'true';

	return true;
}

export interface DocxImeTransformNeutralizerOptions {
	getEditorView?: () => EditorView | null | undefined;
	getRenderedDomContext?: () => RenderedDomContext | null | undefined;
	ownerDocument?: Document;
	onDiagnostic?: (event: DocxImeDiagnosticEvent) => void;
}

export interface DocxImeDiagnosticEvent {
	event:
		| 'attached'
		| 'transform-neutralized'
		| 'anchor-state'
		| 'composition-start'
		| 'composition-end'
		| 'detached';
	details?: Record<string, unknown>;
}

function getNeutralizerWindow(editorRoot: HTMLElement, ownerDocument?: Document): Window {
	return ownerDocument?.defaultView ?? editorRoot.ownerDocument.defaultView ?? window;
}

function toViewportCaretRect(rect: CaretRectLike | null | undefined): ViewportCaretRect | null {
	if (!rect || !isFiniteNumber(rect.left) || !isFiniteNumber(rect.top) || !isFiniteNumber(rect.bottom)) {
		return null;
	}

	const explicitHeight = rect.height;
	const height = explicitHeight !== undefined && explicitHeight > 0
		? explicitHeight
		: Math.max(1, rect.bottom - rect.top);
	if (!isFiniteNumber(height) || height <= 0) {
		return null;
	}

	return {
		left: rect.left,
		top: rect.top,
		bottom: rect.bottom,
		height,
	};
}

function findVisibleCaretRect(editorRoot: HTMLElement): ViewportCaretRect | null {
	const caret = editorRoot.querySelector<HTMLElement>(DOCX_CARET_SELECTOR);
	return toViewportCaretRect(caret?.getBoundingClientRect());
}

function getRenderedCaretRect(view: EditorView, context: RenderedDomContext | null | undefined): ViewportCaretRect | null {
	if (!context?.pagesContainer?.isConnected) {
		return null;
	}

	const position = view.state.selection.head;
	const coordinates = context.getCoordinatesForPosition(position);
	if (!coordinates) {
		return null;
	}

	const zoom = Number.isFinite(context.zoom) && context.zoom > 0 ? context.zoom : 1;
	const pagesRect = context.pagesContainer.getBoundingClientRect();
	const top = pagesRect.top + coordinates.y * zoom;
	const height = Math.max(1, coordinates.height * zoom);
	return {
		left: pagesRect.left + coordinates.x * zoom,
		top,
		bottom: top + height,
		height,
	};
}

function findHiddenProseMirrorRoot(view: EditorView): HTMLElement | null {
	const hiddenRoot = view.dom.closest(DOCX_HIDDEN_PROSEMIRROR_SELECTOR);
	return isHTMLElement(hiddenRoot) ? hiddenRoot : null;
}

function hasHiddenProseMirrorFocus(view: EditorView, hiddenRoot: HTMLElement, ownerDocument: Document): boolean {
	if (view.hasFocus()) {
		return true;
	}

	const activeElement = ownerDocument.activeElement;
	return Boolean(activeElement && hiddenRoot.contains(activeElement));
}

function restoreHiddenProseMirrorAnchor(hiddenRoot: HTMLElement): boolean {
	if (hiddenRoot.dataset[IME_ANCHORED_DATASET_KEY] !== 'true') {
		return false;
	}

	const originalLeft = hiddenRoot.dataset[IME_ORIGINAL_LEFT_DATASET_KEY] ?? '-9999px';
	const originalTop = hiddenRoot.dataset[IME_ORIGINAL_TOP_DATASET_KEY] ?? '0';
	const originalPosition = hiddenRoot.dataset[IME_ORIGINAL_POSITION_DATASET_KEY] ?? '';
	hiddenRoot.style.setProperty('left', originalLeft);
	hiddenRoot.style.setProperty('top', originalTop);
	if (originalPosition) {
		hiddenRoot.style.setProperty('position', originalPosition);
	} else {
		hiddenRoot.style.removeProperty('position');
	}
	delete hiddenRoot.dataset[IME_ANCHORED_DATASET_KEY];
	delete hiddenRoot.dataset[IME_ORIGINAL_LEFT_DATASET_KEY];
	delete hiddenRoot.dataset[IME_ORIGINAL_TOP_DATASET_KEY];
	delete hiddenRoot.dataset[IME_ORIGINAL_POSITION_DATASET_KEY];
	return true;
}

function emitImeDiagnostic(
	options: DocxImeTransformNeutralizerOptions,
	event: DocxImeDiagnosticEvent['event'],
	details?: Record<string, unknown>,
): void {
	options.onDiagnostic?.({ event, details });
}

function roundedCaretRect(rect: ViewportCaretRect | null): Record<string, number> | null {
	if (!rect) {
		return null;
	}

	return {
		left: Math.round(rect.left),
		top: Math.round(rect.top),
		bottom: Math.round(rect.bottom),
		height: Math.round(rect.height),
	};
}

function getHiddenAnchorSnapshot(
	editorRoot: HTMLElement,
	options: DocxImeTransformNeutralizerOptions,
): Record<string, unknown> {
	const view = options.getEditorView?.();
	const hiddenRoot = view ? findHiddenProseMirrorRoot(view) : null;
	return {
		hasEditorView: Boolean(view),
		hasHiddenRoot: Boolean(hiddenRoot),
		anchored: hiddenRoot?.dataset[IME_ANCHORED_DATASET_KEY] === 'true',
		left: hiddenRoot?.style.left || null,
		top: hiddenRoot?.style.top || null,
		hasVisibleCaret: Boolean(findVisibleCaretRect(editorRoot)),
		hasRenderedDomContext: Boolean(options.getRenderedDomContext?.()?.pagesContainer?.isConnected),
	};
}

export function calculateHiddenImeAnchorPosition(
	currentHiddenLeftPx: number,
	currentHiddenTopPx: number,
	hiddenCaretRect: ViewportCaretRect,
	targetCaretRect: ViewportCaretRect,
): HiddenImeAnchorPosition | null {
	if (
		!isFiniteNumber(currentHiddenLeftPx)
		|| !isFiniteNumber(currentHiddenTopPx)
		|| !isFiniteNumber(hiddenCaretRect.left)
		|| !isFiniteNumber(hiddenCaretRect.bottom)
		|| !isFiniteNumber(targetCaretRect.left)
		|| !isFiniteNumber(targetCaretRect.bottom)
	) {
		return null;
	}

	return {
		leftPx: currentHiddenLeftPx + targetCaretRect.left - hiddenCaretRect.left,
		topPx: currentHiddenTopPx + targetCaretRect.bottom - hiddenCaretRect.bottom,
	};
}

export function syncDocxImeHiddenProseMirrorAnchor(
	editorRoot: HTMLElement,
	options: DocxImeTransformNeutralizerOptions = {},
): boolean {
	const view = options.getEditorView?.();
	if (!view) {
		emitImeDiagnostic(options, 'anchor-state', {
			status: 'unavailable',
			reason: 'editor-view-missing',
		});
		return false;
	}

	const ownerDocument = options.ownerDocument ?? editorRoot.ownerDocument;
	const hiddenRoot = findHiddenProseMirrorRoot(view);
	if (!hiddenRoot) {
		emitImeDiagnostic(options, 'anchor-state', {
			status: 'unavailable',
			reason: 'hidden-prosemirror-root-missing',
		});
		return false;
	}

	if (!hasHiddenProseMirrorFocus(view, hiddenRoot, ownerDocument)) {
		const restored = restoreHiddenProseMirrorAnchor(hiddenRoot);
		emitImeDiagnostic(options, 'anchor-state', {
			status: 'inactive',
			reason: 'editor-not-focused',
			restored,
		});
		return false;
	}

	const visibleCaretRect = findVisibleCaretRect(editorRoot);
	const renderedCaretRect = visibleCaretRect
		? null
		: getRenderedCaretRect(view, options.getRenderedDomContext?.());
	const targetCaretRect = visibleCaretRect ?? renderedCaretRect;
	if (!targetCaretRect) {
		emitImeDiagnostic(options, 'anchor-state', {
			status: 'unavailable',
			reason: 'rendered-caret-missing',
		});
		return false;
	}

	let hiddenCaretRect: ViewportCaretRect | null = null;
	try {
		hiddenCaretRect = toViewportCaretRect(view.coordsAtPos(view.state.selection.head));
	} catch {
		hiddenCaretRect = null;
	}
	if (!hiddenCaretRect) {
		emitImeDiagnostic(options, 'anchor-state', {
			status: 'unavailable',
			reason: 'hidden-caret-coordinates-missing',
		});
		return false;
	}

	const currentHiddenRect = hiddenRoot.getBoundingClientRect();
	const currentHiddenLeftPx = parseCssPx(hiddenRoot.style.left) ?? currentHiddenRect.left;
	const currentHiddenTopPx = parseCssPx(hiddenRoot.style.top) ?? currentHiddenRect.top;
	const nextAnchor = calculateHiddenImeAnchorPosition(
		currentHiddenLeftPx,
		currentHiddenTopPx,
		hiddenCaretRect,
		targetCaretRect,
	);
	if (!nextAnchor) {
		emitImeDiagnostic(options, 'anchor-state', {
			status: 'unavailable',
			reason: 'anchor-calculation-failed',
		});
		return false;
	}

	if (hiddenRoot.dataset[IME_ANCHORED_DATASET_KEY] !== 'true') {
		hiddenRoot.dataset[IME_ORIGINAL_LEFT_DATASET_KEY] = hiddenRoot.style.left || '-9999px';
		hiddenRoot.dataset[IME_ORIGINAL_TOP_DATASET_KEY] = hiddenRoot.style.top || '0';
		hiddenRoot.dataset[IME_ORIGINAL_POSITION_DATASET_KEY] = hiddenRoot.style.position || '';
	}
	setDynamicCssProps(hiddenRoot, {
		position: 'fixed',
		left: `${Math.round(nextAnchor.leftPx)}px`,
		top: `${Math.round(nextAnchor.topPx)}px`,
	});
	hiddenRoot.dataset[IME_ANCHORED_DATASET_KEY] = 'true';
	emitImeDiagnostic(options, 'anchor-state', {
		status: 'synced',
		reason: visibleCaretRect ? 'visible-caret' : 'rendered-dom-context',
		targetCaret: roundedCaretRect(targetCaretRect),
		hiddenCaret: roundedCaretRect(hiddenCaretRect),
		anchor: {
			left: Math.round(nextAnchor.leftPx),
			top: Math.round(nextAnchor.topPx),
		},
	});
	return true;
}

export function attachDocxImeTransformNeutralizer(
	editorRoot: HTMLElement,
	options: DocxImeTransformNeutralizerOptions = {},
): () => void {
	const view = getNeutralizerWindow(editorRoot, options.ownerDocument);
	const ownerDocument = options.ownerDocument ?? editorRoot.ownerDocument;
	let frameId: number | null = null;
	let pollIntervalId: number | null = null;
	let suppressObserver = false;
	const retryTimeouts: number[] = [];
	let lastAnchorState = '';

	const diagnosticOptions: DocxImeTransformNeutralizerOptions = {
		...options,
		onDiagnostic: (event) => {
			if (event.event === 'anchor-state') {
				const status = typeof event.details?.status === 'string' ? event.details.status : '';
				const reason = typeof event.details?.reason === 'string' ? event.details.reason : '';
				const signature = `${status}:${reason}`;
				if (signature === lastAnchorState) {
					return;
				}
				lastAnchorState = signature;
			}
			options.onDiagnostic?.(event);
		},
	};

	const run = () => {
		frameId = null;
		if (suppressObserver) {
			return false;
		}

		const wrapper = findDocxEditorZoomWrapper(editorRoot);
		if (!wrapper) {
			return false;
		}

		const transform = wrapper.style.transform;
		if (editorZoomTransformNeedsNeutralization(transform)) {
			const parsedTransform = parseEditorZoomTransform(transform);
			suppressObserver = true;
			try {
				neutralizeDocxEditorZoomWrapper(wrapper);
				emitImeDiagnostic(diagnosticOptions, 'transform-neutralized', {
					originalTransform: transform,
					translateXPx: parsedTransform.translateXPx,
					translateYPx: parsedTransform.translateYPx,
					scale: parsedTransform.scale,
					zoom: wrapper.style.zoom || null,
					marginLeft: wrapper.style.marginLeft || null,
					marginTop: wrapper.style.marginTop || null,
				});
			} finally {
				view.requestAnimationFrame(() => {
					suppressObserver = false;
				});
			}
		}

		return syncDocxImeHiddenProseMirrorAnchor(editorRoot, diagnosticOptions);
	};

	const schedule = () => {
		if (frameId !== null) {
			return;
		}

		frameId = view.requestAnimationFrame(() => {
			run();
		});
	};

	const runImmediately = (event: Event) => {
		if (frameId !== null) {
			view.cancelAnimationFrame(frameId);
			frameId = null;
		}
		const anchored = run();
		if (event.type === 'compositionstart' || event.type === 'compositionend') {
			emitImeDiagnostic(
				diagnosticOptions,
				event.type === 'compositionstart' ? 'composition-start' : 'composition-end',
				{
					anchored,
					...getHiddenAnchorSnapshot(editorRoot, diagnosticOptions),
				},
			);
		}
		schedule();
	};

	const observer = new MutationObserver((mutations) => {
		if (suppressObserver) {
			return;
		}

		for (const mutation of mutations) {
			if (mutation.type === 'childList') {
				schedule();
				return;
			}

			if (
				mutation.type === 'attributes'
				&& mutation.attributeName === 'style'
				&& isHTMLElement(mutation.target)
			) {
				const wrapper = findDocxEditorZoomWrapper(editorRoot);
				if (wrapper && (mutation.target === wrapper || wrapper.contains(mutation.target))) {
					schedule();
					return;
				}
			}
		}
	});

	observer.observe(editorRoot, {
		attributes: true,
		attributeFilter: ['style'],
		childList: true,
		subtree: true,
	});

	const immediateEventOptions = { capture: true };
	const scheduledEventOptions = { capture: true };
	for (const eventName of ['keydown', 'beforeinput', 'compositionstart', 'compositionupdate', 'compositionend', 'input']) {
		ownerDocument.addEventListener(eventName, runImmediately, immediateEventOptions);
	}
	for (const eventName of ['keyup', 'mouseup', 'focusin', 'focusout', 'selectionchange']) {
		ownerDocument.addEventListener(eventName, schedule, scheduledEventOptions);
	}

	schedule();
	for (const delay of [0, 100, 500, 1500]) {
		retryTimeouts.push(view.setTimeout(schedule, delay));
	}
	pollIntervalId = view.setInterval(schedule, 1000);
	emitImeDiagnostic(diagnosticOptions, 'attached', {
		hasZoomWrapper: Boolean(findDocxEditorZoomWrapper(editorRoot)),
		...getHiddenAnchorSnapshot(editorRoot, diagnosticOptions),
	});

	return () => {
		observer.disconnect();
		if (frameId !== null) {
			view.cancelAnimationFrame(frameId);
		}
		if (pollIntervalId !== null) {
			view.clearInterval(pollIntervalId);
		}
		for (const timeoutId of retryTimeouts) {
			if (timeoutId) {
				view.clearTimeout(timeoutId);
			}
		}
		for (const eventName of ['keydown', 'beforeinput', 'compositionstart', 'compositionupdate', 'compositionend', 'input']) {
			ownerDocument.removeEventListener(eventName, runImmediately, immediateEventOptions);
		}
		for (const eventName of ['keyup', 'mouseup', 'focusin', 'focusout', 'selectionchange']) {
			ownerDocument.removeEventListener(eventName, schedule, scheduledEventOptions);
		}
		const editorView = options.getEditorView?.();
		if (editorView) {
			const hiddenRoot = findHiddenProseMirrorRoot(editorView);
			if (hiddenRoot) {
				restoreHiddenProseMirrorAnchor(hiddenRoot);
			}
		}
		emitImeDiagnostic(diagnosticOptions, 'detached', getHiddenAnchorSnapshot(editorRoot, diagnosticOptions));
	};
}
