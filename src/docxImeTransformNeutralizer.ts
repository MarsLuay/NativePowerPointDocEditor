import { isHTMLElement } from './domGuards';

const IME_NEUTRALIZED_DATASET_KEY = 'docxidianImeNeutralized';

export interface ParsedEditorZoomTransform {
	translateXPx: number;
	scale: number;
}

export function parseEditorZoomTransform(transform: string): ParsedEditorZoomTransform {
	let translateXPx = 0;
	let scale = 1;
	const trimmed = transform.trim();

	if (!trimmed || trimmed === 'none') {
		return { translateXPx, scale };
	}

	const translateXMatch = trimmed.match(/translateX\(\s*(-?[\d.]+)px\s*\)/i);
	if (translateXMatch?.[1]) {
		translateXPx = Number.parseFloat(translateXMatch[1]);
	} else {
		const translateMatch = trimmed.match(/translate\(\s*(-?[\d.]+)px(?:\s*,\s*[^)]+)?\s*\)/i);
		if (translateMatch?.[1]) {
			translateXPx = Number.parseFloat(translateMatch[1]);
		}
	}

	const scaleMatch = trimmed.match(/scale\(\s*([\d.]+)(?:\s*,\s*[\d.]+)?\s*\)/i);
	if (scaleMatch?.[1]) {
		scale = Number.parseFloat(scaleMatch[1]);
	}

	return { translateXPx, scale };
}

export function editorZoomTransformNeedsNeutralization(transform: string): boolean {
	const { translateXPx, scale } = parseEditorZoomTransform(transform);
	return translateXPx !== 0 || Math.abs(scale - 1) > 0.0001;
}

export function findDocxEditorZoomWrapper(editorRoot: HTMLElement): HTMLElement | null {
	const pages = editorRoot.querySelector('.paged-editor__pages');
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

	const { translateXPx, scale } = parseEditorZoomTransform(transform);

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

	wrapper.style.removeProperty('transform');
	wrapper.style.removeProperty('transform-origin');
	wrapper.dataset[IME_NEUTRALIZED_DATASET_KEY] = 'true';

	return true;
}

export interface DocxImeTransformNeutralizerOptions {
	ownerDocument?: Document;
}

function getNeutralizerWindow(editorRoot: HTMLElement, ownerDocument?: Document): Window {
	return ownerDocument?.defaultView ?? editorRoot.ownerDocument.defaultView ?? window;
}

export function attachDocxImeTransformNeutralizer(
	editorRoot: HTMLElement,
	options: DocxImeTransformNeutralizerOptions = {},
): () => void {
	const view = getNeutralizerWindow(editorRoot, options.ownerDocument);
	let frameId: number | null = null;
	let pollIntervalId: number | null = null;
	let suppressObserver = false;
	const retryTimeouts: number[] = [];

	const run = () => {
		frameId = null;
		if (suppressObserver) {
			return;
		}

		const wrapper = findDocxEditorZoomWrapper(editorRoot);
		if (!wrapper) {
			return;
		}

		const transform = wrapper.style.transform;
		if (!editorZoomTransformNeedsNeutralization(transform)) {
			return;
		}

		suppressObserver = true;
		try {
			neutralizeDocxEditorZoomWrapper(wrapper);
		} finally {
			view.requestAnimationFrame(() => {
				suppressObserver = false;
			});
		}
	};

	const schedule = () => {
		if (frameId !== null) {
			return;
		}

		frameId = view.requestAnimationFrame(run);
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

	schedule();
	for (const delay of [0, 100, 500, 1500]) {
		retryTimeouts.push(view.setTimeout(schedule, delay));
	}
	pollIntervalId = view.setInterval(schedule, 500);

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
	};
}
