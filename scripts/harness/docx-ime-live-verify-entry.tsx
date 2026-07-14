import React, { useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor, type DocxEditorRef } from '@npde/docx-editor-react';
import editorStyles from '@npde/docx-editor-react/styles.css';
import type { RenderedDomContext } from '@npde/docx-editor-core/plugin-api';
import {
	attachDocxImeTransformNeutralizer,
	countTransformAncestors,
	findDocxEditorZoomWrapper,
	neutralizeDocxEditorZoomWrapper,
} from '../../src/docxImeTransformNeutralizer';

const DOCX_BUFFER = Uint8Array.from(atob(window.__DOCX_BASE64__!), (char) => char.charCodeAt(0)).buffer;

type WrapperInspection = {
	inlineTransform: string;
	inlineZoom: string;
	inlineMarginLeft: string;
	computedTransform: string;
	neutralizedDataset: boolean;
};

type LiveVerifyMetrics = {
	name: string;
	error?: string;
	neutralizerAttached?: boolean;
	wrapper: WrapperInspection | null;
	transformAncestorsOnCaret: number;
	editableFound: boolean;
	hiddenImeRootFound: boolean;
	hiddenImeAnchored: boolean;
	hiddenCaretDelta: { x: number; bottom: number } | null;
	compositionStartAnchored: boolean;
	passed: boolean;
};

function inspectWrapper(editorRoot: HTMLElement): WrapperInspection | null {
	const wrapper = findDocxEditorZoomWrapper(editorRoot);
	if (!wrapper) {
		return null;
	}

	return {
		inlineTransform: wrapper.style.transform || 'none',
		inlineZoom: wrapper.style.zoom || '',
		inlineMarginLeft: wrapper.style.marginLeft || '',
		computedTransform: getComputedStyle(wrapper).transform,
		neutralizedDataset: wrapper.dataset.nativePowerPointDocEditorImeNeutralized === 'true',
	};
}

function findBodyEditable(hostEl: HTMLElement): HTMLElement | null {
	return (
		hostEl.ownerDocument.querySelector<HTMLElement>('[data-native-powerpoint-doc-editor-hidden-prosemirror] .ProseMirror[contenteditable="true"]')
		?? hostEl.ownerDocument.querySelector<HTMLElement>('.paged-editor__hidden-pm .ProseMirror[contenteditable="true"]')
		?? hostEl.ownerDocument.querySelector<HTMLElement>('.ProseMirror[contenteditable="true"]')
		?? hostEl.querySelector<HTMLElement>('[contenteditable="true"]')
	);
}

function findNeutralizerRoot(hostEl: HTMLElement): HTMLElement | null {
	return (
		hostEl.querySelector<HTMLElement>('.native-powerpoint-doc-editor-editor-harness')
		?? hostEl.querySelector<HTMLElement>('[data-native-powerpoint-doc-editor-root]')
		?? hostEl.querySelector<HTMLElement>('[data-native-powerpoint-doc-editor-root]')
	);
}

function collectMetrics(
	scenarioName: string,
	hostEl: HTMLElement,
	editorRef: React.RefObject<DocxEditorRef>,
	compositionStartAnchored: boolean,
): LiveVerifyMetrics {
	const editorRoot = findNeutralizerRoot(hostEl);
	if (!editorRoot) {
		return {
			name: scenarioName,
			error: 'missing editor root',
			wrapper: null,
			transformAncestorsOnCaret: -1,
			editableFound: false,
			hiddenImeRootFound: false,
			hiddenImeAnchored: false,
			hiddenCaretDelta: null,
			compositionStartAnchored,
			passed: false,
		};
	}

	const editable = findBodyEditable(hostEl);
	const zoomWrapper = findDocxEditorZoomWrapper(editorRoot);
	if (zoomWrapper) {
		neutralizeDocxEditorZoomWrapper(zoomWrapper);
	}
	const wrapper = inspectWrapper(editorRoot);
	const transformAncestorsOnCaret = editable ? countTransformAncestors(editable) : -1;
	const view = editorRef.current?.getEditorRef()?.getView() ?? null;
	const hiddenRoot =
		view?.dom.closest<HTMLElement>('[data-native-powerpoint-doc-editor-hidden-prosemirror]')
		?? view?.dom.closest<HTMLElement>('.paged-editor__hidden-pm')
		?? null;
	const visibleCaret = editorRoot.querySelector<HTMLElement>('[data-native-powerpoint-doc-editor-caret]')
		?? editorRoot.querySelector<HTMLElement>('[data-testid="caret"]');
	let hiddenCaretDelta: LiveVerifyMetrics['hiddenCaretDelta'] = null;
	if (view && visibleCaret) {
		try {
			const hiddenCaretRect = view.coordsAtPos(view.state.selection.head);
			const visibleCaretRect = visibleCaret.getBoundingClientRect();
			hiddenCaretDelta = {
				x: Math.round(hiddenCaretRect.left - visibleCaretRect.left),
				bottom: Math.round(hiddenCaretRect.bottom - visibleCaretRect.bottom),
			};
		} catch {
			hiddenCaretDelta = null;
		}
	}
	const hiddenImeAnchored = hiddenRoot?.dataset.nativePowerPointDocEditorImeAnchored === 'true';

	const passed =
		wrapper !== null
		&& wrapper.inlineTransform === 'none'
		&& !wrapper.inlineTransform.includes('scale(')
		&& !wrapper.inlineTransform.includes('translateX(')
		&& transformAncestorsOnCaret === 0
		&& hiddenImeAnchored
		&& hiddenCaretDelta !== null
		&& Math.abs(hiddenCaretDelta.x) <= 2
		&& Math.abs(hiddenCaretDelta.bottom) <= 2
		&& compositionStartAnchored
		&& (wrapper.inlineZoom !== '' || wrapper.inlineMarginLeft !== '' || scenarioName === 'baseline');

	return {
		name: scenarioName,
		neutralizerAttached: editorRoot.dataset.nativePowerPointDocEditorImeAttachProbe === 'true',
		wrapper,
		transformAncestorsOnCaret,
		editableFound: Boolean(editable),
		hiddenImeRootFound: Boolean(hiddenRoot),
		hiddenImeAnchored,
		hiddenCaretDelta,
		compositionStartAnchored,
		passed,
	};
}

function DocxLiveVerifyApp({
	scenarioName,
	initialZoom,
	showOutline,
}: {
	scenarioName: string;
	initialZoom: number;
	showOutline: boolean;
}) {
	const [hostEl, setHostEl] = useState<HTMLDivElement | null>(null);
	const editorRef = useRef<DocxEditorRef>(null);
	const renderedDomContextRef = useRef<RenderedDomContext | null>(null);

	useEffect(() => {
		if (!hostEl) {
			return;
		}

		let detachNeutralizer: (() => void) | undefined;
		let retryTimeouts: number[] = [];
		let compositionStartAnchored = false;

		const attachNeutralizer = (): boolean => {
			if (detachNeutralizer) {
				return true;
			}
			const editorRoot = findNeutralizerRoot(hostEl);
			const editorView = editorRef.current?.getEditorRef()?.getView();
			if (!editorRoot || !editorView) {
				return false;
			}

			detachNeutralizer = attachDocxImeTransformNeutralizer(editorRoot, {
				getEditorView: () => editorRef.current?.getEditorRef()?.getView() ?? null,
				getRenderedDomContext: () => renderedDomContextRef.current,
				onDiagnostic: (event) => {
					if (event.event === 'composition-start' && event.details?.anchored === true) {
						compositionStartAnchored = true;
					}
				},
			});
			editorRoot.dataset.nativePowerPointDocEditorImeAttachProbe = 'true';
			return true;
		};

		const publish = () => {
			attachNeutralizer();
			editorRef.current?.focus();
			window.setTimeout(() => {
				document.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: 'に' }));
				window.setTimeout(() => {
					const metrics = collectMetrics(scenarioName, hostEl, editorRef, compositionStartAnchored);
					document.body.dataset.metrics = encodeURIComponent(JSON.stringify(metrics));
					console.log('LIVE_VERIFY_RESULT:' + JSON.stringify(metrics));
				}, 50);
			}, 50);
		};

		if (!attachNeutralizer()) {
			retryTimeouts = [100, 500, 1500].map((delay) => window.setTimeout(attachNeutralizer, delay));
		}

		const timers = [1000, 2500, 5000, 10000, 15000, 20000].map((delay) => window.setTimeout(publish, delay));

		return () => {
			for (const timeout of retryTimeouts) {
				window.clearTimeout(timeout);
			}
			retryTimeouts = [];
			for (const timer of timers) {
				window.clearTimeout(timer);
			}
			detachNeutralizer?.();
		};
	}, [hostEl, scenarioName]);

	return (
		<div className="native-powerpoint-doc-editor-host" ref={setHostEl}>
			<DocxEditor
				ref={editorRef}
				className="native-powerpoint-doc-editor-editor-harness"
				documentBuffer={DOCX_BUFFER}
				initialZoom={initialZoom}
				showOutline={showOutline}
				showOutlineButton={showOutline}
				showToolbar={false}
				showRuler={false}
				showZoomControl={false}
				onRenderedDomContextReady={(context) => {
					renderedDomContextRef.current = context;
				}}
				onFontsLoaded={() => undefined}
				onError={(error) => {
					const payload: LiveVerifyMetrics = {
						name: scenarioName,
						error: error.message,
						wrapper: null,
						transformAncestorsOnCaret: -1,
						editableFound: false,
						hiddenImeRootFound: false,
						hiddenImeAnchored: false,
						hiddenCaretDelta: null,
						compositionStartAnchored: false,
						passed: false,
					};
					document.body.dataset.metrics = encodeURIComponent(JSON.stringify(payload));
					console.log('LIVE_VERIFY_RESULT:' + JSON.stringify(payload));
				}}
			/>
		</div>
	);
}

declare global {
	interface Window {
		__DOCX_BASE64__?: string;
		__HARNESS_SCENARIO__?: string;
	}
}

const styleTag = document.createElement('style');
styleTag.textContent = `
${editorStyles}
html, body { margin: 0; width: 100%; height: 100%; overflow: hidden; background: #e2e8f0; }
.native-powerpoint-doc-editor-host { display: flex; flex-direction: column; width: 100%; height: 100%; min-height: 0; background: #f8fafc; }
`;
document.head.appendChild(styleTag);

const scenario = window.__HARNESS_SCENARIO__ ?? 'baseline';
const scenarioProps: Record<string, { initialZoom: number; showOutline: boolean }> = {
	baseline: { initialZoom: 1, showOutline: false },
	zoom125Outline: { initialZoom: 1.25, showOutline: true },
};

const props = scenarioProps[scenario] ?? scenarioProps.baseline;
const root = createRoot(document.getElementById('root')!);
root.render(
	<DocxLiveVerifyApp
		scenarioName={scenario}
		initialZoom={props.initialZoom}
		showOutline={props.showOutline}
	/>,
);
