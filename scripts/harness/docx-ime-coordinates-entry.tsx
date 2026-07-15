import React, { useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor } from '@npde/docx-editor-react';
import editorStyles from '@npde/docx-editor-react/styles.css';

const DOCX_BUFFER = Uint8Array.from(atob(window.__DOCX_BASE64__), (char) => char.charCodeAt(0)).buffer;

type TransformAncestor = {
	tag: string;
	className: string;
	transform: string;
};

type ScenarioMetrics = {
	name: string;
	fixedProbeLeft: number;
	fixedProbeTop: number;
	zoomContainerTransform: string | null;
	transformAncestors: TransformAncestor[];
	editableCount: number;
	caretClientRect: { left: number; top: number; bottom: number; right: number } | null;
	caretScreenPoint: { x: number; y: number } | null;
	windowScreenOrigin: { x: number; y: number };
	imeRisk: 'low' | 'high';
};

function collectTransformAncestors(node: Element | null): TransformAncestor[] {
	const ancestors: TransformAncestor[] = [];
	let current = node?.parentElement ?? null;
	while (current && current !== document.documentElement) {
		const transform = getComputedStyle(current).transform;
		if (transform && transform !== 'none') {
			ancestors.push({
				tag: current.tagName.toLowerCase(),
				className: current.className?.toString?.() ?? '',
				transform,
			});
		}
		current = current.parentElement;
	}
	return ancestors;
}

function collectCaretMetrics(hostEl: HTMLElement): Pick<
	ScenarioMetrics,
	'editableCount' | 'caretClientRect' | 'caretScreenPoint' | 'transformAncestors'
> {
	const editables = Array.from(hostEl.querySelectorAll<HTMLElement>('[contenteditable="true"]'));
	const editable =
		editables.find((element) => element.closest('[data-native-powerpoint-doc-editor-page-content]')) ??
		editables.find((element) => !element.closest('[data-native-powerpoint-doc-editor-hidden-prosemirror]')) ??
		editables[0] ??
		null;

	if (!editable) {
		return {
			editableCount: editables.length,
			caretClientRect: null,
			caretScreenPoint: null,
			transformAncestors: [],
		};
	}

	editable.focus();
	const text = editable.textContent ?? '';
	const caretOffset = Math.max(0, text.length);
	const selection = window.getSelection();
	const range = document.createRange();
	const textNode = editable.firstChild;
	if (textNode && textNode.nodeType === Node.TEXT_NODE) {
		range.setStart(textNode, Math.min(caretOffset, textNode.textContent?.length ?? 0));
	} else {
		range.selectNodeContents(editable);
	}
	range.collapse(true);
	selection?.removeAllRanges();
	selection?.addRange(range);

	const rect = range.getBoundingClientRect();
	return {
		editableCount: editables.length,
		caretClientRect: {
			left: rect.left,
			top: rect.top,
			right: rect.right,
			bottom: rect.bottom,
		},
		caretScreenPoint: {
			x: rect.left + window.screenX,
			y: rect.bottom + window.screenY,
		},
		transformAncestors: collectTransformAncestors(editable),
	};
}

function collectScenarioMetrics(name: string, hostEl: HTMLElement): ScenarioMetrics {
	const fixedProbe = document.createElement('div');
	fixedProbe.className = 'native-powerpoint-doc-editor-fixed-probe';
	fixedProbe.style.position = 'fixed';
	fixedProbe.style.left = '0';
	fixedProbe.style.top = '0';
	fixedProbe.style.visibility = 'hidden';
	fixedProbe.style.pointerEvents = 'none';
	hostEl.appendChild(fixedProbe);
	const fixedRect = fixedProbe.getBoundingClientRect();
	fixedProbe.remove();

	const zoomContainer = hostEl.querySelector<HTMLElement>('[data-native-powerpoint-doc-editor-root] > div[style*="transform"]');
	const caretMetrics = collectCaretMetrics(hostEl);
	const hasRiskyTransform = caretMetrics.transformAncestors.length > 0;
	const hasFixedProbeOffset = Math.abs(fixedRect.left) > 1 || Math.abs(fixedRect.top) > 1;

	return {
		name,
		fixedProbeLeft: Math.round(fixedRect.left),
		fixedProbeTop: Math.round(fixedRect.top),
		zoomContainerTransform: zoomContainer?.style.transform ?? null,
		...caretMetrics,
		windowScreenOrigin: { x: window.screenX, y: window.screenY },
		imeRisk: hasRiskyTransform || hasFixedProbeOffset ? 'high' : 'low',
	};
}

function DocxHarnessApp({
	scenarioName,
	workspaceStyle,
	initialZoom,
	showOutline,
}: {
	scenarioName: string;
	workspaceStyle?: React.CSSProperties;
	initialZoom?: number;
	showOutline?: boolean;
}) {
	const hostRef = useRef<HTMLDivElement | null>(null);
	const [ready, setReady] = useState(false);

	useEffect(() => {
		if (!ready || !hostRef.current) {
			return;
		}

		const publish = () => {
			const metrics = collectScenarioMetrics(scenarioName, hostRef.current!);
			document.body.dataset.metrics = encodeURIComponent(JSON.stringify(metrics));
		};

		const timer = window.setTimeout(publish, 2500);
		return () => window.clearTimeout(timer);
	}, [ready, scenarioName]);

	return (
		<div className="workspace-shell" style={workspaceStyle}>
			<div className="native-powerpoint-doc-editor-host" ref={hostRef}>
				<DocxEditor
					documentBuffer={DOCX_BUFFER}
					initialZoom={initialZoom ?? 1}
					showOutline={showOutline ?? false}
					showOutlineButton={showOutline ?? false}
					showToolbar={false}
					showRuler={false}
					showZoomControl={false}
					onFontsLoaded={() => setReady(true)}
					onError={(error) => {
						document.body.dataset.metrics = encodeURIComponent(JSON.stringify({
							name: scenarioName,
							error: error.message,
							imeRisk: 'high',
						}));
					}}
				/>
			</div>
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
.workspace-shell { width: 100%; height: 100%; padding: 24px; box-sizing: border-box; }
.native-powerpoint-doc-editor-host { display: flex; flex-direction: column; width: 100%; height: 100%; min-height: 0; background: #f8fafc; }
.native-powerpoint-doc-editor-fixed-probe { left: 0; position: fixed; top: 0; visibility: hidden; pointer-events: none; }
`;
document.head.appendChild(styleTag);

const scenario = window.__HARNESS_SCENARIO__ ?? 'baseline';
const root = createRoot(document.getElementById('root')!);

const scenarioProps: Record<string, React.ComponentProps<typeof DocxHarnessApp>> = {
	baseline: { scenarioName: 'baseline', initialZoom: 1, showOutline: false },
	zoom125: { scenarioName: 'zoom125', initialZoom: 1.25, showOutline: false },
	outline: { scenarioName: 'outline', initialZoom: 1, showOutline: true },
	zoom125Outline: { scenarioName: 'zoom125Outline', initialZoom: 1.25, showOutline: true },
	obsidianOffset: {
		scenarioName: 'obsidianOffset',
		initialZoom: 1,
		showOutline: false,
		workspaceStyle: { transform: 'translate(180px, 96px)' },
	},
	fullStack: {
		scenarioName: 'fullStack',
		initialZoom: 1.25,
		showOutline: true,
		workspaceStyle: { transform: 'translate(180px, 96px)' },
	},
};

root.render(<DocxHarnessApp {...(scenarioProps[scenario] ?? scenarioProps.baseline)} />);
