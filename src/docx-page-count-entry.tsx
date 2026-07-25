import React, { useEffect, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor } from './docx/runtime/bridge.mjs';
import type { DocxEditorRef } from './docx/runtime/contract';
import { docxEditorRuntimeStyles as editorStyles } from './docx/runtime/styles';

const documentBuffer = Uint8Array.from(atob(window.__DOCX_BASE64__!), (char) => char.charCodeAt(0)).buffer;

function PageCountApp() {
	const editorRef = useRef<DocxEditorRef>(null);

	useEffect(() => {
		let cancelled = false;
		const deadline = Date.now() + 20000;
		const publish = () => {
			if (cancelled) {
				return;
			}
			const totalPages = editorRef.current?.getTotalPages() ?? 0;
			const renderedPages = window.document.querySelectorAll('[data-native-powerpoint-doc-editor-page]').length;
			if ((totalPages > 0 && renderedPages > 0) || Date.now() >= deadline) {
				const layout = editorRef.current?.getEditorRef()?.getLayout();
				const pageText = Array.from(window.document.querySelectorAll<HTMLElement>('[data-native-powerpoint-doc-editor-page]')).map(
					(page) => page.textContent?.replace(/\s+/g, ' ').trim() ?? '',
				);
				const pageFragments = layout?.pages.map((page) => page.fragments.map((fragment) => ({
					kind: fragment.kind,
					blockId: fragment.blockId,
					y: Math.round(fragment.y),
					height: Math.round(fragment.height),
					...fragment.kind === 'paragraph' ? {
						fromLine: fragment.fromLine,
						toLine: fragment.toLine,
						continuesFromPrev: fragment.continuesFromPrev === true,
						continuesOnNext: fragment.continuesOnNext === true,
					} : {},
				}))) ?? [];
				const metrics = { totalPages, renderedPages, pageText, pageFragments };
				window.document.body.dataset.metrics = encodeURIComponent(JSON.stringify(metrics));
				console.log(`DOCX_PAGE_COUNT_RESULT:${JSON.stringify(metrics)}`);
				return;
			}
			window.setTimeout(publish, 100);
		};
		window.setTimeout(publish, 100);
		return () => {
			cancelled = true;
		};
	}, []);

	return (
		<DocxEditor
			ref={editorRef}
			documentBuffer={documentBuffer}
			mode="viewing"
			readOnly
			showToolbar={false}
			showRuler={false}
			showZoomControl={false}
		/>
	);
}

declare global {
	interface Window {
		__DOCX_BASE64__?: string;
	}
}

const styleTag = window.document.createElement('style');
styleTag.textContent = `
${editorStyles}
html, body, #root { margin: 0; width: 100%; min-height: 100%; background: #e2e8f0; }
`;
window.document.head.appendChild(styleTag);

createRoot(window.document.getElementById('root')!).render(<PageCountApp />);
