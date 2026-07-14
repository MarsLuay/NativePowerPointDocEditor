import React, { useEffect, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor, type DocxEditorRef } from '@eigenpal/docx-editor-react';
import editorStyles from '@eigenpal/docx-editor-react/styles.css';
import JSZip from 'jszip';
import { preserveDocxTableCellFontSizes } from '../../src/docxTableCellFontSizePreserver';

const sourceBuffer = Uint8Array.from(atob(window.__DOCX_BASE64__!), (char) => char.charCodeAt(0)).buffer;

interface FontRoundTripMetrics {
	renderedCellFontSizes: Record<string, string | null>;
	editApplied: boolean;
	rawCellCount: number;
	raw24PtSizeCount: number;
	raw24PtComplexSizeCount: number;
	rawTable11PtSizeCount: number;
	rawTable11PtComplexSizeCount: number;
	repaired24PtSizeCount: number;
	repaired24PtComplexSizeCount: number;
	preservationStatus: string;
	restoredRuns: number;
	restoredTags: number;
}

async function documentXml(buffer: ArrayBuffer): Promise<string> {
	const zip = await JSZip.loadAsync(buffer.slice(0));
	const part = zip.file('word/document.xml');
	if (!part) {
		throw new Error('word/document.xml is missing');
	}
	return part.async('string');
}

function countMatches(value: string, pattern: RegExp): number {
	return value.match(pattern)?.length ?? 0;
}

function findExactTextElement(text: string): HTMLElement | null {
	const elements = document.querySelectorAll<HTMLElement>('[data-native-powerpoint-doc-editor-pages] *');
	for (const element of elements) {
		if (element.children.length === 0 && element.textContent?.trim() === text) {
			return element;
		}
	}
	return null;
}

function collectRenderedCellFontSizes(): Record<string, string | null> {
	const expectedCellText = [
		'A1', 'A2', 'A3', 'A4',
		'B1', 'B2', 'B3', 'B4',
		'Size', 'Keep', 'Wide', 'Cell',
		'C1', 'C2', 'C3', 'C4',
		'D1', 'D2', 'D3', 'D4',
	];
	return Object.fromEntries(
		expectedCellText.map((text) => {
			const element = findExactTextElement(text);
			return [text, element ? getComputedStyle(element).fontSize : null];
		}),
	);
}

function FontRoundTripApp() {
	const editorRef = useRef<DocxEditorRef>(null);

	useEffect(() => {
		let cancelled = false;

		const run = async () => {
			const deadline = Date.now() + 15000;
			while (!cancelled && !editorRef.current?.getEditorRef()?.getView() && Date.now() < deadline) {
				await new Promise((resolve) => window.setTimeout(resolve, 100));
			}
			if (cancelled || !editorRef.current) {
				return;
			}

			await new Promise((resolve) => window.setTimeout(resolve, 750));
			const renderedCellFontSizes = collectRenderedCellFontSizes();
			const editorView = editorRef.current.getEditorRef()?.getView();
			let editFrom = -1;
			let editTo = -1;
			editorView?.state.doc.descendants((node, position) => {
				if (node.isText && node.text?.includes('Document defaults are intentionally 11 pt')) {
					editFrom = position;
					editTo = position + node.nodeSize;
					return false;
				}
				return true;
			});
			const italicMark = editorView?.state.schema.marks.italic;
			const editApplied = Boolean(editorView && italicMark && editFrom >= 0 && editTo > editFrom);
			if (editorView && italicMark && editApplied) {
				editorView.dispatch(editorView.state.tr.addMark(editFrom, editTo, italicMark.create()));
			}
			const rawBuffer = await editorRef.current.save({ selective: false });
			if (!rawBuffer) {
				throw new Error('DOCX editor returned no save buffer');
			}

			const rawXml = await documentXml(rawBuffer);
			const rawTableXml = (rawXml.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) ?? []).join('');
			const preserved = await preserveDocxTableCellFontSizes(sourceBuffer, rawBuffer);
			const repairedXml = await documentXml(preserved.buffer);
			const metrics: FontRoundTripMetrics = {
				renderedCellFontSizes,
				editApplied,
				rawCellCount: countMatches(rawXml, /<w:tc\b/g),
				raw24PtSizeCount: countMatches(rawXml, /<w:sz w:val="48"\/>/g),
				raw24PtComplexSizeCount: countMatches(rawXml, /<w:szCs w:val="48"\/>/g),
				rawTable11PtSizeCount: countMatches(rawTableXml, /<w:sz w:val="22"\/>/g),
				rawTable11PtComplexSizeCount: countMatches(rawTableXml, /<w:szCs w:val="22"\/>/g),
				repaired24PtSizeCount: countMatches(repairedXml, /<w:sz w:val="48"\/>/g),
				repaired24PtComplexSizeCount: countMatches(repairedXml, /<w:szCs w:val="48"\/>/g),
				preservationStatus: preserved.status,
				restoredRuns: preserved.restoredRuns,
				restoredTags: preserved.restoredTags,
			};
			document.body.dataset.metrics = encodeURIComponent(JSON.stringify(metrics));
			console.log(`FONT_ROUNDTRIP_RESULT:${JSON.stringify(metrics)}`);
		};

		void run().catch((error) => {
			const message = error instanceof Error ? error.message : String(error);
			document.body.dataset.error = encodeURIComponent(message);
			console.error(`FONT_ROUNDTRIP_ERROR:${message}`);
		});

		return () => {
			cancelled = true;
		};
	}, []);

	return (
		<DocxEditor
			ref={editorRef}
			documentBuffer={sourceBuffer}
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

const styleTag = document.createElement('style');
styleTag.textContent = `
${editorStyles}
html, body, #root { margin: 0; width: 100%; min-height: 100%; background: #e2e8f0; }
`;
document.head.appendChild(styleTag);

createRoot(document.getElementById('root')!).render(<FontRoundTripApp />);
