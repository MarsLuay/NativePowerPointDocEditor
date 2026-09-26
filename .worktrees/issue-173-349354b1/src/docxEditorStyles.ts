import { createDetachedDocxEditorChromeElement } from './docxEditorChromeDom';

export type DocxStyleInjectionMethod = 'already-injected' | 'adoptedStyleSheets' | 'style-element';

export interface DocxStyleInjectionResult {
	method: DocxStyleInjectionMethod;
	fallbackReason?: string;
}

const injectedDocuments = new WeakSet<Document>();
const STYLE_ELEMENT_ATTRIBUTE = 'data-native-powerpoint-doc-editor-styles';

function getErrorMessage(error: unknown): string {
	return error instanceof Error ? error.message : String(error);
}

function injectWithStyleElement(targetDocument: Document, cssText: string): void {
	const existing = targetDocument.querySelector<HTMLStyleElement>(
		`style[${STYLE_ELEMENT_ATTRIBUTE}]`,
	);
	if (existing) {
		existing.textContent = cssText;
		return;
	}

	const parent = targetDocument.head ?? targetDocument.documentElement ?? targetDocument.body;
	if (!parent) {
		throw new Error('DOCX editor styles require a document root.');
	}
	const style = createDetachedDocxEditorChromeElement(parent, 'style');
	style.setAttribute(STYLE_ELEMENT_ATTRIBUTE, 'true');
	style.textContent = cssText;
	parent.appendChild(style);
}

export function ensureDocxDocumentStyles(
	targetDocument: Document,
	cssText: string,
): DocxStyleInjectionResult {
	if (injectedDocuments.has(targetDocument)) {
		return { method: 'already-injected' };
	}

	const targetStyleSheetConstructor = targetDocument.defaultView?.CSSStyleSheet
		?? (typeof CSSStyleSheet !== 'undefined' ? CSSStyleSheet : undefined);
	let fallbackReason: string | undefined;

	if (
		targetStyleSheetConstructor
		&& 'adoptedStyleSheets' in targetDocument
		&& Array.isArray(targetDocument.adoptedStyleSheets)
	) {
		try {
			const styleSheet = new targetStyleSheetConstructor();
			styleSheet.replaceSync(cssText);
			targetDocument.adoptedStyleSheets = [
				...targetDocument.adoptedStyleSheets,
				styleSheet,
			];
			injectedDocuments.add(targetDocument);
			return { method: 'adoptedStyleSheets' };
		} catch (error) {
			fallbackReason = getErrorMessage(error);
		}
	} else {
		fallbackReason = 'adoptedStyleSheets is unavailable for the target document';
	}

	injectWithStyleElement(targetDocument, cssText);
	injectedDocuments.add(targetDocument);
	return {
		method: 'style-element',
		fallbackReason,
	};
}
