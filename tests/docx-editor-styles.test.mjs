import assert from 'node:assert/strict';
import test from 'node:test';
import { pathToFileURL } from 'node:url';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

let stylesModulePromise;

async function loadStylesModule() {
	stylesModulePromise ??= bundleSource('src/docxEditorStyles.ts', 'docx-editor-styles.cjs');
	return import(pathToFileURL(await stylesModulePromise).href);
}

function createRealm(name) {
	const sheets = [];
	class RealmCSSStyleSheet {
		constructor() {
			this.realm = name;
			this.cssText = '';
			sheets.push(this);
		}

		replaceSync(cssText) {
			this.cssText = cssText;
		}
	}

	const document = {
		adoptedStyleSheets: [],
		defaultView: { CSSStyleSheet: RealmCSSStyleSheet },
	};
	return { document, CSSStyleSheet: RealmCSSStyleSheet, sheets };
}

test('DOCX styles use the target document realm and remain document-scoped', async () => {
	const { ensureDocxDocumentStyles } = await loadStylesModule();
	const mainRealm = createRealm('main');
	const popoutRealm = createRealm('popout');
	const previousGlobalStyleSheet = globalThis.CSSStyleSheet;
	globalThis.CSSStyleSheet = mainRealm.CSSStyleSheet;

	try {
		assert.equal(
			ensureDocxDocumentStyles(mainRealm.document, '.main {}').method,
			'adoptedStyleSheets',
		);
		assert.equal(
			ensureDocxDocumentStyles(mainRealm.document, '.main-updated {}').method,
			'already-injected',
		);

		const popoutResult = ensureDocxDocumentStyles(popoutRealm.document, '.popout {}');
		assert.equal(popoutResult.method, 'adoptedStyleSheets');
		assert.equal(mainRealm.document.adoptedStyleSheets.length, 1);
		assert.equal(popoutRealm.document.adoptedStyleSheets.length, 1);
		assert.ok(popoutRealm.document.adoptedStyleSheets[0] instanceof popoutRealm.CSSStyleSheet);
		assert.equal(popoutRealm.document.adoptedStyleSheets[0].realm, 'popout');
		assert.equal(mainRealm.sheets.length, 1);
		assert.equal(popoutRealm.sheets.length, 1);
	} finally {
		globalThis.CSSStyleSheet = previousGlobalStyleSheet;
	}
});

test('DOCX styles fall back when adopted stylesheet ownership is rejected', async () => {
	const { ensureDocxDocumentStyles } = await loadStylesModule();
	let styleElement;
	const adoptedStyleSheets = [];
	const document = {
		defaultView: {
			CSSStyleSheet: class {
				replaceSync() {}
			},
		},
		querySelector() {
			return styleElement ?? null;
		},
		createElement() {
			return {
				setAttribute() {},
				textContent: '',
			};
		},
		head: {
			appendChild(element) {
				styleElement = element;
			},
		},
	};
	document.head.ownerDocument = document;
	Object.defineProperty(document, 'adoptedStyleSheets', {
		get: () => adoptedStyleSheets,
		set: () => {
			throw new Error('cross-document stylesheet');
		},
	});

	const result = ensureDocxDocumentStyles(document, '.fallback {}');
	assert.equal(result.method, 'style-element');
	assert.equal(result.fallbackReason, 'cross-document stylesheet');
	assert.equal(styleElement.textContent, '.fallback {}');
});
