import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadDocxParagraphLayoutRelayoutModule } from './helpers/load-plugin-modules.mjs';

async function loadModule() {
	return await loadDocxParagraphLayoutRelayoutModule();
}

function paragraph(defaultTextFormatting) {
	return {
		type: { name: 'paragraph' },
		attrs: { defaultTextFormatting },
		descendants(callback) {
			// A paragraph typically descends into text nodes, but we'll leave it empty for this mock.
			return true;
		},
	};
}

function doc(...nodes) {
	return {
		descendants(callback) {
			for (let i = 0; i < nodes.length; i++) {
				const node = nodes[i];
				const shouldContinue = callback(node, i);
				if (shouldContinue && node.descendants) {
					node.descendants(callback);
				}
			}
		},
	};
}

test('empty paragraph default font formatting invalidates paragraph layout', async () => {
	const { didParagraphTypographyChange } = await loadModule();
	assert.equal(
		didParagraphTypographyChange(
			doc(paragraph({ fontSize: 8 })),
			doc(paragraph({ fontSize: 24 })),
		),
		true,
	);
});

test('unchanged empty paragraph defaults keep layout stable', async () => {
	const { didParagraphTypographyChange } = await loadModule();
	assert.equal(
		didParagraphTypographyChange(
			doc(paragraph({ fontSize: 8, fontFamily: { ascii: 'Arial' } })),
			doc(paragraph({ fontSize: 8, fontFamily: { ascii: 'Arial' } })),
		),
		false,
	);
});

test('getDocumentParagraphLayoutSignatures extracts signatures for paragraph nodes only', async () => {
	const { getDocumentParagraphLayoutSignatures } = await loadModule();
	const mixedDoc = doc(
		paragraph({ fontSize: 12 }),
		{ type: { name: 'heading' }, descendants() {} },
		paragraph({ fontSize: 14 }),
	);
	const signatures = getDocumentParagraphLayoutSignatures(mixedDoc);
	assert.equal(signatures.length, 2);
	assert.equal(typeof signatures[0], 'string');
	assert.equal(typeof signatures[1], 'string');
	assert.notEqual(signatures[0], signatures[1]);
});

test('stableParagraphLayoutValue handles various types', async () => {
	const { stableParagraphLayoutValue } = await loadModule();

	// null/undefined
	assert.equal(stableParagraphLayoutValue(null), '');
	assert.equal(stableParagraphLayoutValue(undefined), '');

	// primitives
	assert.equal(stableParagraphLayoutValue('test'), 'test');
	assert.equal(stableParagraphLayoutValue(123), '123');
	assert.equal(stableParagraphLayoutValue(true), 'true');

	// objects
	assert.equal(stableParagraphLayoutValue({ a: 1 }), '{"a":1}');
	assert.equal(stableParagraphLayoutValue([1, 2, 3]), '[1,2,3]');

	// circular references
	const circular = {};
	circular.self = circular;
	assert.equal(stableParagraphLayoutValue(circular), '');
});

test('getParagraphListLayoutSignature handles empty and populated lists', async () => {
	const { getParagraphListLayoutSignature } = await loadModule();

	// Empty list attributes
	const emptyListDoc = {
		attrs: {}
	};
	assert.equal(getParagraphListLayoutSignature(emptyListDoc), '');

	// Populated list attributes
	const populatedListDoc = {
		attrs: {
			numPr: { numId: '1', ilvl: '0' },
			listMarker: '1.',
			listMarkerHidden: false,
			listMarkerFontFamily: { ascii: 'Arial' },
			listMarkerFontSize: 12,
			indentLeft: 720,
			indentFirstLine: 0,
			hangingIndent: 360
		}
	};

	const expected = [
		'{"numId":"1","ilvl":"0"}',
		'1.',
		'false',
		'{"ascii":"Arial"}',
		'12',
		'720',
		'0',
		'360'
	].join('\u001f');
	assert.equal(getParagraphListLayoutSignature(populatedListDoc), expected);
});

test('getParagraphTypographySignature handles formatting marks and omits non-typography marks', async () => {
	const { getParagraphTypographySignature } = await loadModule();

	const textNodeWithMarks = {
		isText: true,
		marks: [
			{ type: { name: 'bold' }, attrs: {} },
			{ type: { name: 'fontSize' }, attrs: { size: 24 } },
			{ type: { name: 'link' }, attrs: { href: 'https://example.com' } } // Should be omitted
		]
	};

	const docNode = {
		attrs: {
			styleId: 'Normal',
			alignment: 'center',
			lineSpacing: 240,
			lineSpacingRule: 'auto',
			spaceBefore: 120,
			spaceAfter: 120,
			indentLeft: 0,
			indentRight: 0,
			indentFirstLine: 0,
			hangingIndent: 0,
			outlineLevel: 0,
			defaultTextFormatting: {}
		},
		descendants(callback) {
			callback(textNodeWithMarks);
		}
	};

	const expected = [
		'Normal',
		'center',
		'240',
		'auto',
		'120',
		'120',
		'0',
		'0',
		'0',
		'0',
		'0',
		'{}',
		'bold:{};fontSize:{"size":24}'
	].join('\u001f');
	assert.equal(getParagraphTypographySignature(docNode), expected);
});

test('didListLayoutChange identifies unchanged and changed lists', async () => {
	const { didListLayoutChange } = await loadModule();

	function listParagraph(attrs) {
		return {
			type: { name: 'paragraph' },
			attrs,
			descendants(callback) { return true; }
		};
	}

	const listAttrs1 = { numPr: { numId: '1' } };
	const listAttrs2 = { numPr: { numId: '2' } };

	// Unchanged
	assert.equal(didListLayoutChange(doc(listParagraph(listAttrs1)), doc(listParagraph(listAttrs1))), false);

	// Changed
	assert.equal(didListLayoutChange(doc(listParagraph(listAttrs1)), doc(listParagraph(listAttrs2))), true);
});

test('didParagraphLayoutChange identifies changes in either list layout or typography', async () => {
	const { didParagraphLayoutChange } = await loadModule();

	function para(attrs) {
		return {
			type: { name: 'paragraph' },
			attrs,
			descendants(callback) { return true; }
		};
	}

	const para1 = para({ defaultTextFormatting: { fontSize: 12 }, numPr: { numId: '1' } });
	const para2 = para({ defaultTextFormatting: { fontSize: 14 }, numPr: { numId: '1' } }); // typography changed
	const para3 = para({ defaultTextFormatting: { fontSize: 12 }, numPr: { numId: '2' } }); // list changed
	const para4 = para({ defaultTextFormatting: { fontSize: 12 }, numPr: { numId: '1' } }); // unchanged

	assert.equal(didParagraphLayoutChange(doc(para1), doc(para2)), true);
	assert.equal(didParagraphLayoutChange(doc(para1), doc(para3)), true);
	assert.equal(didParagraphLayoutChange(doc(para1), doc(para4)), false);
});
