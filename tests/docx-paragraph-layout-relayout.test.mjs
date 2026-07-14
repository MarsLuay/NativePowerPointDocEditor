import assert from "node:assert/strict";
import { test } from "node:test";
import { Schema } from "prosemirror-model";
import { loadDocxParagraphLayoutRelayoutModule } from "./helpers/load-plugin-modules.mjs";

const schema = new Schema({
	nodes: {
		doc: { content: "block+" },
		paragraph: {
			attrs: {
				styleId: { default: null },
				alignment: { default: null },
				lineSpacing: { default: null },
				lineSpacingRule: { default: null },
				spaceBefore: { default: null },
				spaceAfter: { default: null },
				indentLeft: { default: null },
				indentRight: { default: null },
				indentFirstLine: { default: null },
				hangingIndent: { default: null },
				outlineLevel: { default: null },
				defaultFontSize: { default: null },
				defaultFontFamily: { default: null },
				numPr: { default: null },
				listMarker: { default: null },
			},
			content: "inline*",
			group: "block",
		},
		text: { group: "inline" },
	},
	marks: {
		bold: {},
		italic: {},
		fontSize: { attrs: { size: { default: null }, sizeCs: { default: null } } },
	},
});

function paragraph(attrs, text, marks = []) {
	return schema.node("paragraph", attrs, text ? schema.text(text, marks) : undefined);
}

test("didParagraphTypographyChange detects heading style changes", async () => {
	const {
		didParagraphTypographyChange,
		didParagraphLayoutChange,
		didListLayoutChange,
	} = await loadDocxParagraphLayoutRelayoutModule();

	const before = schema.node("doc", null, [
		paragraph({ styleId: "Normal" }, "Hello"),
		paragraph({ styleId: "Normal" }, "World"),
	]);
	const after = schema.node("doc", null, [
		paragraph({ styleId: "Heading1" }, "Hello", [schema.marks.bold.create()]),
		paragraph({ styleId: "Heading1" }, "World", [schema.marks.bold.create()]),
	]);

	assert.equal(didListLayoutChange(before, after), false);
	assert.equal(didParagraphTypographyChange(before, after), true);
	assert.equal(didParagraphLayoutChange(before, after), true);
});

test("didParagraphLayoutChange ignores unchanged documents", async () => {
	const { didParagraphLayoutChange } = await loadDocxParagraphLayoutRelayoutModule();
	const doc = schema.node("doc", null, [paragraph({ styleId: "Normal" }, "Hello")]);
	assert.equal(didParagraphLayoutChange(doc, doc), false);
});

test("stableParagraphLayoutValue avoids default object stringification", async () => {
	const { stableParagraphLayoutValue } = await loadDocxParagraphLayoutRelayoutModule();
	assert.equal(stableParagraphLayoutValue({ level: 2 }), '{"level":2}');
	const circular = {};
	circular.self = circular;
	assert.equal(stableParagraphLayoutValue(circular), "");
});
