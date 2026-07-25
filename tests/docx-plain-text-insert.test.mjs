import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Fragment, Schema, Slice } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { loadDocxPlainTextInsertModule } from "./helpers/load-plugin-modules.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const schema = new Schema({
	nodes: {
		doc: { content: "block+" },
		paragraph: {
			content: "inline*",
			group: "block",
			attrs: {
				defaultTextFormatting: { default: null },
				_originalFormatting: { default: null },
				styleId: { default: null },
				numPr: { default: null },
				listIsBullet: { default: null },
				listNumFmt: { default: null },
				listMarker: { default: null },
			},
		},
		text: { group: "inline" },
	},
	marks: {
		bold: {},
		fontFamily: {
			attrs: {
				ascii: { default: null },
				hAnsi: { default: null },
			},
		},
		fontSize: {
			attrs: { size: { default: null }, sizeCs: { default: null } },
		},
	},
});

function paragraph(text, attrs = null, marks = []) {
	return schema.node(
		"paragraph",
		attrs,
		text ? schema.text(text, marks) : undefined,
	);
}

function countParagraphs(doc) {
	let count = 0;
	doc.descendants((node) => {
		if (node.isTextblock) {
			count += 1;
		}
		return true;
	});
	return count;
}

test("resolveSafePlainTextInsertRange keeps same-parent replace intact", async () => {
	const { resolveSafePlainTextInsertRange } = await loadDocxPlainTextInsertModule();
	const doc = schema.node("doc", null, [
		paragraph("Python"),
		paragraph("TypeScript"),
	]);
	const state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, 1, 5),
	});

	const range = resolveSafePlainTextInsertRange(state, 1, 5);
	assert.deepEqual(range, { from: 1, to: 5, collapsedCrossBlock: false });
});

test("cross-block plain insert collapses selection and preserves paragraphs", async () => {
	const {
		buildPlainTextInsertTransaction,
		countDocTextblocks,
	} = await loadDocxPlainTextInsertModule();

	const doc = schema.node("doc", null, [
		paragraph("Python"),
		paragraph("TypeScript"),
		paragraph("Rust"),
	]);
	const from = 3;
	const to = doc.content.size - 3;
	const state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, from, to),
	});

	assert.equal(countDocTextblocks(state.doc), 3);
	assert.equal(state.selection.$from.sameParent(state.selection.$to), false);

	const unsafe = state.tr.insertText("SQL", from, to);
	assert.equal(countParagraphs(unsafe.doc), 1, "raw insertText must wipe intervening paragraphs");

	const { transaction, range } = buildPlainTextInsertTransaction(state, "SQL", from, to);
	assert.equal(range.collapsedCrossBlock, true);
	assert.equal(range.from, range.to);
	assert.equal(transaction.selection.empty, true, "safe insert should not leave the old range selected");
	assert.equal(countDocTextblocks(transaction.doc), 3);
	assert.match(transaction.doc.textContent, /SQL/);
	assert.match(transaction.doc.textContent, /TypeScript/);
	assert.match(transaction.doc.textContent, /Rust/);
});

test("multi-line plain paste keeps blank lines and empty-para fonts", async () => {
	const {
		buildPlainTextParagraphSlice,
		splitPlainTextLines,
	} = await loadDocxPlainTextInsertModule();

	const sample = [
		"TECHNICAL SKILLS",
		"",
		"",
		"Operating Systems: Windows 10/11, Linux (Debian)",
		"",
		"Programming & Scripting: Java, TypeScript",
	].join("\n");

	assert.deepEqual(splitPlainTextLines(sample).filter((line) => line === ""), ["", "", ""]);

	const fontMarks = [
		schema.marks.fontFamily.create({ ascii: "Georgia", hAnsi: "Georgia" }),
		schema.marks.fontSize.create({ size: 48, sizeCs: 48 }),
	];
	const doc = schema.node("doc", null, [
		paragraph("", {
			defaultTextFormatting: {
				fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
				fontSize: 48,
			},
		}),
	]);
	let state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, 1),
	});
	state = state.apply(state.tr.setStoredMarks(fontMarks));

	const { slice } = buildPlainTextParagraphSlice(state, sample, 1);
	const transaction = state.tr.replaceRange(1, 1, slice);
	const emptyParas = [];
	transaction.doc.forEach((node) => {
		if (node.isTextblock && node.textContent.length === 0) {
			emptyParas.push(node);
		}
	});

	assert.ok(emptyParas.length >= 3, `expected blank section paras, got ${emptyParas.length}`);
	for (const empty of emptyParas) {
		assert.deepEqual(empty.attrs.defaultTextFormatting, {
			fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
			fontSize: 48,
			fontSizeCs: 48,
		});
	}
	assert.match(transaction.doc.textContent, /TECHNICAL SKILLS/);
	assert.match(transaction.doc.textContent, /Operating Systems/);
});

test("multi-line plain paste inside a bullet inherits its list attrs", async () => {
	const { buildPlainTextParagraphSlice } = await loadDocxPlainTextInsertModule();
	const listAttrs = {
		numPr: { numId: 1, ilvl: 0 },
		listIsBullet: true,
		listNumFmt: "bullet",
		listMarker: "•",
	};
	const doc = schema.node("doc", null, [paragraph("Existing", listAttrs)]);
	const state = EditorState.create({ doc, selection: TextSelection.create(doc, 1) });

	const { slice } = buildPlainTextParagraphSlice(state, "Windows\nLinux", 1);

	slice.content.forEach((node) => {
		assert.deepEqual(node.attrs.numPr, { numId: 1, ilvl: 0 });
		assert.equal(node.attrs.listIsBullet, true);
		assert.equal(node.attrs.listMarker, "•");
	});
});

test("multi-line paste over cross-block selection replaces without wiping neighbors", async () => {
	const { buildPlainTextParagraphSlice, countDocTextblocks } = await loadDocxPlainTextInsertModule();

	const doc = schema.node("doc", null, [
		paragraph("BEFORE"),
		paragraph("Operating Systems: old"),
		paragraph("Programming: old"),
		paragraph("AFTER"),
	]);
	let from = -1;
	let to = -1;
	doc.descendants((node, pos) => {
		if (!node.isTextblock) return;
		if (node.textContent.startsWith("Operating Systems: old") && from < 0) {
			from = pos + 1;
		}
		if (node.textContent === "AFTER" && to < 0) {
			to = pos + 1; // start of AFTER paragraph content
		}
	});
	assert.ok(from > 0 && to > from, `expected cross-block range, got ${from}-${to}`);

	const state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, from, to),
	});
	assert.equal(state.selection.$from.sameParent(state.selection.$to), false);

	const sample = ["TECHNICAL SKILLS", "", "Operating Systems: Windows"].join("\n");
	const { slice } = buildPlainTextParagraphSlice(state, sample, from);
	const transaction = state.tr.replaceRange(from, to, slice);

	assert.ok(countDocTextblocks(transaction.doc) >= 4);
	assert.match(transaction.doc.textContent, /BEFORE/);
	assert.match(transaction.doc.textContent, /AFTER/);
	assert.match(transaction.doc.textContent, /TECHNICAL SKILLS/);
	assert.match(transaction.doc.textContent, /Operating Systems: Windows/);
	assert.doesNotMatch(transaction.doc.textContent, /Programming: old/);
});

test("multi-line paste collapses the old cross-block selection after insertion", async () => {
	const { insertPlainTextAsParagraphs } = await loadDocxPlainTextInsertModule();
	const doc = schema.node("doc", null, [
		paragraph("BEFORE"),
		paragraph("replace one"),
		paragraph("replace two"),
		paragraph("AFTER"),
	]);
	let from = -1;
	let to = -1;
	doc.descendants((node, pos) => {
		if (!node.isTextblock) return;
		if (node.textContent === "replace one") from = pos + 1;
		if (node.textContent === "replace two") to = pos + node.nodeSize - 1;
	});
	const state = EditorState.create({ doc, selection: TextSelection.create(doc, from, to) });
	let dispatched = null;
	const view = {
		state,
		dispatch(transaction) {
			dispatched = transaction;
		},
	};

	insertPlainTextAsParagraphs(view, "new first\nnew second", from, to);

	assert.ok(dispatched, "paste should dispatch a transaction");
	assert.equal(dispatched.selection.empty, true, "paste should leave a caret, not select inserted text");
	assert.match(dispatched.selection.$from.parent.textContent, /new second/);
});

test("rich cross-block paste preserves per-run marks and replaces the full selection", async () => {
	const { insertRichClipboardSlice } = await loadDocxPlainTextInsertModule();
	const doc = schema.node("doc", null, [
		paragraph("BEFORE"),
		paragraph("replace one"),
		paragraph("replace two"),
		paragraph("AFTER"),
	]);
	let from = -1;
	let to = -1;
	doc.descendants((node, pos) => {
		if (!node.isTextblock) return;
		if (node.textContent === "replace one") from = pos + 1;
		if (node.textContent === "AFTER") to = pos + 1;
	});
	const state = EditorState.create({ doc, selection: TextSelection.create(doc, from, to) });
	const bold = schema.marks.bold.create();
	const fontSize = schema.marks.fontSize.create({ size: 36, sizeCs: 36 });
	const slice = new Slice(
		Fragment.from([
			paragraph("TECHNICAL SKILLS", null, [bold, fontSize]),
			paragraph("", {
				defaultTextFormatting: { fontSize: 36, fontSizeCs: 36 },
			}),
			paragraph("Operating Systems: Windows", {
				numPr: { numId: 1, ilvl: 0 },
				listIsBullet: true,
				listNumFmt: "bullet",
				listMarker: "•",
			}),
		]),
		1,
		1,
	);
	let dispatched = null;
	const view = {
		state,
		dispatch(transaction) {
			dispatched = transaction;
		},
	};

	const range = insertRichClipboardSlice(view, slice, from, to);

	assert.equal(range.collapsedCrossBlock, true);
	assert.ok(dispatched, "rich paste should dispatch a transaction");
	assert.equal(dispatched.selection.empty, true, "rich paste should leave a caret");
	assert.match(dispatched.doc.textContent, /BEFORE/);
	assert.match(dispatched.doc.textContent, /AFTER/);
	assert.match(dispatched.doc.textContent, /TECHNICAL SKILLS/);
	assert.match(dispatched.doc.textContent, /Operating Systems: Windows/);
	assert.doesNotMatch(dispatched.doc.textContent, /replace one|replace two/);

	let title = null;
	dispatched.doc.descendants((node) => {
		if (node.isTextblock && node.textContent === "TECHNICAL SKILLS") {
			title = node;
			return false;
		}
		return true;
	});
	assert.ok(title, "expected formatted title paragraph");
	assert.ok(title.firstChild.marks.some((mark) => mark.type.name === "bold"), "title must stay bold");

	let emptyParagraph = null;
	let bulletParagraph = null;
	dispatched.doc.descendants((node) => {
		if (!node.isTextblock) return true;
		if (node.textContent.length === 0) emptyParagraph = node;
		if (node.textContent === "Operating Systems: Windows") bulletParagraph = node;
		return true;
	});
	assert.deepEqual(emptyParagraph?.attrs.defaultTextFormatting, { fontSize: 36, fontSizeCs: 36 });
	assert.deepEqual(bulletParagraph?.attrs.numPr, { numId: 1, ilvl: 0 });
	assert.equal(bulletParagraph?.attrs.listIsBullet, true);
});

test("DOCX paste preserves rich HTML and routes only plain text through paragraph insert", async () => {
	const source = await readFile(path.join(projectRoot, "src/DocxReactView.tsx"), "utf8");
	const helper = await readFile(path.join(projectRoot, "src/docxPlainTextInsert.ts"), "utf8");
	const menus = await readFile(
		path.join(
			projectRoot,
			"docx-editor/packages/react/src/components/DocxEditor/hooks/useContextMenus.ts",
		),
		"utf8",
	);

	assert.match(source, /from '\.\/docxPlainTextInsert'/);
	assert.match(source, /insertPlainTextAsParagraphs/);
	assert.match(source, /DOCX multi-line plain paste inserted as paragraphs/);
	assert.match(source, /multi-line-paragraphs/);
	assert.match(source, /if \(hasHtml\)/);
	assert.match(source, /DOCX paste deferred to HTML clipboardParser/);
	assert.match(source, /handlePaste\(view, event, slice: Slice\)/);
	assert.match(source, /insertRichClipboardSlice/);
	assert.match(source, /isSuggestionModeActive/);
	assert.match(source, /if \(!crossBlock\)/);
	assert.match(source, /before the section heading/);
	assert.match(source, /DOCX rich paste replaced through structured slice/);
	assert.match(source, /summarizeRichClipboardSlice/);
	assert.match(
		source,
		/routing multi-line or cross-block HTML through the plain\s+\/\/ text helper silently strips those marks/,
	);
	assert.match(helper, /splitPlainTextLines/);
	assert.match(helper, /buildPlainTextParagraphSlice/);
	assert.match(helper, /replacedCrossBlock|crossBlock/);
	assert.match(helper, /TextSelection\.near/);
	assert.match(helper, /summarizeTransactionMeta/);
	assert.match(source, /DOCX keyboard history shortcut received/);
	assert.match(menus, /pasteAsPlainText/);
	assert.match(menus, /setData\('text\/plain'/);
	assert.doesNotMatch(menus, /view\.dispatch\(view\.state\.tr\.insertText\(text\)\)/);
});
