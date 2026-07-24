import assert from 'node:assert/strict';
import { test } from 'node:test';
import JSZip from 'jszip';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

async function createDocxBuffer(parts) {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', '<Types/>');
	for (const [partPath, xml] of Object.entries(parts)) {
		zip.file(partPath, xml);
	}
	return zip.generateAsync({ type: 'arraybuffer' });
}

function wrapBody(...inner) {
	return [
		'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',
		...inner,
		'</w:body></w:document>',
	].join('');
}

let cachedDocxDescribeModule;

async function loadDocxDescribeModule() {
	if (cachedDocxDescribeModule) return cachedDocxDescribeModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-docx-describe-test-'));
	const outfile = path.join(outputDirectory, 'docx-describe.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/docxDescribe.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cachedDocxDescribeModule = require(outfile);
	return cachedDocxDescribeModule;
}

test('describeDocxFromBuffer maps paragraphs, styles, and runs', async () => {
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();
	const buffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Abstract</w:t></w:r></w:p>',
			'<w:p><w:r><w:t>Body text</w:t></w:r></w:p>',
		),
	});

	const snapshot = await describeDocxFromBuffer(buffer, 'notes/doc.docx');
	assert.equal(snapshot.format, 'docx');
	assert.equal(snapshot.file, 'notes/doc.docx');
	assert.equal(snapshot.blockCount, 2);
	assert.equal(snapshot.blocks[0]?.id, 'body/p[0]');
	assert.equal(snapshot.blocks[0]?.kind, 'paragraph');
	assert.equal(snapshot.blocks[0]?.style, 'Heading1');
	assert.equal(snapshot.blocks[0]?.text, 'Abstract');
	assert.equal(snapshot.blocks[0]?.runs[0]?.id, 'body/p[0]/r[0]');
	assert.equal(snapshot.blocks[0]?.runs[0]?.bold, true);
	assert.equal(snapshot.blocks[1]?.id, 'body/p[1]');
	assert.equal(snapshot.blocks[1]?.text, 'Body text');
});

test('describeDocxFromBuffer maps table cells with stable ids', async () => {
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();
	const buffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:tbl>',
			'<w:tr>',
			'<w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>',
			'<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>',
			'</w:tr>',
			'<w:tr>',
			'<w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>',
			'</w:tr>',
			'</w:tbl>',
		),
	});

	const snapshot = await describeDocxFromBuffer(buffer, 'table.docx');
	assert.equal(snapshot.blockCount, 1);
	const table = snapshot.blocks[0];
	assert.equal(table?.id, 'body/tbl[0]');
	assert.equal(table?.kind, 'table');
	assert.equal(table?.rows, 2);
	assert.equal(table?.cols, 2);
	assert.equal(table?.cells?.[0]?.id, 'body/tbl[0]/tr[0]/tc[0]');
	assert.equal(table?.cells?.[0]?.text, 'A1');
	assert.equal(table?.cells?.[2]?.id, 'body/tbl[0]/tr[1]/tc[0]');
	assert.equal(table?.cells?.[2]?.text, 'A2');
});

function wrapHeader(...inner) {
	return [
		'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
		...inner,
		'</w:hdr>',
	].join('');
}

function wrapFooter(...inner) {
	return [
		'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
		...inner,
		'</w:ftr>',
	].join('');
}

function wrapFootnotes(...inner) {
	return [
		'<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
		...inner,
		'</w:footnotes>',
	].join('');
}

test('describeDocxFromBuffer includes headers, footers, footnotes, and comments', async () => {
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();
	const buffer = await createDocxBuffer({
		'word/document.xml': wrapBody('<w:p><w:r><w:t>Body</w:t></w:r></w:p>'),
		'word/headers/header1.xml': wrapHeader('<w:p><w:r><w:t>Header text</w:t></w:r></w:p>'),
		'word/footers/footer1.xml': wrapFooter('<w:p><w:r><w:t>Footer text</w:t></w:r></w:p>'),
		'word/footnotes.xml': wrapFootnotes(
			'<w:footnote w:type="separator"><w:p/></w:footnote>',
			'<w:footnote w:id="1"><w:p><w:r><w:t>Footnote one</w:t></w:r></w:p></w:footnote>',
		),
		'word/comments.xml': [
			'<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
			'<w:comment w:id="0" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">',
			'<w:p><w:r><w:t>Please revise.</w:t></w:r></w:p>',
			'</w:comment>',
			'</w:comments>',
		].join(''),
	});

	const snapshot = await describeDocxFromBuffer(buffer, 'parts.docx');
	assert.ok(snapshot.scope.sources.includes('word/headers/header1.xml'));
	assert.ok(snapshot.scope.sources.includes('word/footers/footer1.xml'));
	assert.ok(snapshot.scope.sources.includes('word/footnotes.xml'));
	assert.ok(snapshot.scope.sources.includes('word/comments.xml'));

	const header = snapshot.blocks.find((block) => block.id === 'header/1/p[0]');
	const footer = snapshot.blocks.find((block) => block.id === 'footer/1/p[0]');
	const footnote = snapshot.blocks.find((block) => block.id === 'footnotes/fn[1]/p[0]');
	const comment = snapshot.blocks.find((block) => block.id === 'comments/c[0]');

	assert.equal(header?.kind, 'paragraph');
	assert.equal(header?.text, 'Header text');
	assert.equal(footer?.text, 'Footer text');
	assert.equal(footnote?.text, 'Footnote one');
	assert.equal(comment?.kind, 'comment');
	assert.equal(comment?.author, 'Reviewer');
	assert.equal(comment?.text, 'Please revise.');
	assert.equal(comment?.parentId, undefined);
});

test('describeDocxFromBuffer threads comment replies via commentsExtended', async () => {
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();
	const buffer = await createDocxBuffer({
		'word/document.xml': wrapBody('<w:p><w:r><w:t>Body</w:t></w:r></w:p>'),
		'word/comments.xml': [
			'<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">',
			'<w:comment w:id="1" w:author="Mars" w:date="2026-07-15T21:18:10Z">',
			'<w:p w14:paraId="AAAAAAA1"><w:r><w:t>make fancier</w:t></w:r></w:p>',
			'</w:comment>',
			'<w:comment w:id="2" w:author="Mars" w:date="2026-07-15T21:19:00Z">',
			'<w:p w14:paraId="AAAAAAA2"><w:r><w:t>add more polish</w:t></w:r></w:p>',
			'</w:comment>',
			'</w:comments>',
		].join(''),
		'word/commentsExtended.xml': [
			'<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">',
			'<w15:commentEx w15:paraId="AAAAAAA1" w15:done="0"/>',
			'<w15:commentEx w15:paraId="AAAAAAA2" w15:paraIdParent="AAAAAAA1" w15:done="0"/>',
			'</w15:commentsEx>',
		].join(''),
	});

	const snapshot = await describeDocxFromBuffer(buffer, 'reply.docx');
	assert.ok(snapshot.scope.sources.includes('word/commentsExtended.xml'));
	const parent = snapshot.blocks.find((block) => block.id === 'comments/c[1]');
	const reply = snapshot.blocks.find((block) => block.id === 'comments/c[2]');
	assert.equal(parent?.text, 'make fancier');
	assert.equal(parent?.parentId, undefined);
	assert.equal(reply?.text, 'add more polish');
	assert.equal(reply?.parentId, 1);
});
