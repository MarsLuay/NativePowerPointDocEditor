import assert from 'node:assert/strict';
import { test } from 'node:test';
import JSZip from 'jszip';
import { build } from 'esbuild';
import { docxEditorAliases } from './helpers/docx-esbuild-aliases.mjs';
import { mkdtemp, mkdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const COMMENT_PARTS = [
	'word/comments.xml',
	'word/commentsExtended.xml',
	'word/commentsIds.xml',
	'word/commentsExtensible.xml',
];
const COMMENT_RELATIONSHIP_PARTS = COMMENT_PARTS.map((partPath) => `word/_rels/${partPath.split('/').pop()}.rels`);
const COMMENT_RELATIONSHIP_TYPES = [
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
	'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
	'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
	'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
];

let cachedDocxServiceModule;
let cachedDescribeModule;
let cachedObsidianStub;

async function loadDocxServiceModule() {
	if (cachedDocxServiceModule) return cachedDocxServiceModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-docx-remove-comments-'));
	const obsidianModuleDirectory = path.join(outputDirectory, 'node_modules', 'obsidian');
	await mkdir(obsidianModuleDirectory, { recursive: true });
	await writeFile(path.join(obsidianModuleDirectory, 'package.json'), '{"name":"obsidian","main":"index.js"}');
	await writeFile(
		path.join(obsidianModuleDirectory, 'index.js'),
		'class TFile { constructor(path, extension) { this.path = path; this.extension = extension; } } module.exports = { TFile };',
	);
	cachedObsidianStub = require(path.join(obsidianModuleDirectory, 'index.js'));
	const outfile = path.join(outputDirectory, 'docx-service.cjs');
	await build({
		alias: docxEditorAliases,
		absWorkingDir: outputDirectory,
		entryPoints: [path.join(projectRoot, 'src/ai/docxDocumentService.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		external: ['obsidian'],
	});
	cachedDocxServiceModule = require(outfile);
	return cachedDocxServiceModule;
}

async function loadDocxDescribeModule() {
	if (cachedDescribeModule) return cachedDescribeModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-docx-remove-comments-describe-'));
	const outfile = path.join(outputDirectory, 'docx-describe.cjs');
	await build({
		alias: docxEditorAliases,
		entryPoints: [path.join(projectRoot, 'src/ai/docxDescribe.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cachedDescribeModule = require(outfile);
	return cachedDescribeModule;
}

function createMockVault(initialFiles) {
	const { TFile } = cachedObsidianStub;
	const store = new Map(initialFiles);
	return {
		store,
		getAbstractFileByPath(filePath) {
			if (!store.has(filePath)) return null;
			return new TFile(filePath, filePath.split('.').pop() ?? '');
		},
		async readBinary(file) {
			const bytes = store.get(file.path);
			if (!bytes) throw new Error(`Missing file: ${file.path}`);
			return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
		},
		async modifyBinary(file, buffer) {
			store.set(file.path, Buffer.from(buffer));
		},
	};
}

function anchoredParagraph(text, id = 1) {
	return [
		'<w:p>',
		`<w:commentRangeStart w:id="${id}"/>`,
		`<w:r><w:t>${text}</w:t></w:r>`,
		`<w:commentRangeEnd w:id="${id}"/>`,
		`<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`,
		'</w:p>',
	].join('');
}

async function createCommentedDocxBuffer() {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', [
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
		'<Default Extension="xml" ContentType="application/xml"/>',
		'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
		'<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
		'<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>',
		'<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>',
		'<Override PartName="/word/commentsExtensible.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"/>',
		'</Types>',
	].join(''));
	zip.file('_rels/.rels', [
		`<Relationships xmlns="${RELS_NS}">`,
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
		'</Relationships>',
	].join(''));
	zip.file('word/_rels/document.xml.rels', [
		`<Relationships xmlns="${RELS_NS}">`,
		`<Relationship Id="rId2" Type="${COMMENT_RELATIONSHIP_TYPES[0]}" Target="comments.xml"/>`,
		`<Relationship Id="rId3" Type="${COMMENT_RELATIONSHIP_TYPES[1]}" Target="commentsExtended.xml"/>`,
		`<Relationship Id="rId4" Type="${COMMENT_RELATIONSHIP_TYPES[2]}" Target="commentsIds.xml"/>`,
		`<Relationship Id="rId5" Type="${COMMENT_RELATIONSHIP_TYPES[3]}" Target="commentsExtensible.xml"/>`,
		'<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/kept.png"/>',
		'</Relationships>',
	].join(''));
	zip.file('word/document.xml', [
		`<w:document xmlns:w="${WORD_NS}"><w:body>`,
		anchoredParagraph('Body text'),
		'<w:p><w:ins w:id="9" w:author="Reviewer"><w:r><w:t>Tracked change stays</w:t></w:r></w:ins></w:p>',
		'</w:body></w:document>',
	].join(''));
	zip.file('word/headers/header1.xml', `<w:hdr xmlns:w="${WORD_NS}">${anchoredParagraph('Header text')}</w:hdr>`);
	zip.file('word/footers/footer1.xml', `<w:ftr xmlns:w="${WORD_NS}">${anchoredParagraph('Footer text')}</w:ftr>`);
	zip.file('word/footnotes.xml', `<w:footnotes xmlns:w="${WORD_NS}"><w:footnote w:id="1">${anchoredParagraph('Footnote text')}</w:footnote></w:footnotes>`);
	zip.file('word/endnotes.xml', `<w:endnotes xmlns:w="${WORD_NS}"><w:endnote w:id="1">${anchoredParagraph('Endnote text')}</w:endnote></w:endnotes>`);
	zip.file('word/comments.xml', [
		`<w:comments xmlns:w="${WORD_NS}">`,
		'<w:comment w:id="1" w:author="Mars"><w:p><w:r><w:t>make fancier</w:t></w:r></w:p></w:comment>',
		'</w:comments>',
	].join(''));
	zip.file('word/commentsExtended.xml', '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="1"/></w15:commentsEx>');
	zip.file('word/commentsIds.xml', '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="1" w16cid:durableId="1"/></w16cid:commentsIds>');
	zip.file('word/commentsExtensible.xml', '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"><w16cex:comment w16cex:durableId="1"/></w16cex:commentsExtensible>');
	zip.file('word/_rels/comments.xml.rels', `<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/comment-only.png"/></Relationships>`);
	zip.file('word/media/kept.png', Buffer.from([1, 2, 3]));
	zip.file('word/media/comment-only.png', Buffer.from([4, 5, 6]));
	return zip.generateAsync({ type: 'arraybuffer' });
}

test('docx.removeComments removes annotations and package metadata without touching unrelated content', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();
	const docPath = 'notes/comments.docx';
	const initialBuffer = await createCommentedDocxBuffer();
	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const dryRun = await service.apply(docPath, [{ op: 'docx.removeComments' }], { dryRun: true });
	assert.equal(dryRun.ok, true, JSON.stringify(dryRun.errors));
	assert.equal((await service.save(docPath)).ok, true);
	let output = await JSZip.loadAsync(vault.store.get(docPath));
	assert.ok(output.file('word/comments.xml'), 'dry run must not mutate the cached session');

	const apply = await service.apply(docPath, [{ op: 'docx.removeComments' }]);
	assert.equal(apply.ok, true, JSON.stringify(apply.errors));
	assert.ok(apply.changed?.includes('word/comments.xml'));
	assert.deepEqual(apply.preview, [{ id: 'comments', field: 'removeAll', before: 1, after: 0 }]);
	assert.equal((await service.save(docPath)).ok, true);

	const savedBytes = vault.store.get(docPath);
	output = await JSZip.loadAsync(savedBytes);
	for (const partPath of [...COMMENT_PARTS, ...COMMENT_RELATIONSHIP_PARTS]) {
		assert.equal(output.file(partPath), null, `${partPath} must be removed`);
	}

	const contentTypes = await output.file('[Content_Types].xml').async('string');
	for (const partPath of COMMENT_PARTS) {
		assert.doesNotMatch(contentTypes, new RegExp(`PartName="/${partPath}"`));
	}
	const rels = await output.file('word/_rels/document.xml.rels').async('string');
	for (const relationshipType of COMMENT_RELATIONSHIP_TYPES) {
		assert.doesNotMatch(rels, new RegExp(relationshipType));
	}
	assert.match(rels, /Target="media\/kept\.png"/);
	assert.ok(output.file('word/media/kept.png'));
	assert.ok(output.file('word/media/comment-only.png'), 'comment-only media is intentionally preserved');

	for (const [partPath, expectedText] of [
		['word/document.xml', 'Body text'],
		['word/headers/header1.xml', 'Header text'],
		['word/footers/footer1.xml', 'Footer text'],
		['word/footnotes.xml', 'Footnote text'],
		['word/endnotes.xml', 'Endnote text'],
	]) {
		const xml = await output.file(partPath).async('string');
		assert.match(xml, new RegExp(expectedText));
		assert.doesNotMatch(xml, /<w:commentRange(?:Start|End)\b/);
		assert.doesNotMatch(xml, /<w:commentReference\b/);
	}
	const documentXml = await output.file('word/document.xml').async('string');
	assert.match(documentXml, /<w:ins\b/);
	assert.match(documentXml, /Tracked change stays/);

	const snapshot = await describeDocxFromBuffer(
		savedBytes.buffer.slice(savedBytes.byteOffset, savedBytes.byteOffset + savedBytes.byteLength),
		docPath,
	);
	assert.equal(snapshot.scope.review?.hasComments, false);
	assert.equal(snapshot.scope.review?.hasTrackChanges, true);
	assert.equal(snapshot.blocks.some((block) => block.kind === 'comment'), false);

	const repeat = await service.apply(docPath, [{ op: 'docx.removeComments' }]);
	assert.equal(repeat.ok, true, JSON.stringify(repeat.errors));
	assert.deepEqual(repeat.changed, []);
	assert.equal(repeat.preview, undefined);
});
