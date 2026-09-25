import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);
let cached;
let cachedResolver;

async function loadModule() {
	if (cached) return cached;
	const directory = await mkdtemp(path.join(tmpdir(), 'npde-docx-anchors-'));
	const outfile = path.join(directory, 'docx-ooxml.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/docxOoxml.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cached = require(outfile);
	return cached;
}

async function loadResolver() {
	if (cachedResolver) return cachedResolver;
	const directory = await mkdtemp(path.join(tmpdir(), 'npde-docx-anchor-resolver-'));
	const outfile = path.join(directory, 'docx-block-resolver.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/docxBlockResolver.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cachedResolver = require(outfile);
	return cachedResolver;
}

test('DOCX paragraph anchors round-trip and are generated uniquely for editable paragraphs', async () => {
	const { ensureParagraphAnchors, parseParagraph, getParagraphAnchor } = await loadModule();
	const source = '<w:document xmlns:w="urn:w"><w:body><w:p w14:paraId="A1B2C3D4"><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body></w:document>';
	const anchored = ensureParagraphAnchors(source);
	const anchors = [...anchored.matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map((match) => match[1]);
	assert.equal(anchors.length, 2);
	assert.equal(new Set(anchors).size, anchors.length);
	assert.equal(anchors[0], 'A1B2C3D4');
	assert.match(anchored, /xmlns:w14="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordml"/);
	assert.equal(getParagraphAnchor(anchored.match(/<w:p\b[^>]*>/g)[1]), anchors[1]);
	assert.equal(parseParagraph(anchored.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)[0]).anchor, anchors[0]);
});

test('persistent anchors resolve without positional retargeting across DOCX parts', async () => {
	const { findParagraphByAnchor } = await loadResolver();
	const bodyInner = '<w:p w14:paraId="11111111"><w:r><w:t>first</w:t></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:t>second</w:t></w:r></w:p>';
	assert.equal(findParagraphByAnchor(bodyInner, '22222222').id, 'body/p[1]');
	assert.equal(findParagraphByAnchor(bodyInner, '11111111', 'header/2').id, 'header/2/p[0]');
	assert.throws(() => findParagraphByAnchor(bodyInner, '99999999'), /was not found/);
});
