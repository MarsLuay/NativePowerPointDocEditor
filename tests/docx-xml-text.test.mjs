import assert from 'node:assert/strict';
import test from 'node:test';
import { loadDocxXmlTextModule } from './helpers/load-plugin-modules.mjs';

test('decodeDocxXmlEntities decodes standard XML entities', async () => {
    const { decodeDocxXmlEntities } = await loadDocxXmlTextModule();
	assert.equal(decodeDocxXmlEntities('&lt;'), '<');
	assert.equal(decodeDocxXmlEntities('&gt;'), '>');
	assert.equal(decodeDocxXmlEntities('&amp;'), '&');
	assert.equal(decodeDocxXmlEntities('&quot;'), '"');
	assert.equal(decodeDocxXmlEntities('&apos;'), "'");
});

test('decodeDocxXmlEntities decodes decimal code points', async () => {
    const { decodeDocxXmlEntities } = await loadDocxXmlTextModule();
	assert.equal(decodeDocxXmlEntities('&#65;'), 'A');
	assert.equal(decodeDocxXmlEntities('&#32;'), ' ');
	assert.equal(decodeDocxXmlEntities('&#128512;'), '😀');
});

test('decodeDocxXmlEntities decodes hex code points', async () => {
    const { decodeDocxXmlEntities } = await loadDocxXmlTextModule();
	assert.equal(decodeDocxXmlEntities('&#x41;'), 'A');
	assert.equal(decodeDocxXmlEntities('&#x20;'), ' ');
	assert.equal(decodeDocxXmlEntities('&#x1F600;'), '😀');
});

test('decodeDocxXmlEntities ignores invalid entities or leaves them unchanged', async () => {
    const { decodeDocxXmlEntities } = await loadDocxXmlTextModule();
	assert.equal(decodeDocxXmlEntities('&unknown;'), '&unknown;');
	assert.equal(decodeDocxXmlEntities('&#;'), '&#;');
	assert.equal(decodeDocxXmlEntities('&#x;'), '&#x;');
});

test('decodeDocxXmlEntities handles mixed content', async () => {
    const { decodeDocxXmlEntities } = await loadDocxXmlTextModule();
	assert.equal(
		decodeDocxXmlEntities('Hello &lt;World&gt;! &#65;&#x42;&amp;C'),
		'Hello <World>! AB&C'
	);
});

test('decodeDocxXmlEntities returns original string if no entities', async () => {
    const { decodeDocxXmlEntities } = await loadDocxXmlTextModule();
	assert.equal(decodeDocxXmlEntities('Hello World!'), 'Hello World!');
});

test('decodeDocxXmlEntities handles invalid high code points gracefully', async () => {
    const { decodeDocxXmlEntities } = await loadDocxXmlTextModule();
	assert.throws(() => decodeDocxXmlEntities('&#9999999999;'), RangeError);
	assert.throws(() => decodeDocxXmlEntities('&#xFFFFFFFFF;'), RangeError);
});

test('normalizeDocxExtractedText normalizes line endings and spaces', async () => {
    const { normalizeDocxExtractedText } = await loadDocxXmlTextModule();
    assert.equal(normalizeDocxExtractedText('A\rB'), 'A\nB');
    assert.equal(normalizeDocxExtractedText('A \t\nB'), 'A\nB');
    assert.equal(normalizeDocxExtractedText('A\n\t B'), 'A\nB');
});

test('normalizeDocxExtractedText collapses consecutive newlines to maximum of two', async () => {
    const { normalizeDocxExtractedText } = await loadDocxXmlTextModule();
    assert.equal(normalizeDocxExtractedText('A\n\n\n\nB'), 'A\n\nB');
    assert.equal(normalizeDocxExtractedText('A\n\nB'), 'A\n\nB');
});

test('normalizeDocxExtractedText trims surrounding whitespace', async () => {
    const { normalizeDocxExtractedText } = await loadDocxXmlTextModule();
    assert.equal(normalizeDocxExtractedText(' \n  Hello World \n '), 'Hello World');
});

test('extractDocxTextFromXml extracts plain text from simple paragraph', async () => {
    const { extractDocxTextFromXml } = await loadDocxXmlTextModule();
    const xml = '<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>';
    assert.equal(extractDocxTextFromXml(xml), 'Hello World');
});

test('extractDocxTextFromXml handles tabs and breaks across runs', async () => {
    const { extractDocxTextFromXml } = await loadDocxXmlTextModule();
    const xml = '<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r></w:p>';
    assert.equal(extractDocxTextFromXml(xml), 'A\tB\nC');
});

test('extractDocxTextFromXml handles decoded entities', async () => {
    const { extractDocxTextFromXml } = await loadDocxXmlTextModule();
    const xml = '<w:p><w:r><w:t>A &amp; B</w:t></w:r></w:p>';
    assert.equal(extractDocxTextFromXml(xml), 'A & B');
});

test('extractDocxTextFromXml separates paragraphs with newlines', async () => {
    const { extractDocxTextFromXml } = await loadDocxXmlTextModule();
    const xml = '<w:p><w:r><w:t>P1</w:t></w:r></w:p><w:p><w:r><w:t>P2</w:t></w:r></w:p>';
    assert.equal(extractDocxTextFromXml(xml), 'P1\nP2');
});

test('extractDocxRunText extracts text with tabs and breaks without normalizing', async () => {
    const { extractDocxRunText } = await loadDocxXmlTextModule();
    const xml = '<w:r><w:t>A </w:t><w:br/><w:t> B</w:t><w:tab/><w:t>C</w:t></w:r>';
    assert.equal(extractDocxRunText(xml), 'A \n B\tC');
});

test('extractDocxRunText parses partial tags', async () => {
    const { extractDocxRunText } = await loadDocxXmlTextModule();
    const xml = '<w:t xml:space="preserve"> Hello </w:t>';
    assert.equal(extractDocxRunText(xml), ' Hello ');
});
