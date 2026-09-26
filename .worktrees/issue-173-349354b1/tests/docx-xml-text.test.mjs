import assert from 'node:assert/strict';
import test from 'node:test';
import { decodeDocxXmlEntities } from '../src/docxXmlText.ts';

test('decodeDocxXmlEntities decodes standard XML entities', () => {
	assert.equal(decodeDocxXmlEntities('&lt;'), '<');
	assert.equal(decodeDocxXmlEntities('&gt;'), '>');
	assert.equal(decodeDocxXmlEntities('&amp;'), '&');
	assert.equal(decodeDocxXmlEntities('&quot;'), '"');
	assert.equal(decodeDocxXmlEntities('&apos;'), "'");
});

test('decodeDocxXmlEntities decodes decimal code points', () => {
	assert.equal(decodeDocxXmlEntities('&#65;'), 'A');
	assert.equal(decodeDocxXmlEntities('&#32;'), ' ');
	assert.equal(decodeDocxXmlEntities('&#128512;'), '😀');
});

test('decodeDocxXmlEntities decodes hex code points', () => {
	assert.equal(decodeDocxXmlEntities('&#x41;'), 'A');
	assert.equal(decodeDocxXmlEntities('&#x20;'), ' ');
	assert.equal(decodeDocxXmlEntities('&#x1F600;'), '😀');
});

test('decodeDocxXmlEntities ignores invalid entities or leaves them unchanged', () => {
	assert.equal(decodeDocxXmlEntities('&unknown;'), '&unknown;');
	assert.equal(decodeDocxXmlEntities('&#;'), '&#;');
	assert.equal(decodeDocxXmlEntities('&#x;'), '&#x;');
});

test('decodeDocxXmlEntities handles mixed content', () => {
	assert.equal(
		decodeDocxXmlEntities('Hello &lt;World&gt;! &#65;&#x42;&amp;C'),
		'Hello <World>! AB&C'
	);
});

test('decodeDocxXmlEntities returns original string if no entities', () => {
	assert.equal(decodeDocxXmlEntities('Hello World!'), 'Hello World!');
});

test('decodeDocxXmlEntities handles invalid high code points gracefully', () => {
	assert.throws(() => decodeDocxXmlEntities('&#9999999999;'), RangeError);
	assert.throws(() => decodeDocxXmlEntities('&#xFFFFFFFFF;'), RangeError);
});
