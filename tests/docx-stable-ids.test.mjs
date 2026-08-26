import assert from 'node:assert/strict';
import { test } from 'node:test';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

async function loadDocxStableIdsModule() {
	const outfile = await bundleSource('src/ai/docxStableIds.ts', 'docx-stable-ids.cjs');
	return import(`file://${outfile}`);
}

test('parseStableLocation parses body locations correctly', async () => {
	const { parseStableLocation } = await loadDocxStableIdsModule();

	assert.deepEqual(parseStableLocation('body/p[0]'), {
		part: 'body',
		partNumber: null,
		kind: 'paragraph',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('body/p[12]/r[3]'), {
		part: 'body',
		partNumber: null,
		kind: 'run',
		paragraphIndex: 12,
		runIndex: 3,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('body/tbl[2]'), {
		part: 'body',
		partNumber: null,
		kind: 'table',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 2,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('body/tbl[1]/tr[4]/tc[5]'), {
		part: 'body',
		partNumber: null,
		kind: 'cell',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 1,
		rowIndex: 4,
		colIndex: 5,
	});
});

test('parseStableLocation parses header locations correctly', async () => {
	const { parseStableLocation } = await loadDocxStableIdsModule();

	assert.deepEqual(parseStableLocation('header/1/p[2]'), {
		part: 'header',
		partNumber: 1,
		kind: 'paragraph',
		paragraphIndex: 2,
		runIndex: null,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('header/2/p[0]/r[5]'), {
		part: 'header',
		partNumber: 2,
		kind: 'run',
		paragraphIndex: 0,
		runIndex: 5,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('header/3/tbl[0]'), {
		part: 'header',
		partNumber: 3,
		kind: 'table',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 0,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('header/1/tbl[2]/tr[3]/tc[4]'), {
		part: 'header',
		partNumber: 1,
		kind: 'cell',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 2,
		rowIndex: 3,
		colIndex: 4,
	});
});

test('parseStableLocation parses footer locations correctly', async () => {
	const { parseStableLocation } = await loadDocxStableIdsModule();

	assert.deepEqual(parseStableLocation('footer/1/p[0]'), {
		part: 'footer',
		partNumber: 1,
		kind: 'paragraph',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('footer/2/p[1]/r[2]'), {
		part: 'footer',
		partNumber: 2,
		kind: 'run',
		paragraphIndex: 1,
		runIndex: 2,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('footer/1/tbl[0]'), {
		part: 'footer',
		partNumber: 1,
		kind: 'table',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 0,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('footer/1/tbl[0]/tr[1]/tc[2]'), {
		part: 'footer',
		partNumber: 1,
		kind: 'cell',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 0,
		rowIndex: 1,
		colIndex: 2,
	});
});

test('parseStableLocation parses footnote and endnote locations correctly', async () => {
	const { parseStableLocation } = await loadDocxStableIdsModule();

	assert.deepEqual(parseStableLocation('footnotes/fn[0]/p[1]'), {
		part: 'footnotes',
		partNumber: 0,
		kind: 'paragraph',
		paragraphIndex: 1,
		runIndex: null,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('footnotes/fn[1]/p[2]/r[3]'), {
		part: 'footnotes',
		partNumber: 1,
		kind: 'run',
		paragraphIndex: 2,
		runIndex: 3,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('endnotes/en[0]/tbl[1]'), {
		part: 'endnotes',
		partNumber: 0,
		kind: 'table',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 1,
		rowIndex: null,
		colIndex: null,
	});

	assert.deepEqual(parseStableLocation('endnotes/en[2]/tbl[0]/tr[1]/tc[2]'), {
		part: 'endnotes',
		partNumber: 2,
		kind: 'cell',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 0,
		rowIndex: 1,
		colIndex: 2,
	});
});

test('parseStableLocation returns null for invalid IDs', async () => {
	const { parseStableLocation } = await loadDocxStableIdsModule();

	assert.equal(parseStableLocation(''), null);
	assert.equal(parseStableLocation('body'), null);
	assert.equal(parseStableLocation('body/'), null);
	assert.equal(parseStableLocation('body/p'), null);
	assert.equal(parseStableLocation('body/p[abc]'), null);
	assert.equal(parseStableLocation('header/p[0]'), null);
	assert.equal(parseStableLocation('footer/1'), null);
	assert.equal(parseStableLocation('footnotes/p[0]'), null);
	assert.equal(parseStableLocation('invalid/p[0]'), null);
});

test('location builders construct correct IDs', async () => {
	const {
		paragraphIdForLocation,
		runIdForLocation,
		tableIdForLocation,
		cellIdForLocation,
	} = await loadDocxStableIdsModule();

	const locBodyPara = {
		part: 'body',
		partNumber: null,
		kind: 'paragraph',
		paragraphIndex: 5,
		runIndex: null,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	};
	assert.equal(paragraphIdForLocation(locBodyPara), 'body/p[5]');

	const locHeaderRun = {
		part: 'header',
		partNumber: 2,
		kind: 'run',
		paragraphIndex: 1,
		runIndex: 3,
		tableIndex: null,
		rowIndex: null,
		colIndex: null,
	};
	assert.equal(runIdForLocation(locHeaderRun), 'header/2/p[1]/r[3]');

	const locFooterTable = {
		part: 'footer',
		partNumber: 1,
		kind: 'table',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 4,
		rowIndex: null,
		colIndex: null,
	};
	assert.equal(tableIdForLocation(locFooterTable), 'footer/1/tbl[4]');

	const locEndnoteCell = {
		part: 'endnotes',
		partNumber: 0,
		kind: 'cell',
		paragraphIndex: 0,
		runIndex: null,
		tableIndex: 1,
		rowIndex: 2,
		colIndex: 3,
	};
	assert.equal(cellIdForLocation(locEndnoteCell), 'endnotes/en[0]/tbl[1]/tr[2]/tc[3]');
});
