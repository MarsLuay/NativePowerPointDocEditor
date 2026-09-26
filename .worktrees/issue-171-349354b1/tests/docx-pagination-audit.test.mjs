import assert from 'node:assert/strict';
import test from 'node:test';
import {
	classifyPaginationDelta,
	formatPaginationAuditStatus,
} from '../scripts/lib/docx-pagination-audit.mjs';

test('DOCX pagination audit classifies only extra preview pages as over-pagination', () => {
	assert.equal(classifyPaginationDelta(1), 'over-pagination');
	assert.equal(formatPaginationAuditStatus(classifyPaginationDelta(1)), 'OVER');
});

test('DOCX pagination audit treats fewer preview pages as a reference-renderer difference', () => {
	assert.equal(classifyPaginationDelta(-1), 'reference-difference');
	assert.equal(formatPaginationAuditStatus(classifyPaginationDelta(-1)), 'REFERENCE-DIFF');
	assert.equal(classifyPaginationDelta(0), 'match');
	assert.equal(formatPaginationAuditStatus(classifyPaginationDelta(0)), 'PASS');
});
