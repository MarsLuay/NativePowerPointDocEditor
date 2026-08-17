import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { test } from 'node:test';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
let modulePromise;

function loadSelectionModule() {
	modulePromise ??= bundleSource(
		'src/docxViewSelection.ts',
		'docx-view-selection.cjs',
	).then(outfile => require(outfile));
	return modulePromise;
}

function view(path, id) {
	return {
		id,
		getLoadedDocumentPath: () => path,
	};
}

test('selectDocxViewForPath prefers the active matching duplicate view', async () => {
	const { selectDocxViewForPath } = await loadSelectionModule();
	const hidden = view('Applications/resume.docx', 'hidden');
	const active = view('Applications/resume.docx', 'active');

	assert.equal(
		selectDocxViewForPath([hidden, active], active, 'Applications/resume.docx'),
		active,
	);
});

test('selectDocxViewForPath falls back to the first matching view', async () => {
	const { selectDocxViewForPath } = await loadSelectionModule();
	const firstMatch = view('Applications/resume.docx', 'first');
	const secondMatch = view('Applications/resume.docx', 'second');
	const otherActive = view('Applications/other.docx', 'active');

	assert.equal(
		selectDocxViewForPath(
			[firstMatch, secondMatch],
			otherActive,
			'Applications/resume.docx',
		),
		firstMatch,
	);
	assert.equal(
		selectDocxViewForPath([otherActive], otherActive, 'Applications/resume.docx'),
		null,
	);
});
