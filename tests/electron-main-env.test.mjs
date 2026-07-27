import assert from 'node:assert/strict';
import { test } from 'node:test';
import { electronMainEnv } from '../scripts/lib/text-offset-harness.mjs';

test('electronMainEnv strips only the Node-mode Electron flag', () => {
	const environment = electronMainEnv(
		{ EXPLICIT: 'override', ELECTRON_DISABLE_SECURITY_WARNINGS: 'true' },
		{
			ELECTRON_RUN_AS_NODE: '1',
			EXPLICIT: 'parent',
			PRESERVED: 'value',
		},
	);

	assert.equal(environment.ELECTRON_RUN_AS_NODE, undefined);
	assert.equal(environment.PRESERVED, 'value');
	assert.equal(environment.EXPLICIT, 'override');
	assert.equal(environment.ELECTRON_DISABLE_SECURITY_WARNINGS, 'true');
});
