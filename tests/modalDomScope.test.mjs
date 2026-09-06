import assert from 'node:assert/strict';
import test from 'node:test';
import { loadModalDomScopeModule } from './helpers/load-plugin-modules.mjs';

test('modalDomScope', async (t) => {
	const { openModalDomScope, loadModalDomScope, closeModalDomScope } = await loadModalDomScopeModule();

	await t.test('openModalDomScope returns a Component instance', () => {
		const scope = openModalDomScope();
		assert.ok(scope, 'scope should be returned');
		assert.equal(typeof scope.load, 'function', 'scope should be a Component with load method');
		assert.equal(typeof scope.unload, 'function', 'scope should be a Component with unload method');
	});

	await t.test('loadModalDomScope calls load() on the component', () => {
		let loadCalled = false;
		const scope = {
			load() {
				loadCalled = true;
			}
		};
		loadModalDomScope(scope);
		assert.equal(loadCalled, true, 'load() should be called');
	});

	await t.test('closeModalDomScope calls unload() on the component', () => {
		let unloadCalled = false;
		const scope = {
			unload() {
				unloadCalled = true;
			}
		};
		closeModalDomScope(scope);
		assert.equal(unloadCalled, true, 'unload() should be called');
	});

	await t.test('closeModalDomScope gracefully handles undefined', () => {
		assert.doesNotThrow(() => {
			closeModalDomScope(undefined);
		});
	});
});
