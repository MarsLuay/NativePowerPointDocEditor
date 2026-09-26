import assert from 'node:assert/strict';
import test from 'node:test';
import { JSDOM } from 'jsdom';
import { loadInlineTextGeometryModule } from './helpers/load-plugin-modules.mjs';

function installContentBounds(window) {
	const layout = (element) => {
		if (element.style?.display === 'none' || element.classList?.contains('native-powerpoint-geometry-measure-suspended')) {
			return { left: 0, top: 0, width: 0, height: 0, empty: true };
		}
		const own = element.__box ?? null;
		let left = own ? own.left : Infinity;
		let top = own ? own.top : Infinity;
		let right = own ? own.left + own.width : -Infinity;
		let bottom = own ? own.top + own.height : -Infinity;
		for (const child of element.children ?? []) {
			const box = layout(child);
			if (box.empty) continue;
			left = Math.min(left, box.left);
			top = Math.min(top, box.top);
			right = Math.max(right, box.left + box.width);
			bottom = Math.max(bottom, box.top + box.height);
		}
		if (!Number.isFinite(left) || !Number.isFinite(top)) {
			return { left: 0, top: 0, width: 0, height: 0, empty: true };
		}
		return { left, top, width: right - left, height: bottom - top, empty: false };
	};

	for (const element of window.document.querySelectorAll('*')) {
		element.getBoundingClientRect = () => {
			const box = layout(element);
			return { left: box.left, top: box.top, width: box.width, height: box.height };
		};
	}
}

function box(element, rect) {
	element.__box = rect;
	return element;
}

test('text selection overlay does not inflate shape bounds and stays idempotent', async () => {
	const { InlineTextGeometry } = await loadInlineTextGeometryModule();
	const dom = new JSDOM('<!doctype html><div id="pane"></div>');
	const { document, window } = dom.window;
	globalThis.SVGElement ??= window.SVGElement;
	const pane = document.getElementById('pane');
	box(pane, { left: 0, top: 0, width: 800, height: 600 });
	pane.scrollLeft = 0;
	pane.scrollTop = 0;

	const shape = document.createElement('g');
	box(shape, { left: 227, top: 262, width: 216, height: 54 });
	const text = document.createElement('text');
	box(text, { left: 227, top: 262, width: 180, height: 40 });
	const caret = document.createElement('line');
	caret.className = 'native-powerpoint-svg-caret';
	box(caret, { left: 240, top: 262, width: 2, height: 40 });
	const selection = document.createElement('rect');
	selection.className = 'native-powerpoint-svg-selection';
	box(selection, { left: 11.31, top: 155.5, width: 431.69, height: 160.5 });
	shape.append(text, caret, selection);
	document.body.append(pane, shape);
	installContentBounds(window);

	const geometry = new InlineTextGeometry(() => pane);
	const first = geometry.getElementBox(shape);
	const second = geometry.getElementBox(shape);

	assert.deepEqual(first, { left: 227, top: 262, width: 216, height: 54 });
	assert.deepEqual(second, first);
	assert.equal(selection.classList.contains('native-powerpoint-geometry-measure-suspended'), false);
	assert.equal(caret.classList.contains('native-powerpoint-geometry-measure-suspended'), false);
});

test('issue 169 double-click selection boxes stay on the text-box frame', async () => {
	const { InlineTextGeometry } = await loadInlineTextGeometryModule();
	const dom = new JSDOM('<!doctype html><div id="pane"></div>');
	const { document, window } = dom.window;
	const pane = document.getElementById('pane');
	box(pane, { left: 0, top: 0, width: 800, height: 600 });
	pane.scrollLeft = 0;
	pane.scrollTop = 0;

	const frame = { left: 207, top: 202.5, width: 216, height: 54 };
	const shape = document.createElement('g');
	box(shape, frame);
	const text = document.createElement('text');
	box(text, { left: 207, top: 210, width: 180, height: 30 });
	const selection = document.createElement('rect');
	selection.className = 'native-powerpoint-svg-selection';
	selection.setAttribute('data-npde-geometry-overlay', 'true');
	shape.append(text, selection);
	document.body.append(pane, shape);
	installContentBounds(window);

	const geometry = new InlineTextGeometry(() => pane);
	const inflated = [
		{ left: -28.69000244140625, top: 36.5, width: 451.69000244140625, height: 220 },
		{ left: -70, top: 36.5, width: 493, height: 220 },
	];
	for (const outlier of inflated) {
		box(selection, outlier);
		assert.ok(outlier.width > frame.width);
		assert.deepEqual(geometry.getElementBox(shape), frame);
	}
	box(selection, { left: 0, top: 0, width: 0, height: 0 });
	assert.deepEqual(geometry.getElementBox(shape), frame);
});

test('grouped content and image geometry still union without the selection rect', async () => {
	const { InlineTextGeometry } = await loadInlineTextGeometryModule();
	const dom = new JSDOM('<!doctype html><div id="pane"></div>');
	const { document, window } = dom.window;
	const pane = document.getElementById('pane');
	box(pane, { left: 0, top: 0, width: 400, height: 400 });
	pane.scrollLeft = 0;
	pane.scrollTop = 0;

	const group = document.createElement('g');
	const image = document.createElement('image');
	box(image, { left: 10, top: 10, width: 40, height: 30 });
	const sibling = document.createElement('text');
	box(sibling, { left: 80, top: 12, width: 50, height: 20 });
	const selection = document.createElement('rect');
	selection.setAttribute('data-npde-geometry-overlay', 'true');
	box(selection, { left: 0, top: 0, width: 400, height: 400 });
	group.append(image, sibling, selection);
	document.body.append(pane, group);
	installContentBounds(window);

	const geometry = new InlineTextGeometry(() => pane);
	assert.deepEqual(geometry.getElementBox(group), { left: 10, top: 10, width: 120, height: 30 });
	assert.deepEqual(geometry.getElementBox(image), { left: 10, top: 10, width: 40, height: 30 });
});
