import { describe, expect, it } from 'bun:test';
import { synchronizeSelectionOverlayAfterTransaction } from './selectionOverlayRenderGate';

describe('synchronizeSelectionOverlayAfterTransaction', () => {
	it('clears the old caret and waits for the matching layout after a document edit', () => {
		const events: string[] = [];
		const state = { selection: { from: 8, to: 8 } };

		synchronizeSelectionOverlayAfterTransaction(true, state, {
			requestRender: () => events.push('request-render'),
			clearSelectionOverlay: () => events.push('clear-overlay'),
			updateSelectionOverlay: () => events.push('update-overlay'),
		});

		expect(events).toEqual(['request-render', 'clear-overlay']);
	});

	it('updates the overlay immediately for a selection-only transaction', () => {
		const events: string[] = [];
		const state = { selection: { from: 8, to: 8 } };

		synchronizeSelectionOverlayAfterTransaction(false, state, {
			requestRender: () => events.push('request-render'),
			clearSelectionOverlay: () => events.push('clear-overlay'),
			updateSelectionOverlay: () => events.push('update-overlay'),
		});

		expect(events).toEqual(['request-render', 'update-overlay']);
	});
});
