/**
 * Keep painted selection geometry tied to the DOM that owns it.
 *
 * A document transaction invalidates the current page layout. The caller has
 * already requested a render through LayoutSelectionGate, so retaining or
 * recomputing the caret from stale geometry would display it at the wrong
 * character until that render finishes.
 */
export function synchronizeSelectionOverlayAfterTransaction<T>(
	documentChanged: boolean,
	state: T,
	controls: {
		requestRender: () => void;
		clearSelectionOverlay: () => void;
		updateSelectionOverlay: (state: T) => void;
	}
): void {
	controls.requestRender();

	if (documentChanged) {
		controls.clearSelectionOverlay();
		return;
	}

	controls.updateSelectionOverlay(state);
}
