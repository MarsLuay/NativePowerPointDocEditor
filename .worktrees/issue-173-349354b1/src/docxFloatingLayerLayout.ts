export interface FloatingLayerRect {
	left: number;
	right: number;
	top: number;
	bottom: number;
	width: number;
	height: number;
}

export interface FloatingLayerViewport {
	width: number;
	height: number;
}

export interface FloatingLayerPosition {
	left: number;
	top: number;
}

export interface FormattingDropdownScrollTransition {
	savedScrollLeft: number | null;
	visualOffset: number;
	restoreScrollLeft: number | null;
}

export function getFormattingDropdownScrollTransition(
	open: boolean,
	currentScrollLeft: number,
	savedScrollLeft: number | null,
): FormattingDropdownScrollTransition {
	if (open) {
		const preservedScrollLeft = savedScrollLeft ?? currentScrollLeft;
		return {
			savedScrollLeft: preservedScrollLeft,
			visualOffset: preservedScrollLeft,
			restoreScrollLeft: null,
		};
	}

	return {
		savedScrollLeft: null,
		visualOffset: 0,
		restoreScrollLeft: savedScrollLeft,
	};
}

export function isFullscreenDialogLayer(
	role: string | null,
	ariaModal: string | null,
	hasDirectDialogChild: boolean,
): boolean {
	return (role === 'dialog' && ariaModal === 'true') || hasDirectDialogChild;
}

export function calculateClampedFloatingLayerPosition(
	rect: FloatingLayerRect,
	viewport: FloatingLayerViewport,
	position: FloatingLayerPosition,
	margin = 8,
): FloatingLayerPosition {
	const availableWidth = Math.max(0, viewport.width - margin * 2);
	const availableHeight = Math.max(0, viewport.height - margin * 2);
	let deltaX = 0;
	let deltaY = 0;

	if (rect.width > availableWidth || rect.left < margin) {
		deltaX = margin - rect.left;
	} else if (rect.right > viewport.width - margin) {
		deltaX = viewport.width - margin - rect.right;
	}

	if (rect.height > availableHeight || rect.top < margin) {
		deltaY = margin - rect.top;
	} else if (rect.bottom > viewport.height - margin) {
		deltaY = viewport.height - margin - rect.bottom;
	}

	return {
		left: position.left + deltaX,
		top: position.top + deltaY,
	};
}
