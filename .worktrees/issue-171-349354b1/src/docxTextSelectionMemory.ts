export interface DocxTextSelectionRange {
	from: number;
	to: number;
}

export interface DocxTextSelectionEventContext {
	eventType: string;
	insideEditorPages: boolean;
	insideHiddenProseMirror: boolean;
}

export function shouldResetDocxPreservedTextSelection(
	context: DocxTextSelectionEventContext,
): boolean {
	return context.insideEditorPages
		|| (context.eventType === 'keydown' && context.insideHiddenProseMirror);
}

export function updateDocxPreservedTextSelection(
	previous: DocxTextSelectionRange | null,
	current: DocxTextSelectionRange | null,
	resetForEditorInput: boolean,
): DocxTextSelectionRange | null {
	if (resetForEditorInput) {
		return null;
	}
	if (current && current.to > current.from) {
		return current;
	}
	return previous;
}
