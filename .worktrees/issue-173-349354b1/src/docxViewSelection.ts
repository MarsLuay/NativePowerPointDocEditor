export interface LoadedDocxView {
	getLoadedDocumentPath(): string | null;
}

export function selectDocxViewForPath<T extends LoadedDocxView>(
	views: readonly T[],
	activeView: T | null,
	path: string,
): T | null {
	if (activeView?.getLoadedDocumentPath() === path) {
		return activeView;
	}
	return views.find(view => view.getLoadedDocumentPath() === path) ?? null;
}
