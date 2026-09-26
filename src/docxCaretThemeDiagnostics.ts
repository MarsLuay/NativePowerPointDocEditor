export interface CaretThemeSnapshot {
	file?: string;
	wantDark: boolean;
	roots: number;
	docCaretVar: string | null;
	caretBackground: string | null;
	rootHasDarkClass: boolean;
}

export const CARET_THEME_LOG_MESSAGE = 'DOCX caret pinned to page ink';
export const CARET_THEME_SUPPRESSED_MESSAGE = 'DOCX caret theme diagnostics suppressed';

export interface CaretThemeDiagnosticEntry {
	message: string;
	data: Record<string, unknown>;
}

function caretThemeKey(snapshot: CaretThemeSnapshot): string {
	return [
		snapshot.wantDark ? '1' : '0',
		String(snapshot.roots),
		snapshot.docCaretVar ?? '',
		snapshot.caretBackground ?? '',
		snapshot.rootHasDarkClass ? '1' : '0',
	].join('\u0000');
}

/**
 * Records caret/theme telemetry once per distinct state. Identical follow-up
 * syncs stay silent until the next change, which includes one suppressed count.
 */
export function createCaretThemeDiagnostics(
	emit: (entry: CaretThemeDiagnosticEntry) => void,
): { record(snapshot: CaretThemeSnapshot): void } {
	let lastKey: string | null = null;
	let suppressed = 0;

	return {
		record(snapshot) {
			const key = caretThemeKey(snapshot);
			if (lastKey === key) {
				suppressed += 1;
				return;
			}

			if (suppressed > 0) {
				emit({
					message: CARET_THEME_SUPPRESSED_MESSAGE,
					data: {
						file: snapshot.file,
						suppressed,
					},
				});
				suppressed = 0;
			}

			lastKey = key;
			emit({
				message: CARET_THEME_LOG_MESSAGE,
				data: { ...snapshot },
			});
		},
	};
}
