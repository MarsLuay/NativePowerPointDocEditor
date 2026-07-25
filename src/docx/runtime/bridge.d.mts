import type {
	DocxCommand,
	DocxEditorComponent,
	TranslationFunction,
	Translations,
} from './contract';
import type { EditorView } from 'prosemirror-view';

export const DocxEditor: DocxEditorComponent;

export function clearParagraphMeasureCache(): void;

export function insertTable(rows: number, columns: number): DocxCommand;
export function setFontSize(size: number): DocxCommand;
export function setFontFamily(fontFamily: string): DocxCommand;
export function setLineSpacing(value: number): DocxCommand;
export function insertImageFromFile(
	view: EditorView,
	file: File,
	options: {
		maxWidth?: number;
		onError?: (error: Error) => void;
		onInserted?: (dimensions?: { width: number; height: number }) => void;
	},
): void;
export function isSuggestionModeActive(state: unknown): boolean;

export function loadFontFromBuffer(fontFamily: string, buffer: ArrayBuffer): Promise<boolean>;

export const en: Translations;
export function deepMerge(base: Translations, override: Translations | undefined): Translations;
export function createT(strings: Translations, language?: string): TranslationFunction;
export function loadDocxEditorLocale(locale: string): Promise<Translations | undefined>;
