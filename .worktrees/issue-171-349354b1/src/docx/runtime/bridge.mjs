export { DocxEditor } from /* vendored runtime alias */ '@npde/docx-editor-react';

export {
	clearParagraphMeasureCache,
} from /* vendored runtime alias */ '@npde/docx-editor-core/layout-bridge';

export {
	insertTable,
	setFontSize,
	setFontFamily,
	setLineSpacing,
	insertImageFromFile,
} from /* vendored runtime alias */ '@npde/docx-editor-core/prosemirror/commands';

export {
	textFormattingToMarks,
} from /* vendored runtime alias */ '@npde/docx-editor-core/prosemirror/commands/formatting';

export {
	loadFontFromBuffer,
} from /* vendored runtime alias */ '@npde/docx-editor-core/utils';

export {
	isSuggestionModeActive,
} from /* vendored runtime alias */ '@npde/docx-editor-core/prosemirror/plugins';

export {
	createT,
	deepMerge,
	en,
} from /* vendored runtime alias */ '@npde/docx-editor-i18n';

const localeLoaders = {
	en: () => import('@npde/docx-editor-i18n/en'),
	he: () => import('@npde/docx-editor-i18n/he'),
	pl: () => import('@npde/docx-editor-i18n/pl'),
	'pt-BR': () => import('@npde/docx-editor-i18n/pt-BR'),
	tr: () => import('@npde/docx-editor-i18n/tr'),
	'zh-CN': () => import('@npde/docx-editor-i18n/zh-CN'),
};

export async function loadDocxEditorLocale(locale) {
	const loader = localeLoaders[locale];
	if (!loader) {
		return undefined;
	}

	return (await loader()).default;
}
