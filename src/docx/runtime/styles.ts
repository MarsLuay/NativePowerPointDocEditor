import proseMirrorEditorStyles from '../../../vendor/docx-editor-runtime/core/dist/prosemirror/editor.css';
import editorStyles from '../../../vendor/docx-editor-runtime/react/dist/styles.css';

export const docxEditorRuntimeStyles = [
	proseMirrorEditorStyles,
	editorStyles,
].join('\n');
