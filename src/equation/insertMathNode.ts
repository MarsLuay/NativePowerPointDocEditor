import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

import type { CompiledEquation } from './compileEquation';

export function insertMathNodeAtSelection(view: EditorView, equation: CompiledEquation): boolean {
	const mathType = view.state.schema.nodes.math;
	if (!mathType) {
		return false;
	}

	const node = mathType.create({
		display: equation.display,
		ommlXml: equation.ommlXml,
		plainText: equation.plainText,
	});
	const { $from, from, to } = view.state.selection;
	if (from === to && !$from.parent.canReplaceWith($from.index(), $from.indexAfter(), mathType)) {
		return false;
	}

	let transaction = view.state.tr.replaceWith(from, to, node);
	const after = transaction.mapping.map(from) + node.nodeSize;
	if (after <= transaction.doc.content.size) {
		transaction = transaction.setSelection(TextSelection.near(transaction.doc.resolve(after)));
	}
	view.dispatch(transaction.scrollIntoView());
	view.focus();
	return true;
}
