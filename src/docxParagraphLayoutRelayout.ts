import type { Node as ProseMirrorNode } from 'prosemirror-model';

const TYPOGRAPHY_MARK_TYPES = ['bold', 'italic', 'fontSize', 'fontFamily', 'textColor', 'underline', 'strike'] as const;

export function stableParagraphLayoutValue(value: unknown): string {
	if (value === null || value === undefined) {
		return '';
	}

	if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
		return String(value);
	}

	try {
		return JSON.stringify(value) ?? '';
	} catch {
		return '';
	}
}

export function getParagraphListLayoutSignature(node: ProseMirrorNode): string {
	const attrs = node.attrs as Record<string, unknown>;
	const hasListLayout = attrs.numPr != null
		|| attrs.listMarker != null
		|| attrs.listMarkerHidden != null
		|| attrs.listMarkerFontFamily != null
		|| attrs.listMarkerFontSize != null;

	if (!hasListLayout) {
		return '';
	}

	return [
		stableParagraphLayoutValue(attrs.numPr),
		stableParagraphLayoutValue(attrs.listMarker),
		stableParagraphLayoutValue(attrs.listMarkerHidden),
		stableParagraphLayoutValue(attrs.listMarkerFontFamily),
		stableParagraphLayoutValue(attrs.listMarkerFontSize),
		stableParagraphLayoutValue(attrs.indentLeft),
		stableParagraphLayoutValue(attrs.indentFirstLine),
		stableParagraphLayoutValue(attrs.hangingIndent),
	].join('\u001f');
}

function getParagraphRunTypographySignature(node: ProseMirrorNode): string {
	const parts: string[] = [];

	node.descendants((child) => {
		if (!child.isText) {
			return true;
		}

		for (const mark of child.marks) {
			if ((TYPOGRAPHY_MARK_TYPES as readonly string[]).includes(mark.type.name)) {
				parts.push(`${mark.type.name}:${stableParagraphLayoutValue(mark.attrs)}`);
			}
		}

		return true;
	});

	return parts.join(';');
}

export function getParagraphTypographySignature(node: ProseMirrorNode): string {
	const attrs = node.attrs as Record<string, unknown>;

	return [
		stableParagraphLayoutValue(attrs.styleId),
		stableParagraphLayoutValue(attrs.alignment),
		stableParagraphLayoutValue(attrs.lineSpacing),
		stableParagraphLayoutValue(attrs.lineSpacingRule),
		stableParagraphLayoutValue(attrs.spaceBefore),
		stableParagraphLayoutValue(attrs.spaceAfter),
		stableParagraphLayoutValue(attrs.indentLeft),
		stableParagraphLayoutValue(attrs.indentRight),
		stableParagraphLayoutValue(attrs.indentFirstLine),
		stableParagraphLayoutValue(attrs.hangingIndent),
		stableParagraphLayoutValue(attrs.outlineLevel),
		stableParagraphLayoutValue(attrs.defaultFontSize),
		stableParagraphLayoutValue(attrs.defaultFontFamily),
		getParagraphRunTypographySignature(node),
	].join('\u001f');
}

export function getDocumentParagraphLayoutSignatures(doc: ProseMirrorNode): string[] {
	const signatures: string[] = [];

	doc.descendants((node) => {
		if (node.type.name !== 'paragraph') {
			return true;
		}

		signatures.push([
			getParagraphListLayoutSignature(node),
			getParagraphTypographySignature(node),
		].join('\u001e'));
		return false;
	});

	return signatures;
}

export function didListLayoutChange(before: ProseMirrorNode, after: ProseMirrorNode): boolean {
	const getListSignatures = (doc: ProseMirrorNode) => {
		const signatures: string[] = [];
		doc.descendants((node) => {
			if (node.type.name !== 'paragraph') {
				return true;
			}

			signatures.push(getParagraphListLayoutSignature(node));
			return false;
		});
		return signatures;
	};

	const beforeListSignatures = getListSignatures(before);
	const afterListSignatures = getListSignatures(after);
	const signatureCount = Math.max(beforeListSignatures.length, afterListSignatures.length);

	for (let index = 0; index < signatureCount; index += 1) {
		if ((beforeListSignatures[index] ?? '') !== (afterListSignatures[index] ?? '')) {
			return true;
		}
	}

	return false;
}

export function didParagraphTypographyChange(before: ProseMirrorNode, after: ProseMirrorNode): boolean {
	const getTypographySignatures = (doc: ProseMirrorNode) => {
		const signatures: string[] = [];
		doc.descendants((node) => {
			if (node.type.name !== 'paragraph') {
				return true;
			}

			signatures.push(getParagraphTypographySignature(node));
			return false;
		});
		return signatures;
	};

	const beforeTypographySignatures = getTypographySignatures(before);
	const afterTypographySignatures = getTypographySignatures(after);
	const signatureCount = Math.max(beforeTypographySignatures.length, afterTypographySignatures.length);

	for (let index = 0; index < signatureCount; index += 1) {
		if ((beforeTypographySignatures[index] ?? '') !== (afterTypographySignatures[index] ?? '')) {
			return true;
		}
	}

	return false;
}

export function didParagraphLayoutChange(before: ProseMirrorNode, after: ProseMirrorNode): boolean {
	return didListLayoutChange(before, after) || didParagraphTypographyChange(before, after);
}
