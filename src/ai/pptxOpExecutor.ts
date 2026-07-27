import type { Vault } from 'obsidian';
import type { ChartDataUpdate } from '../ChartData';
import { getImageMimeType } from '../PowerPointInsertModals';
import type {
	ImageCrop,
	InsertableShapeGeometry,
	PresentationEngine,
	RunStyleChange,
	SlideLayoutKind,
} from '../PresentationEngine';
import type { ShapeTransform } from 'pptx-svg';
import type { ParagraphListStyle } from '../SlideInsertions';
import type { DrawingParagraphText } from '../powerpoint/drawingmlText';
import { isEditableShapeIndex } from '../powerpoint/svgUtils';
import { readRasterImageDimensions } from '../powerpoint/imageDimensions';
import { fitImageWithinBounds } from '../powerpoint/imageFit';
import { AI_ERROR_CODES, createAiError } from './errors';
import { describePptxFromEngine } from './pptxDescribe';
import { pptxShapeId } from './pptxIds';
import type { ApplyPreviewChange, DocumentOp } from './types';
import { readVaultBinaryFile } from './vaultBinary';

export interface PptxOpExecutionResult {
	changedIds: string[];
	preview: ApplyPreviewChange[];
	affectedSlideIndices: Set<number>;
	warnings: string[];
}

export interface PptxOpExecutionContext {
	engine: PresentationEngine;
	vault: Vault;
	filePath: string;
	dryRun: boolean;
}

const INSERTABLE_GEOMETRIES = new Set<InsertableShapeGeometry>([
	'rect',
	'ellipse',
	'roundRect',
	'line',
	'rightArrow',
	'leftArrow',
	'upArrow',
	'downArrow',
]);

function asRecord(op: DocumentOp): Record<string, unknown> {
	return op;
}

function requireNumber(value: unknown, field: string): number {
	if (typeof value !== 'number' || !Number.isFinite(value)) {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be a number.`, { field });
	}
	return value;
}

function requireString(value: unknown, field: string): string {
	if (typeof value !== 'string') {
		throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `${field} must be a string.`, { field });
	}
	return value;
}

function requireReplacementParagraphs(value: unknown): DrawingParagraphText[] {
	if (!Array.isArray(value) || value.length === 0) {
		throw createAiError(
			AI_ERROR_CODES.SCHEMA_INVALID,
			'paragraphs must be a non-empty array of native PowerPoint paragraphs.',
			{ field: 'paragraphs' },
		);
	}

	return value.map((item, index) => {
		if (!item || typeof item !== 'object' || Array.isArray(item)) {
			throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Each paragraph must be an object.', {
				field: `paragraphs[${index}]`,
			});
		}
		const paragraph = item as Record<string, unknown>;
		const text = requireString(paragraph.text, `paragraphs[${index}].text`);
		if (/[\r\n]/.test(text)) {
			throw createAiError(
				AI_ERROR_CODES.SCHEMA_INVALID,
				'Paragraph text cannot contain line breaks; use one paragraphs[] entry per PowerPoint paragraph.',
				{ field: `paragraphs[${index}].text` },
			);
		}
		const listStyle = requireString(paragraph.listStyle, `paragraphs[${index}].listStyle`);
		if (listStyle !== 'none' && listStyle !== 'bullet' && listStyle !== 'number') {
			throw createAiError(
				AI_ERROR_CODES.SCHEMA_INVALID,
				'listStyle must be none, bullet, or number.',
				{ field: `paragraphs[${index}].listStyle` },
			);
		}
		const bold = paragraph.bold;
		if (bold !== undefined && typeof bold !== 'boolean') {
			throw createAiError(
				AI_ERROR_CODES.SCHEMA_INVALID,
				'bold must be a boolean when provided.',
				{ field: `paragraphs[${index}].bold` },
			);
		}
		return bold === undefined ? { text, listStyle } : { text, listStyle, bold };
	});
}

function requireTransform(value: unknown): ShapeTransform {
	const record = value as Record<string, unknown>;
	return {
		x: requireNumber(record.x, 'transform.x'),
		y: requireNumber(record.y, 'transform.y'),
		cx: requireNumber(record.cx, 'transform.cx'),
		cy: requireNumber(record.cy, 'transform.cy'),
		rot: requireNumber(record.rot, 'transform.rot'),
	};
}

function assertEditableShape(slideIndex: number, shapeIndex: number): void {
	if (!isEditableShapeIndex(shapeIndex)) {
		throw createAiError(
			AI_ERROR_CODES.OBJECT_NOT_EDITABLE,
			`Shape ${pptxShapeId(slideIndex, shapeIndex)} is not editable.`,
		);
	}
}

function assertSlideInRange(engine: PresentationEngine, slideIndex: number): void {
	if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= engine.slideCount) {
		throw createAiError(
			AI_ERROR_CODES.SLIDE_NOT_FOUND,
			`Slide index ${slideIndex} is out of range (0..${Math.max(engine.slideCount - 1, 0)}).`,
		);
	}
}

function findShapeSnapshot(
	engine: PresentationEngine,
	filePath: string,
	slideIndex: number,
	shapeIndex: number,
) {
	const snapshot = describePptxFromEngine(engine, filePath);
	const slide = snapshot.slides.find((entry) => entry.index === slideIndex);
	const shape = slide?.shapes.find((entry) => entry.index === shapeIndex);
	if (!shape) {
		throw createAiError(
			AI_ERROR_CODES.SHAPE_NOT_FOUND,
			`Shape ${pptxShapeId(slideIndex, shapeIndex)} was not found.`,
		);
	}
	return shape;
}

function tryFindShapeSnapshot(
	engine: PresentationEngine,
	filePath: string,
	slideIndex: number,
	shapeIndex: number,
) {
	try {
		return findShapeSnapshot(engine, filePath, slideIndex, shapeIndex);
	} catch {
		return null;
	}
}

async function applyTransformToInsertedShape(
	engine: PresentationEngine,
	slideIndex: number,
	shapeIndex: number,
	transform: ShapeTransform,
): Promise<void> {
	await engine.applyInsertedShapeTransform(slideIndex, shapeIndex, transform);
}

function parseImageFit<TFit extends string>(
	payload: Record<string, unknown>,
	defaultFit: TFit,
	allowedFits: readonly TFit[],
): TFit {
	const fit = payload.fit;
	if (fit === undefined) return defaultFit;
	if (typeof fit === 'string' && allowedFits.includes(fit as TFit)) return fit as TFit;
	throw createAiError(
		AI_ERROR_CODES.SCHEMA_INVALID,
		`fit must be ${allowedFits.map((value) => `"${value}"`).join(' or ')}.`,
		{ field: 'fit' },
	);
}

function fitInsertedImageTransform(
	transform: ShapeTransform,
	imageBytes: Uint8Array,
): ShapeTransform {
	const fittedSize = fitImageWithinBounds(
		readRasterImageDimensions(imageBytes),
		transform.cx,
		transform.cy,
	);
	return {
		...transform,
		x: transform.x + (transform.cx - fittedSize.width) / 2,
		y: transform.y + (transform.cy - fittedSize.height) / 2,
		cx: fittedSize.width,
		cy: fittedSize.height,
	};
}

export async function executePptxOp(
	context: PptxOpExecutionContext,
	op: DocumentOp,
): Promise<PptxOpExecutionResult> {
	const payload = asRecord(op);
	const opId = requireString(op.op, 'op');
	const result: PptxOpExecutionResult = {
		changedIds: [],
		preview: [],
		affectedSlideIndices: new Set<number>(),
		warnings: [],
	};

	switch (opId) {
		case 'pptx.updateShapeText': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const text = requireString(payload.text, 'text');
			if (/[\r\n]/.test(text)) {
				throw createAiError(
					AI_ERROR_CODES.SCHEMA_INVALID,
					'pptx.updateShapeText is single-paragraph. Use pptx.replaceShapeParagraphs for multiple paragraphs or lists.',
					{ op: opId, field: 'text' },
				);
			}
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const before = tryFindShapeSnapshot(context.engine, context.filePath, slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			result.preview.push({ id: changedId, field: 'text', before: before?.text ?? null, after: text });
			if (!context.dryRun) {
				await context.engine.updateShapeText(slideIndex, shapeIndex, text);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.replaceShapeParagraphs': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const paragraphs = requireReplacementParagraphs(payload.paragraphs);
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const before = tryFindShapeSnapshot(context.engine, context.filePath, slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			result.preview.push({
				id: changedId,
				field: 'paragraphs',
				before: before?.paragraphs ?? null,
				after: paragraphs,
			});
			if (!context.dryRun) {
				await context.engine.replaceShapeParagraphs(slideIndex, shapeIndex, paragraphs);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.deleteShape': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const before = findShapeSnapshot(context.engine, context.filePath, slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			result.preview.push({ id: changedId, field: 'shape', before, after: null });
			if (!context.dryRun) {
				await context.engine.deleteShape(slideIndex, shapeIndex);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.deleteShapes': {
			// Internal batch op produced by coalescePptxOps — not a public catalog id.
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			assertSlideInRange(context.engine, slideIndex);
			if (!Array.isArray(payload.shapeIndexes) || payload.shapeIndexes.length === 0) {
				throw createAiError(
					AI_ERROR_CODES.SCHEMA_INVALID,
					'shapeIndexes must be a non-empty array of numbers.',
					{ field: 'shapeIndexes' },
				);
			}
			const shapeIndexes = payload.shapeIndexes.map((value, index) => {
				if (typeof value !== 'number' || !Number.isFinite(value)) {
					throw createAiError(
						AI_ERROR_CODES.SCHEMA_INVALID,
						`shapeIndexes[${index}] must be a number.`,
						{ field: `shapeIndexes[${index}]` },
					);
				}
				assertEditableShape(slideIndex, value);
				return value;
			});
			for (const shapeIndex of shapeIndexes) {
				const before = findShapeSnapshot(context.engine, context.filePath, slideIndex, shapeIndex);
				const changedId = pptxShapeId(slideIndex, shapeIndex);
				result.preview.push({ id: changedId, field: 'shape', before, after: null });
				result.changedIds.push(changedId);
			}
			if (!context.dryRun) {
				await context.engine.deleteShapes(slideIndex, shapeIndexes);
			}
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.updateParagraphText': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const paragraphIndex = requireNumber(payload.paragraphIndex, 'paragraphIndex');
			const text = requireString(payload.text, 'text');
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const before = context.engine.getParagraphRunText(slideIndex, shapeIndex, paragraphIndex);
			if (before === null) {
				throw createAiError(AI_ERROR_CODES.SHAPE_NOT_FOUND, 'Paragraph not found on shape.', { op: opId });
			}
			const changedId = `${pptxShapeId(slideIndex, shapeIndex)}/p:${paragraphIndex}`;
			result.preview.push({ id: changedId, field: 'text', before, after: text });
			if (!context.dryRun) {
				await context.engine.updateParagraphText(slideIndex, shapeIndex, paragraphIndex, text);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.updateTextRun': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const paragraphIndex = requireNumber(payload.paragraphIndex, 'paragraphIndex');
			const runIndex = requireNumber(payload.runIndex, 'runIndex');
			const text = requireString(payload.text, 'text');
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = `${pptxShapeId(slideIndex, shapeIndex)}/p:${paragraphIndex}/r:${runIndex}`;
			result.preview.push({ id: changedId, field: 'text', before: null, after: text });
			if (!context.dryRun) {
				await context.engine.updateTextRun(slideIndex, shapeIndex, paragraphIndex, runIndex, text);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.replaceText': {
			const query = requireString(payload.query, 'query');
			const replacement = requireString(payload.replacement, 'replacement');
			const matchCase = payload.matchCase === true;
			const slideIndex = payload.slideIndex;
			const shapeIndex = payload.shapeIndex;
			const options: { matchCase?: boolean; slideIndex?: number; shapeIndex?: number } = { matchCase };
			if (typeof slideIndex === 'number' && typeof shapeIndex === 'number') {
				assertSlideInRange(context.engine, slideIndex);
				assertEditableShape(slideIndex, shapeIndex);
				options.slideIndex = slideIndex;
				options.shapeIndex = shapeIndex;
				result.affectedSlideIndices.add(slideIndex);
			} else {
				for (let index = 0; index < context.engine.slideCount; index++) {
					result.affectedSlideIndices.add(index);
				}
			}
			result.preview.push({ id: context.filePath, field: 'replaceText', before: query, after: replacement });
			if (!context.dryRun) {
				const count = await context.engine.replaceText(query, replacement, options);
				if (count === 0) {
					result.warnings.push(`No matches found for "${query}".`);
				}
			}
			result.changedIds.push(context.filePath);
			return result;
		}
		case 'pptx.setRunStyle': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const paragraphIndex = requireNumber(payload.paragraphIndex, 'paragraphIndex');
			const runIndex = requireNumber(payload.runIndex, 'runIndex');
			const style = payload.style as RunStyleChange;
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = `${pptxShapeId(slideIndex, shapeIndex)}/p:${paragraphIndex}/r:${runIndex}`;
			result.preview.push({ id: changedId, field: 'style', before: null, after: style });
			if (!context.dryRun) {
				await context.engine.setRunStyle(
					slideIndex,
					shapeIndex,
					{ paragraphIndex, runIndex },
					style,
				);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.setParagraphAlignment': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const paragraphIndex = requireNumber(payload.paragraphIndex, 'paragraphIndex');
			const align = requireString(payload.align, 'align') as 'l' | 'ctr' | 'r' | 'just';
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = `${pptxShapeId(slideIndex, shapeIndex)}/p:${paragraphIndex}`;
			result.preview.push({ id: changedId, field: 'align', before: null, after: align });
			if (!context.dryRun) {
				await context.engine.setParagraphAlignment(slideIndex, shapeIndex, paragraphIndex, align);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.applyListStyle': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const paragraphIndex = requireNumber(payload.paragraphIndex, 'paragraphIndex');
			const style = requireString(payload.style, 'style') as ParagraphListStyle;
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = `${pptxShapeId(slideIndex, shapeIndex)}/p:${paragraphIndex}`;
			result.preview.push({ id: changedId, field: 'listStyle', before: null, after: style });
			if (!context.dryRun) {
				await context.engine.applyListStyle(slideIndex, shapeIndex, paragraphIndex, style);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.updateTransform': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const transform = requireTransform(payload.transform);
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const before = tryFindShapeSnapshot(context.engine, context.filePath, slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			for (const field of ['x', 'y', 'cx', 'cy', 'rot'] as const) {
				const beforeValue = before?.transform[field] ?? null;
				if (beforeValue !== transform[field]) {
					result.preview.push({
						id: changedId,
						field: `transform.${field}`,
						before: beforeValue,
						after: transform[field],
					});
				}
			}
			if (!context.dryRun) {
				await context.engine.updateShapeTransform(slideIndex, shapeIndex, transform);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.setShapeFillColor': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const hex = requireString(payload.hex, 'hex').trim();
			if (!/^#?[0-9A-Fa-f]{6}$/.test(hex)) {
				throw createAiError(
					AI_ERROR_CODES.SCHEMA_INVALID,
					'hex must be a six-digit RGB color, with or without #.',
					{ field: 'hex' },
				);
			}
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			const normalizedHex = hex.replace(/^#/, '').toUpperCase();
			result.preview.push({
				id: changedId,
				field: 'fill',
				before: context.engine.getShapeVisualStyle(slideIndex, shapeIndex)?.fill ?? null,
				after: normalizedHex,
			});
			if (!context.dryRun) {
				await context.engine.setShapeFillColor(slideIndex, shapeIndex, normalizedHex);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.reorderShapes': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const mode = requireString(payload.mode, 'mode') as 'front' | 'back' | 'forward' | 'backward';
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			result.preview.push({ id: changedId, field: 'zOrder', before: mode, after: mode });
			if (!context.dryRun) {
				await context.engine.reorderShapes(slideIndex, [shapeIndex], mode);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.groupShapes': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndices = payload.shapeIndices;
			if (!Array.isArray(shapeIndices) || shapeIndices.length < 2) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'shapeIndices must contain at least two integers.');
			}
			const indices = shapeIndices.map((value, index) => requireNumber(value, `shapeIndices[${index}]`));
			assertSlideInRange(context.engine, slideIndex);
			for (const shapeIndex of indices) assertEditableShape(slideIndex, shapeIndex);
			if (!context.dryRun) {
				const groupIndex = await context.engine.groupShapes(slideIndex, indices);
				result.changedIds.push(pptxShapeId(slideIndex, groupIndex));
			} else {
				result.changedIds.push(pptxShapeId(slideIndex, indices[0] ?? 0));
			}
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.ungroupShapes': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			if (!context.dryRun) {
				const childIndices = await context.engine.ungroupShapes(slideIndex, shapeIndex);
				for (const childIndex of childIndices) {
					result.changedIds.push(pptxShapeId(slideIndex, childIndex));
				}
			} else {
				result.changedIds.push(pptxShapeId(slideIndex, shapeIndex));
			}
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.flipShape': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const axis = requireString(payload.axis, 'axis') as 'horizontal' | 'vertical';
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			result.preview.push({ id: changedId, field: 'flip', before: null, after: axis });
			if (!context.dryRun) {
				await context.engine.flipShape(slideIndex, shapeIndex, axis);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.addImage': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const vaultImagePath = requireString(payload.vaultImagePath, 'vaultImagePath');
			const transform = requireTransform(payload.transform);
			const fit = parseImageFit(payload, 'contain', ['contain', 'stretch'] as const);
			assertSlideInRange(context.engine, slideIndex);
			const image = await readVaultBinaryFile(context.vault, vaultImagePath);
			let shapeIndex = 0;
			if (!context.dryRun) {
				shapeIndex = await context.engine.addImage(slideIndex, image.bytes, getImageMimeType(image.extension));
				// Renderer-side insert returns a composite shape index that can diverge
				// from the serialized spTree (group children / graphicFrame merge). Use
				// the renderer transform path — not applyInsertedShapeTransform (OOXML).
				await context.engine.updateShapeTransform(
					slideIndex,
					shapeIndex,
					fit === 'contain' ? fitInsertedImageTransform(transform, image.bytes) : transform,
				);
			}
			result.changedIds.push(pptxShapeId(slideIndex, shapeIndex));
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.addShape': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const geometry = requireString(payload.geometry, 'geometry');
			const transform = requireTransform(payload.transform);
			if (!INSERTABLE_GEOMETRIES.has(geometry as InsertableShapeGeometry)) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, `Unsupported geometry: ${geometry}.`);
			}
			assertSlideInRange(context.engine, slideIndex);
			let shapeIndex = 0;
			if (!context.dryRun) {
				shapeIndex = await context.engine.addShapeGeometry(slideIndex, geometry as InsertableShapeGeometry);
				await applyTransformToInsertedShape(context.engine, slideIndex, shapeIndex, transform);
			}
			result.changedIds.push(pptxShapeId(slideIndex, shapeIndex));
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.addTextBox': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const transform = requireTransform(payload.transform);
			assertSlideInRange(context.engine, slideIndex);
			let shapeIndex = 0;
			if (!context.dryRun) {
				shapeIndex = await context.engine.addTextBox(slideIndex);
				await applyTransformToInsertedShape(context.engine, slideIndex, shapeIndex, transform);
			}
			result.changedIds.push(pptxShapeId(slideIndex, shapeIndex));
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.addTable': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const rows = requireNumber(payload.rows, 'rows');
			const cols = requireNumber(payload.cols, 'cols');
			const transform = requireTransform(payload.transform);
			assertSlideInRange(context.engine, slideIndex);
			let shapeIndex = 0;
			if (!context.dryRun) {
				shapeIndex = await context.engine.addTable(slideIndex, rows, cols);
				await applyTransformToInsertedShape(context.engine, slideIndex, shapeIndex, transform);
			}
			result.changedIds.push(pptxShapeId(slideIndex, shapeIndex));
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.addChart': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const transform = requireTransform(payload.transform);
			assertSlideInRange(context.engine, slideIndex);
			let shapeIndex = 0;
			if (!context.dryRun) {
				shapeIndex = await context.engine.addChart(slideIndex);
				await applyTransformToInsertedShape(context.engine, slideIndex, shapeIndex, transform);
			}
			result.changedIds.push(pptxShapeId(slideIndex, shapeIndex));
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.addSlide': {
			const afterIndex = requireNumber(payload.afterIndex, 'afterIndex');
			const layout = (typeof payload.layout === 'string' ? payload.layout : 'blank') as SlideLayoutKind;
			let slideIndex = Math.max(0, afterIndex + 1);
			if (!context.dryRun) {
				const moveResult = layout === 'blank'
					? await context.engine.addSlide(afterIndex)
					: await context.engine.addSlideWithLayout(afterIndex, layout);
				slideIndex = moveResult.slideIndex;
			}
			result.changedIds.push(`slide:${slideIndex}`);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.deleteSlide': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			assertSlideInRange(context.engine, slideIndex);
			if (!context.dryRun) {
				await context.engine.deleteSlide(slideIndex);
			}
			result.changedIds.push(`slide:${slideIndex}`);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.moveSlide': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const direction = requireNumber(payload.direction, 'direction') as -1 | 1;
			assertSlideInRange(context.engine, slideIndex);
			let targetIndex = slideIndex + direction;
			if (!context.dryRun) {
				const moved = await context.engine.moveSlide(slideIndex, direction);
				targetIndex = moved.slideIndex;
			}
			result.changedIds.push(`slide:${targetIndex}`);
			result.affectedSlideIndices.add(targetIndex);
			return result;
		}
		case 'pptx.duplicateSlide': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			assertSlideInRange(context.engine, slideIndex);
			let insertedIndex = slideIndex + 1;
			if (!context.dryRun) {
				const duplicated = await context.engine.duplicateSlide(slideIndex);
				insertedIndex = duplicated.slideIndex;
			}
			result.changedIds.push(`slide:${insertedIndex}`);
			result.affectedSlideIndices.add(insertedIndex);
			return result;
		}
		case 'pptx.reorderSlides': {
			const order = payload.order;
			if (!Array.isArray(order)) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'order must be an array of slide indices.');
			}
			const normalized = order.map((value, index) => requireNumber(value, `order[${index}]`));
			if (!context.dryRun) {
				await context.engine.reorderSlides(normalized);
			}
			for (const slideIndex of normalized) {
				result.changedIds.push(`slide:${slideIndex}`);
				result.affectedSlideIndices.add(slideIndex);
			}
			return result;
		}
		case 'pptx.setSlideBackground': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			assertSlideInRange(context.engine, slideIndex);
			const colorHex = typeof payload.colorHex === 'string' ? payload.colorHex : null;
			const vaultImagePath = typeof payload.vaultImagePath === 'string' ? payload.vaultImagePath : null;
			if (!colorHex && !vaultImagePath) {
				throw createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Provide colorHex or vaultImagePath.');
			}
			const changedId = `slide:${slideIndex}`;
			if (vaultImagePath) {
				const image = await readVaultBinaryFile(context.vault, vaultImagePath);
				result.preview.push({ id: changedId, field: 'backgroundImage', before: null, after: vaultImagePath });
				if (!context.dryRun) {
					await context.engine.setSlideBackgroundImage(slideIndex, image.bytes, getImageMimeType(image.extension));
				}
			} else if (colorHex) {
				result.preview.push({ id: changedId, field: 'backgroundColor', before: null, after: colorHex });
				if (!context.dryRun) {
					await context.engine.setSlideBackgroundColor(slideIndex, colorHex);
				}
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.setImageCrop': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const crop = payload.crop as ImageCrop;
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			result.preview.push({ id: changedId, field: 'crop', before: null, after: crop });
			if (!context.dryRun) {
				await context.engine.setImageCrop(slideIndex, shapeIndex, crop);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.fitImageToFrame': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			if (!context.engine.isImageShape(slideIndex, shapeIndex)) {
				throw createAiError(
					AI_ERROR_CODES.OBJECT_NOT_EDITABLE,
					`Shape ${pptxShapeId(slideIndex, shapeIndex)} is not an image.`,
				);
			}
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			const before = context.engine.getImageCrop(slideIndex, shapeIndex);
			const after = await context.engine.getImageFitCrop(slideIndex, shapeIndex);
			result.preview.push({ id: changedId, field: 'crop', before, after });
			if (!context.dryRun) {
				await context.engine.fitImageToFrame(slideIndex, shapeIndex);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.resetImage': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			if (!context.dryRun) {
				const reset = await context.engine.resetImage(slideIndex, shapeIndex);
				if (!reset.changed) {
					result.warnings.push(`Image ${changedId} already has no crop or visual effects to reset.`);
					return result;
				}
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.replaceImage': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const vaultImagePath = requireString(payload.vaultImagePath, 'vaultImagePath');
			const fit = parseImageFit(payload, 'cover', ['cover', 'stretch'] as const);
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const before = findShapeSnapshot(context.engine, context.filePath, slideIndex, shapeIndex);
			const wasImage = context.engine.isImageShape(slideIndex, shapeIndex);
			const image = await readVaultBinaryFile(context.vault, vaultImagePath);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			result.preview.push({
				id: changedId,
				field: wasImage ? 'image' : 'convertToImage',
				before: wasImage ? null : before,
				after: vaultImagePath,
			});
			if (!wasImage) {
				result.warnings.push(
					`Shape ${changedId} was not a picture; converted it into a picture filling the same transform box.`,
				);
			}
			if (!context.dryRun) {
				const resultIndex = await context.engine.replaceImage(
					slideIndex,
					shapeIndex,
					image.bytes,
					getImageMimeType(image.extension),
					fit,
				);
				result.changedIds.push(pptxShapeId(slideIndex, resultIndex));
			} else {
				result.changedIds.push(changedId);
			}
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		case 'pptx.updateChartData': {
			const slideIndex = requireNumber(payload.slideIndex, 'slideIndex');
			const shapeIndex = requireNumber(payload.shapeIndex, 'shapeIndex');
			const data = payload.data as ChartDataUpdate;
			assertSlideInRange(context.engine, slideIndex);
			assertEditableShape(slideIndex, shapeIndex);
			const changedId = pptxShapeId(slideIndex, shapeIndex);
			result.preview.push({ id: changedId, field: 'chartData', before: null, after: data });
			if (!context.dryRun) {
				await context.engine.updateChartData(slideIndex, shapeIndex, data);
			}
			result.changedIds.push(changedId);
			result.affectedSlideIndices.add(slideIndex);
			return result;
		}
		default:
			throw createAiError(AI_ERROR_CODES.UNKNOWN_OP, `Unknown PPTX operation: ${opId}.`, { op: opId });
	}
}
