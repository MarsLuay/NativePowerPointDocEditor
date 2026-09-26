import { TFile, normalizePath, type Vault } from 'obsidian';

import type { PresentationEngine } from '../PresentationEngine';
import {
	exportSlidesToPdf,
	SLIDE_PDF_EXPORT_DPI,
	SLIDE_PDF_EXPORT_SCALE,
	slideSizeEmuToPdfPoints,
} from '../PowerPointExport';
import {
	getAvailableNumberedPath,
	getVaultFolderPrefix,
	sanitizeArtifactBaseName,
	writeVaultBinaryArtifact,
	type ArtifactWriteTarget,
} from '../export/artifactPaths';
import { debugLog } from '../logger';
import { createSvgElementFromString, sanitizeSvg } from '../SvgSecurity';
import { normalizeSvgForDisplay } from '../powerpoint/svgUtils';
import { AI_ERROR_CODES, createAiError, isAiErrorDetail, type AiErrorDetail } from './errors';

export type ExportPdfConflict = 'replace' | 'keep-both';

export interface ExportPdfOptions {
	/** Vault-relative output path. Defaults to `<sourceBasename>.pdf` beside the source. */
	outputPath?: string;
	/** 0-based slide indices. Defaults to every slide. */
	slideIndices?: number[];
	/** Conflict policy when the output path already exists. Default: `keep-both`. */
	conflict?: ExportPdfConflict;
	/** Legacy SVG-unit scale when slide EMU size is unavailable. */
	scale?: number;
	/** Raster DPI from physical slide size. Default 150. */
	dpi?: number;
}

export interface ExportPdfResult {
	ok: boolean;
	path?: string;
	bytes?: number;
	slideCount?: number;
	errors: AiErrorDetail[];
}

function resolveOwnerDocument(): Document {
	return activeDocument;
}

export function buildSlideSvgElementForExport(
	engine: PresentationEngine,
	slideIndex: number,
	ownerDocument: Document,
): SVGSVGElement | null {
	if (slideIndex < 0 || slideIndex >= engine.slideCount) {
		return null;
	}
	const rendered = engine.renderSlide(slideIndex);
	const scanned = sanitizeSvg(rendered.svg, { mode: 'compatibility' });
	if (!scanned.svg) {
		return null;
	}
	const element = createSvgElementFromString(scanned.svg, ownerDocument);
	if (!element) {
		return null;
	}
	engine.applyFontFidelity(element);
	if (typeof engine.formatChartAxisLabels === 'function') {
		engine.formatChartAxisLabels(element, slideIndex);
	}
	normalizeSvgForDisplay(element);
	return element;
}

export function resolveExportPdfOutputTarget(
	vault: Vault,
	sourceFile: TFile,
	options: ExportPdfOptions,
): ArtifactWriteTarget {
	const conflict = options.conflict ?? 'keep-both';
	const folderPrefix = getVaultFolderPrefix(sourceFile.parent?.path);
	const defaultBase = sanitizeArtifactBaseName(sourceFile.basename, 'presentation');
	const requestedPath = normalizePath(
		options.outputPath?.trim()
			? options.outputPath.trim()
			: `${folderPrefix}${defaultBase}.pdf`,
	);
	if (!requestedPath.toLowerCase().endsWith('.pdf')) {
		throw createAiError(
			AI_ERROR_CODES.SCHEMA_INVALID,
			`outputPath must end with .pdf (got ${requestedPath}).`,
			{ path: requestedPath, field: 'outputPath' },
		);
	}

	const exists = (candidate: string): boolean => Boolean(vault.getAbstractFileByPath(candidate));
	const existing = vault.getAbstractFileByPath(requestedPath);
	if (!existing) {
		return { path: requestedPath, existingFile: null, replace: false };
	}
	if (!(existing instanceof TFile)) {
		throw createAiError(
			AI_ERROR_CODES.FILE_EXISTS,
			`Output path exists and is not a file: ${requestedPath}.`,
			{ path: requestedPath },
		);
	}
	const existingFile = existing;
	if (conflict === 'replace') {
		return { path: requestedPath, existingFile, replace: true };
	}
	return {
		path: getAvailableNumberedPath(requestedPath, exists),
		existingFile: null,
		replace: false,
	};
}

export async function exportPresentationToPdfBytes(
	engine: PresentationEngine,
	options: Pick<ExportPdfOptions, 'slideIndices' | 'scale' | 'dpi'> = {},
	ownerDocument: Document = resolveOwnerDocument(),
): Promise<{ bytes: ArrayBuffer; slideCount: number }> {
	const indices =
		options.slideIndices && options.slideIndices.length > 0
			? options.slideIndices
			: Array.from({ length: engine.slideCount }, (_, index) => index);
	if (indices.length === 0) {
		throw createAiError(AI_ERROR_CODES.SLIDE_NOT_FOUND, 'Presentation has no slides to export.');
	}

	const elements: SVGSVGElement[] = [];
	for (const index of indices) {
		if (!Number.isInteger(index) || index < 0 || index >= engine.slideCount) {
			throw createAiError(
				AI_ERROR_CODES.SLIDE_NOT_FOUND,
				`Slide index out of range: ${index} (slideCount=${engine.slideCount}).`,
				{ field: 'slideIndices' },
			);
		}
		const element = buildSlideSvgElementForExport(engine, index, ownerDocument);
		if (!element) {
			throw createAiError(
				AI_ERROR_CODES.VALIDATION_FAILED,
				`Slide ${index} could not be rendered for PDF export.`,
				{ field: 'slideIndices' },
			);
		}
		elements.push(element);
	}

	const scale = options.scale ?? SLIDE_PDF_EXPORT_SCALE;
	const dpi = options.dpi ?? SLIDE_PDF_EXPORT_DPI;
	let pageSizePoints: { width: number; height: number } | undefined;
	try {
		const slideSize = await engine.getSlideSizeEmu();
		pageSizePoints = slideSizeEmuToPdfPoints(slideSize.cx, slideSize.cy);
	} catch {
		pageSizePoints = undefined;
	}
	debugLog('export', 'AI PPTX PDF export prepared slides', {
		slideCount: elements.length,
		slideIndices: indices,
		scale,
		dpi,
		pageSizePoints,
	});
	const bytes = await exportSlidesToPdf(elements, ownerDocument, {
		scale,
		dpi,
		...(pageSizePoints ? { pageSizePoints } : {}),
	});
	return { bytes, slideCount: elements.length };
}

export async function writeExportPdfArtifact(
	vault: Vault,
	sourceFile: TFile,
	bytes: ArrayBuffer,
	options: ExportPdfOptions,
): Promise<{ path: string; bytes: number }> {
	const target = resolveExportPdfOutputTarget(vault, sourceFile, options);
	const written = await writeVaultBinaryArtifact(vault, target, bytes);
	debugLog('export', 'AI NPDE PDF export written', {
		op: 'ai-export-pdf',
		sourcePath: sourceFile.path,
		targetPath: written.path,
		bytes: bytes.byteLength,
		replaced: target.replace,
	});
	return { path: written.path, bytes: bytes.byteLength };
}

export function toExportPdfFailure(error: unknown, path?: string): ExportPdfResult {
	if (isAiErrorDetail(error)) {
		return { ok: false, errors: [error] };
	}
	return {
		ok: false,
		errors: [
			createAiError(AI_ERROR_CODES.VALIDATION_FAILED, String(error), path ? { path } : {}),
		],
	};
}
