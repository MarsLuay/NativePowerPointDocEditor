import type { TFile, Vault } from 'obsidian';
import JSZip from 'jszip';
import { describeDocxFromBuffer } from './docxDescribe';
import { DocxPatchSession } from './docxPatchSession';
import { parseDocumentBody } from './docxOoxml';
import { extractDocxText } from '../docxTextExtractor';
import { AI_ERROR_CODES, createAiError } from './errors';

const REQUIRED_DOCX_PARTS = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml'] as const;

export function validateDocxDocumentXmlLight(documentXml: string): void {
	if (!/<w:document\b/.test(documentXml) || !/<w:body\b/.test(documentXml)) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'document.xml is missing w:document or w:body.');
	}
	parseDocumentBody(documentXml);
}

export async function validateDocxPackage(output: ArrayBuffer): Promise<void> {
	const zip = await JSZip.loadAsync(output.slice(0));
	for (const part of REQUIRED_DOCX_PARTS) {
		if (!zip.file(part)) {
			throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, `Exported DOCX is missing ${part}.`);
		}
	}
	const documentXml = await zip.file('word/document.xml')?.async('string');
	if (!documentXml) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Exported DOCX is missing word/document.xml.');
	}
	validateDocxDocumentXmlLight(documentXml);
}

export async function validateDocxRoundTrip(sourceBuffer: ArrayBuffer, output: ArrayBuffer): Promise<void> {
	await validateDocxPackage(output);
	await DocxPatchSession.load(output);

	const snapshot = await describeDocxFromBuffer(output, 'roundtrip.docx');
	if (!Number.isInteger(snapshot.blockCount) || snapshot.blockCount < 0) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'DOCX describe() returned an invalid block count after export.');
	}

	const outputText = await extractDocxText(output);
	const sourceText = await extractDocxText(sourceBuffer);
	if (sourceText.length > 0 && outputText.length === 0) {
		throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'DOCX export lost all extractable text.');
	}

	const sourceZip = await JSZip.loadAsync(sourceBuffer.slice(0));
	const outputZip = await JSZip.loadAsync(output.slice(0));
	for (const partPath of Object.keys(sourceZip.files)) {
		if (!partPath.startsWith('word/media/')) continue;
		if (!outputZip.file(partPath) && !hasReplacementMedia(outputZip, partPath)) {
			throw createAiError(
				AI_ERROR_CODES.VALIDATION_FAILED,
				`DOCX export dropped existing media part ${partPath}.`,
			);
		}
	}
}

function hasReplacementMedia(zip: JSZip, originalPath: string): boolean {
	const mediaDirectory = 'word/media/';
	return Object.keys(zip.files).some((path) => path.startsWith(mediaDirectory) && path !== originalPath);
}

export async function saveDocxToVault(
	vault: Vault,
	file: TFile,
	session: DocxPatchSession,
	sourceBuffer: ArrayBuffer,
): Promise<ArrayBuffer> {
	const output = await session.export();
	await validateDocxRoundTrip(sourceBuffer, output);
	await vault.modifyBinary(file, output);
	return output;
}
