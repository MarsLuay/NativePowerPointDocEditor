import JSZip from 'jszip';
import { DOCX_DOCUMENT_PATH, listDocxDescribeParts } from './docxParts';
import { AI_ERROR_CODES, createAiError } from './errors';

export class DocxPatchSession {
	private readonly partXml = new Map<string, string>();

	private constructor(
		private readonly zip: JSZip,
		documentXml: string,
	) {
		this.partXml.set(DOCX_DOCUMENT_PATH, documentXml);
	}

	static async load(buffer: ArrayBuffer): Promise<DocxPatchSession> {
		const zip = await JSZip.loadAsync(buffer.slice(0));
		const documentXml = await zip.file(DOCX_DOCUMENT_PATH)?.async('string');
		if (!documentXml) {
			throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Missing word/document.xml in DOCX package.');
		}

		const session = new DocxPatchSession(zip, documentXml);
		for (const listed of listDocxDescribeParts(zip)) {
			if (listed.path === DOCX_DOCUMENT_PATH) continue;
			const xml = await zip.file(listed.path)?.async('string');
			if (xml) {
				session.partXml.set(listed.path, xml);
			}
		}
		return session;
	}

	getDocumentXml(): string {
		return this.getPartXml(DOCX_DOCUMENT_PATH);
	}

	setDocumentXml(documentXml: string): void {
		this.setPartXml(DOCX_DOCUMENT_PATH, documentXml);
	}

	getPartXml(partPath: string): string {
		const xml = this.partXml.get(partPath);
		if (!xml) {
			throw createAiError(AI_ERROR_CODES.VALIDATION_FAILED, `Missing DOCX part: ${partPath}.`);
		}
		return xml;
	}

	hasPart(partPath: string): boolean {
		return this.partXml.has(partPath);
	}

	setPartXml(partPath: string, partXml: string): void {
		this.partXml.set(partPath, partXml);
		this.zip.file(partPath, partXml);
	}

	listLoadedPartPaths(): string[] {
		return [...this.partXml.keys()].sort();
	}

	getZip(): JSZip {
		return this.zip;
	}

	async export(): Promise<ArrayBuffer> {
		return this.zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
	}
}
