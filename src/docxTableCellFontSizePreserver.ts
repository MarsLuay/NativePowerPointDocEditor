import JSZip from 'jszip';

interface XmlRange {
	start: number;
	end: number;
	xml: string;
}

export type DocxTableCellFontSizePreservationStatus =
	| 'source-unavailable'
	| 'document-part-unavailable'
	| 'no-table-cells'
	| 'cell-count-mismatch'
	| 'checked'
	| 'restored';

export interface PreservedDocxTableCellFontSizes {
	buffer: ArrayBuffer;
	restoredRuns: number;
	restoredTags: number;
	status: DocxTableCellFontSizePreservationStatus;
	sourceCellCount: number;
	outputCellCount: number;
	matchedCellCount: number;
	skippedTextChangedCells: number;
	skippedRunCountChangedCells: number;
	sourceRunsWithDirectSize: number;
}

function isNameChar(char: string | undefined): boolean {
	return Boolean(char && !/\s|\/|>/.test(char));
}

function readXmlTag(xml: string, start: number) {
	if (xml[start] !== '<' || xml[start + 1] === '!' || xml[start + 1] === '?') {
		return null;
	}

	const close = xml.indexOf('>', start + 1);
	if (close < 0) {
		return null;
	}

	const isClosing = xml[start + 1] === '/';
	const nameStart = start + (isClosing ? 2 : 1);
	let nameEnd = nameStart;
	while (isNameChar(xml[nameEnd])) {
		nameEnd += 1;
	}

	if (nameEnd === nameStart) {
		return null;
	}

	const rawTag = xml.slice(start, close + 1);
	return {
		close,
		isClosing,
		isSelfClosing: !isClosing && /\/\s*>$/.test(rawTag),
		name: xml.slice(nameStart, nameEnd),
	};
}

function findXmlElementRanges(xml: string, elementName: string): XmlRange[] {
	const ranges: XmlRange[] = [];
	let depth = 0;
	let rangeStart = -1;
	let index = 0;

	while (index < xml.length) {
		const tagStart = xml.indexOf('<', index);
		if (tagStart < 0) {
			break;
		}

		const tag = readXmlTag(xml, tagStart);
		if (!tag) {
			index = tagStart + 1;
			continue;
		}

		if (tag.name === elementName) {
			if (tag.isClosing) {
				if (depth > 0) {
					depth -= 1;
					if (depth === 0 && rangeStart >= 0) {
						const end = tag.close + 1;
						ranges.push({ start: rangeStart, end, xml: xml.slice(rangeStart, end) });
						rangeStart = -1;
					}
				}
			} else if (!tag.isSelfClosing) {
				if (depth === 0) {
					rangeStart = tagStart;
				}
				depth += 1;
			}
		}

		index = tag.close + 1;
	}

	return ranges;
}

function decodeXmlText(value: string) {
	return value
		.replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(Number.parseInt(hex, 16)))
		.replace(/&#([0-9]+);/g, (_, decimal: string) => String.fromCodePoint(Number.parseInt(decimal, 10)))
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, '&');
}

function extractVisibleText(xml: string) {
	const parts: string[] = [];
	const textPattern = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
	let match: RegExpExecArray | null;

	while ((match = textPattern.exec(xml)) !== null) {
		parts.push(decodeXmlText(match[1] ?? ''));
	}

	return parts.join('');
}

function getDirectRunSizeTags(runXml: string) {
	const runProperties = runXml.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/)?.[0] ?? '';
	if (!runProperties) {
		return [];
	}

	const tags: string[] = [];
	for (const tagName of ['w:sz', 'w:szCs']) {
		const tag = runProperties.match(new RegExp(`<${tagName}\\b[^>]*/>`))?.[0];
		if (tag) {
			tags.push(tag);
		}
	}

	return tags;
}

function tagNameFromSelfClosingTag(tag: string) {
	return tag.match(/^<([^\s/>]+)/)?.[1] ?? '';
}

function runHasDirectSizeTag(runXml: string, tagName: string) {
	return new RegExp(`<${tagName}\\b[^>]*/>`).test(runXml);
}

function addMissingSizeTagsToRun(runXml: string, sourceTags: string[]) {
	const missingTags = sourceTags.filter((tag) => {
		const tagName = tagNameFromSelfClosingTag(tag);
		return tagName && !runHasDirectSizeTag(runXml, tagName);
	});

	if (missingTags.length === 0) {
		return { xml: runXml, restoredTags: 0 };
	}

	if (/<w:rPr\b[^>]*>/.test(runXml)) {
		return {
			xml: runXml.replace('</w:rPr>', `${missingTags.join('')}</w:rPr>`),
			restoredTags: missingTags.length,
		};
	}

	const runStartEnd = runXml.indexOf('>');
	if (runStartEnd < 0) {
		return { xml: runXml, restoredTags: 0 };
	}

	return {
		xml: `${runXml.slice(0, runStartEnd + 1)}<w:rPr>${missingTags.join('')}</w:rPr>${runXml.slice(runStartEnd + 1)}`,
		restoredTags: missingTags.length,
	};
}

function preserveCellRunFontSizes(sourceCellXml: string, outputCellXml: string) {
	if (extractVisibleText(sourceCellXml) !== extractVisibleText(outputCellXml)) {
		return {
			xml: outputCellXml,
			restoredRuns: 0,
			restoredTags: 0,
			matchStatus: 'text-changed' as const,
			sourceRunsWithDirectSize: 0,
		};
	}

	const sourceRuns = findXmlElementRanges(sourceCellXml, 'w:r');
	const outputRuns = findXmlElementRanges(outputCellXml, 'w:r');
	if (sourceRuns.length !== outputRuns.length) {
		return {
			xml: outputCellXml,
			restoredRuns: 0,
			restoredTags: 0,
			matchStatus: 'run-count-changed' as const,
			sourceRunsWithDirectSize: 0,
		};
	}

	let nextCellXml = outputCellXml;
	let restoredRuns = 0;
	let restoredTags = 0;
	let sourceRunsWithDirectSize = 0;

	for (let index = outputRuns.length - 1; index >= 0; index -= 1) {
		const sourceRun = sourceRuns[index];
		const outputRun = outputRuns[index];
		if (!sourceRun || !outputRun || extractVisibleText(sourceRun.xml) !== extractVisibleText(outputRun.xml)) {
			continue;
		}

		const sourceTags = getDirectRunSizeTags(sourceRun.xml);
		if (sourceTags.length === 0) {
			continue;
		}
		sourceRunsWithDirectSize += 1;

		const patched = addMissingSizeTagsToRun(outputRun.xml, sourceTags);
		if (patched.restoredTags === 0) {
			continue;
		}

		nextCellXml = `${nextCellXml.slice(0, outputRun.start)}${patched.xml}${nextCellXml.slice(outputRun.end)}`;
		restoredRuns += 1;
		restoredTags += patched.restoredTags;
	}

	return {
		xml: nextCellXml,
		restoredRuns,
		restoredTags,
		matchStatus: 'matched' as const,
		sourceRunsWithDirectSize,
	};
}

function preserveDocumentTableCellFontSizes(sourceDocumentXml: string, outputDocumentXml: string) {
	const sourceCells = findXmlElementRanges(sourceDocumentXml, 'w:tc');
	const outputCells = findXmlElementRanges(outputDocumentXml, 'w:tc');
	const diagnostics = {
		sourceCellCount: sourceCells.length,
		outputCellCount: outputCells.length,
		matchedCellCount: 0,
		skippedTextChangedCells: 0,
		skippedRunCountChangedCells: 0,
		sourceRunsWithDirectSize: 0,
	};
	if (sourceCells.length === 0) {
		return {
			xml: outputDocumentXml,
			restoredRuns: 0,
			restoredTags: 0,
			status: 'no-table-cells' as const,
			...diagnostics,
		};
	}
	if (sourceCells.length !== outputCells.length) {
		return {
			xml: outputDocumentXml,
			restoredRuns: 0,
			restoredTags: 0,
			status: 'cell-count-mismatch' as const,
			...diagnostics,
		};
	}

	let nextDocumentXml = outputDocumentXml;
	let restoredRuns = 0;
	let restoredTags = 0;

	for (let index = outputCells.length - 1; index >= 0; index -= 1) {
		const sourceCell = sourceCells[index];
		const outputCell = outputCells[index];
		if (!sourceCell || !outputCell) {
			continue;
		}

		const patched = preserveCellRunFontSizes(sourceCell.xml, outputCell.xml);
		if (patched.matchStatus === 'text-changed') {
			diagnostics.skippedTextChangedCells += 1;
		} else if (patched.matchStatus === 'run-count-changed') {
			diagnostics.skippedRunCountChangedCells += 1;
		} else {
			diagnostics.matchedCellCount += 1;
			diagnostics.sourceRunsWithDirectSize += patched.sourceRunsWithDirectSize;
		}
		if (patched.restoredTags === 0) {
			continue;
		}

		nextDocumentXml = `${nextDocumentXml.slice(0, outputCell.start)}${patched.xml}${nextDocumentXml.slice(outputCell.end)}`;
		restoredRuns += patched.restoredRuns;
		restoredTags += patched.restoredTags;
	}

	return {
		xml: nextDocumentXml,
		restoredRuns,
		restoredTags,
		status: restoredTags > 0 ? 'restored' as const : 'checked' as const,
		...diagnostics,
	};
}

export async function preserveDocxTableCellFontSizes(sourceBuffer: ArrayBuffer | null | undefined, outputBuffer: ArrayBuffer): Promise<PreservedDocxTableCellFontSizes> {
	if (!sourceBuffer) {
		return {
			buffer: outputBuffer,
			restoredRuns: 0,
			restoredTags: 0,
			status: 'source-unavailable',
			sourceCellCount: 0,
			outputCellCount: 0,
			matchedCellCount: 0,
			skippedTextChangedCells: 0,
			skippedRunCountChangedCells: 0,
			sourceRunsWithDirectSize: 0,
		};
	}

	const [sourceZip, outputZip] = await Promise.all([
		JSZip.loadAsync(sourceBuffer.slice(0)),
		JSZip.loadAsync(outputBuffer.slice(0)),
	]);
	const sourceDocument = sourceZip.file('word/document.xml');
	const outputDocument = outputZip.file('word/document.xml');
	if (!sourceDocument || !outputDocument) {
		return {
			buffer: outputBuffer,
			restoredRuns: 0,
			restoredTags: 0,
			status: 'document-part-unavailable',
			sourceCellCount: 0,
			outputCellCount: 0,
			matchedCellCount: 0,
			skippedTextChangedCells: 0,
			skippedRunCountChangedCells: 0,
			sourceRunsWithDirectSize: 0,
		};
	}

	const [sourceDocumentXml, outputDocumentXml] = await Promise.all([
		sourceDocument.async('string'),
		outputDocument.async('string'),
	]);
	const patched = preserveDocumentTableCellFontSizes(sourceDocumentXml, outputDocumentXml);
	if (patched.restoredTags === 0) {
		return {
			buffer: outputBuffer,
			restoredRuns: 0,
			restoredTags: 0,
			status: patched.status,
			sourceCellCount: patched.sourceCellCount,
			outputCellCount: patched.outputCellCount,
			matchedCellCount: patched.matchedCellCount,
			skippedTextChangedCells: patched.skippedTextChangedCells,
			skippedRunCountChangedCells: patched.skippedRunCountChangedCells,
			sourceRunsWithDirectSize: patched.sourceRunsWithDirectSize,
		};
	}

	outputZip.file('word/document.xml', patched.xml);
	return {
		buffer: await outputZip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' }),
		restoredRuns: patched.restoredRuns,
		restoredTags: patched.restoredTags,
		status: patched.status,
		sourceCellCount: patched.sourceCellCount,
		outputCellCount: patched.outputCellCount,
		matchedCellCount: patched.matchedCellCount,
		skippedTextChangedCells: patched.skippedTextChangedCells,
		skippedRunCountChangedCells: patched.skippedRunCountChangedCells,
		sourceRunsWithDirectSize: patched.sourceRunsWithDirectSize,
	};
}
