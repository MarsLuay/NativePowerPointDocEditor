export function docxParagraphId(index: number): string {
	return `body/p[${index}]`;
}

export function docxRunId(paragraphIndex: number, runIndex: number): string {
	return `${docxParagraphId(paragraphIndex)}/r[${runIndex}]`;
}

export function docxTableId(index: number): string {
	return `body/tbl[${index}]`;
}

export function docxCellId(tableIndex: number, rowIndex: number, colIndex: number): string {
	return `${docxTableId(tableIndex)}/tr[${rowIndex}]/tc[${colIndex}]`;
}
