import JSZip from 'jszip';

const REVIEW_MARKUP_PATTERN = /<w:(?:ins|del|moveFrom|moveTo|commentRangeStart|commentReference|rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange|sectPrChange|numberingChange|cellIns|cellDel|cellMerge)\b/;
const COMMENT_PATTERN = /<w:comment\b/;
const REVIEW_PART_PATTERNS = [
	/^word\/document\.xml$/,
	/^word\/headers\/header\d+\.xml$/,
	/^word\/footers\/footer\d+\.xml$/,
	/^word\/footnotes\.xml$/,
	/^word\/endnotes\.xml$/,
	/^word\/comments\.xml$/,
];

function isReviewPart(path: string): boolean {
	return REVIEW_PART_PATTERNS.some(pattern => pattern.test(path));
}

export async function hasReviewMarkup(buffer: ArrayBuffer) {
	try {
		const zip = await JSZip.loadAsync(buffer.slice(0));
		const reviewParts = Object.keys(zip.files).filter(isReviewPart);

		for (const partPath of reviewParts) {
			const xml = await zip.file(partPath)?.async('string');
			if (!xml) {
				continue;
			}
			if (REVIEW_MARKUP_PATTERN.test(xml) || (partPath === 'word/comments.xml' && COMMENT_PATTERN.test(xml))) {
				return true;
			}
		}

		return false;
	} catch {
		return false;
	}
}
