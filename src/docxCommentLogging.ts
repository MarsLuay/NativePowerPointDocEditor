/**
 * Small comment payloads for debug logs / save summaries.
 * Includes replies (`parentId != null`) — top-level-only filters hide them.
 */

export interface DocxCommentLogSource {
	id: number;
	author?: string | null;
	date?: string | null;
	parentId?: number | null;
	done?: boolean | null;
	/** Plain text already extracted, or `content` paragraphs for extraction. */
	text?: string | null;
	content?: ReadonlyArray<{
		content?: ReadonlyArray<{
			type?: string;
			content?: ReadonlyArray<{ type?: string; text?: string }>;
		}>;
	}> | null;
}

export interface DocxCommentLogEntry {
	id: number;
	parentId: number | null;
	author: string | null;
	date: string | null;
	done: boolean;
	text: string;
	kind: 'comment' | 'reply';
}

export interface DocxCommentsLogSummary {
	total: number;
	topLevel: number;
	replies: number;
	done: number;
	comments: DocxCommentLogEntry[];
}

const MAX_COMMENT_TEXT_CHARS = 500;

type ParagraphNode = NonNullable<DocxCommentLogSource['content']>[number];
type ParagraphChildNode = NonNullable<ParagraphNode['content']>[number];
type RunChildNode = NonNullable<ParagraphChildNode['content']>[number];

function extractTextFromRunChild(runChild: RunChildNode): string {
	return runChild.type === 'text' && typeof runChild.text === 'string' ? runChild.text : '';
}

function extractTextFromParagraphChild(child: ParagraphChildNode): string {
	if (child.type !== 'run' || !child.content) {
		return '';
	}
	return child.content.map(extractTextFromRunChild).join('');
}

function extractTextFromParagraph(paragraph: ParagraphNode): string {
	if (!paragraph.content) {
		return '';
	}
	return paragraph.content.map(extractTextFromParagraphChild).join('');
}

export function extractDocxCommentPlainText(comment: DocxCommentLogSource): string {
	if (typeof comment.text === 'string' && comment.text.length > 0) {
		return truncateCommentText(comment.text);
	}

	const paragraphs = comment.content;
	if (!paragraphs?.length) {
		return '';
	}

	const rawText = paragraphs.map(extractTextFromParagraph).join('');
	return truncateCommentText(rawText);
}

export function truncateCommentText(text: string, maxChars = MAX_COMMENT_TEXT_CHARS): string {
	if (text.length <= maxChars) {
		return text;
	}
	return `${text.slice(0, maxChars)}…`;
}

export function summarizeDocxComment(comment: DocxCommentLogSource): DocxCommentLogEntry {
	const parentId = comment.parentId == null ? null : Number(comment.parentId);
	return {
		id: Number(comment.id),
		parentId,
		author: comment.author ?? null,
		date: comment.date ?? null,
		done: comment.done === true,
		text: extractDocxCommentPlainText(comment),
		kind: parentId == null ? 'comment' : 'reply',
	};
}

/** Full set: top-level comments and threaded replies. */
export function summarizeDocxComments(
	comments: ReadonlyArray<DocxCommentLogSource> | null | undefined,
): DocxCommentsLogSummary {
	const entries = (comments ?? []).map(summarizeDocxComment);
	const replies = entries.filter((entry) => entry.kind === 'reply').length;
	const done = entries.filter((entry) => entry.done).length;
	return {
		total: entries.length,
		topLevel: entries.length - replies,
		replies,
		done,
		comments: entries,
	};
}
