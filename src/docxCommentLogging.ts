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

export function extractDocxCommentPlainText(comment: DocxCommentLogSource): string {
	if (typeof comment.text === 'string' && comment.text.length > 0) {
		return truncateCommentText(comment.text);
	}

	const paragraphs = comment.content;
	if (!paragraphs?.length) {
		return '';
	}

	const parts: string[] = [];
	for (const paragraph of paragraphs) {
		for (const child of paragraph.content ?? []) {
			if (child.type !== 'run') continue;
			for (const runChild of child.content ?? []) {
				if (runChild.type === 'text' && typeof runChild.text === 'string') {
					parts.push(runChild.text);
				}
			}
		}
	}
	return truncateCommentText(parts.join(''));
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
