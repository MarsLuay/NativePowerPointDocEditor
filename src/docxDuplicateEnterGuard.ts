/**
 * One physical Enter on iPad was reaching the editor twice, about 200ms apart,
 * and each delivery split a paragraph. Key repeat stays allowed. A second
 * non-repeat Enter with no keyup in between is the duplicate path. A real tap
 * always keyups first, so five quick Enters still make five lines.
 */

export const DUPLICATE_ENTER_WINDOW_MS = 350;

const DOCX_EDITING_TARGET_SELECTOR = [
	'.paged-editor__hidden-pm',
	'.docx-editor-root',
	'[data-testid="docx-editor"]',
].join(', ');

export interface ParagraphEnterLike {
	key?: string;
	code?: string;
	repeat?: boolean;
	shiftKey?: boolean;
	altKey?: boolean;
	ctrlKey?: boolean;
	metaKey?: boolean;
	isComposing?: boolean;
	inputType?: string;
}

export interface DuplicateEnterGuard {
	observeKeyDown(
		event: ParagraphEnterLike,
		now: number,
		paragraphCount: number | null,
	): 'allow' | 'suppress';
	observeKeyUp(event: ParagraphEnterLike): void;
	observeBeforeInput(
		event: ParagraphEnterLike,
		paragraphCount: number,
		now: number,
	): 'allow' | 'suppress';
}

export function isDocxEditingTarget(target: EventTarget | null): boolean {
	const element = targetElement(target);
	return Boolean(element?.closest(DOCX_EDITING_TARGET_SELECTOR));
}

export function isPlainEnterKey(event: ParagraphEnterLike): boolean {
	if (event.isComposing || event.shiftKey || event.altKey || event.ctrlKey || event.metaKey) {
		return false;
	}
	return event.key === 'Enter' || event.code === 'NumpadEnter' || event.code === 'Enter';
}

export function createDuplicateEnterGuard(): DuplicateEnterGuard {
	let armedUntil = 0;
	let paragraphCountAtArm: number | null = null;
	let waitingForKeyUp = false;

	return {
		observeKeyDown(event, now, paragraphCount) {
			if (!isPlainEnterKey(event)) {
				return 'allow';
			}
			if (event.repeat === true) {
				return 'allow';
			}
			if (waitingForKeyUp && now < armedUntil) {
				return 'suppress';
			}
			waitingForKeyUp = true;
			armedUntil = now + DUPLICATE_ENTER_WINDOW_MS;
			paragraphCountAtArm = paragraphCount;
			return 'allow';
		},
		observeKeyUp(event) {
			if (!isPlainEnterKey(event)) {
				return;
			}
			waitingForKeyUp = false;
		},
		observeBeforeInput(event, paragraphCount, now) {
			if (event.isComposing || event.inputType !== 'insertParagraph') {
				return 'allow';
			}
			if (paragraphCountAtArm === null || now >= armedUntil) {
				return 'allow';
			}
			if (paragraphCount > paragraphCountAtArm) {
				return 'suppress';
			}
			return 'allow';
		},
	};
}

function targetElement(target: EventTarget | null): Element | null {
	if (!target || typeof target !== 'object') {
		return null;
	}
	if ('closest' in target && typeof (target as Element).closest === 'function') {
		return target as Element;
	}
	const parent = 'parentElement' in target ? (target as { parentElement?: Element | null }).parentElement : null;
	return parent ?? null;
}
