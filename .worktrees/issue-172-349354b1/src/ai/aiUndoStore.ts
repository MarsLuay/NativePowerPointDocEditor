export const AI_EDIT_UNDO_LABEL = 'AI edit';
export const AI_UNDO_LIMIT = 20;

export type AiUndoPayload =
	| { kind: 'pptx'; buffer: ArrayBuffer; currentSlide: number }
	| { kind: 'docx'; buffer: ArrayBuffer };

export interface AiUndoEntry {
	label: string;
	before: AiUndoPayload;
}

interface PathStacks {
	undo: AiUndoEntry[];
	redo: AiUndoEntry[];
}

export class AiUndoStore {
	private readonly stacks = new Map<string, PathStacks>();

	private getStacks(path: string): PathStacks {
		let stacks = this.stacks.get(path);
		if (!stacks) {
			stacks = { undo: [], redo: [] };
			this.stacks.set(path, stacks);
		}
		return stacks;
	}

	record(path: string, entry: AiUndoEntry): void {
		const stacks = this.getStacks(path);
		stacks.undo.push(entry);
		if (stacks.undo.length > AI_UNDO_LIMIT) {
			stacks.undo.shift();
		}
		stacks.redo = [];
	}

	canUndo(path: string): boolean {
		return (this.stacks.get(path)?.undo.length ?? 0) > 0;
	}

	canRedo(path: string): boolean {
		return (this.stacks.get(path)?.redo.length ?? 0) > 0;
	}

	popUndo(path: string): AiUndoEntry | null {
		const stacks = this.getStacks(path);
		return stacks.undo.pop() ?? null;
	}

	pushRedo(path: string, entry: AiUndoEntry): void {
		const stacks = this.getStacks(path);
		stacks.redo.push(entry);
		if (stacks.redo.length > AI_UNDO_LIMIT) {
			stacks.redo.shift();
		}
	}

	popRedo(path: string): AiUndoEntry | null {
		const stacks = this.getStacks(path);
		return stacks.redo.pop() ?? null;
	}

	drainUndo(path: string): AiUndoEntry[] {
		const stacks = this.stacks.get(path);
		if (!stacks || stacks.undo.length === 0) {
			return [];
		}
		return stacks.undo.splice(0);
	}

	clear(path: string): void {
		this.stacks.delete(path);
	}
}

export const aiUndoStore = new AiUndoStore();
