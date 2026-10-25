declare module '@codemirror/view' {
	export class EditorView {
		constructor(config: { state: unknown; parent?: Element });
		readonly state: {
			doc: { toString(): string; length: number };
			selection: { main: { from: number; to: number } };
		};
		plugin(extension: unknown): { onKeydown?(event: KeyboardEvent, view: EditorView): void } | null;
		dispatch(spec: unknown): void;
		destroy(): void;
	}
}

declare module '@codemirror/state' {
	export class EditorState {
		static create(config: { doc?: string; extensions?: unknown }): EditorState;
	}
	export const EditorSelection: {
		single(anchor: number, head?: number): unknown;
	};
}
