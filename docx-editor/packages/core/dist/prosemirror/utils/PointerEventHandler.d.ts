/**
 * Pointer Event Handler
 *
 * Centralized input handling for all pointer events.
 * Provides single source of truth for click, drag, and focus management.
 * @packageDocumentation
 * @public
 */
import { ClickPositionResolver } from './ClickPositionResolver.js';

/**
 * Pointer Event Handler
 *
 * Centralized input handling for all pointer events.
 * Provides single source of truth for click, drag, and focus management.
 * @packageDocumentation
 * @public
 */

/**
 * Interface for the editor that the handler controls.
 */
interface EditorInterface {
    /** Set selection to a position (collapsed) */
    setSelection(pos: number): void;
    /** Set selection range */
    setSelectionRange(from: number, to: number): void;
    /** Get current selection */
    getSelection(): {
        from: number;
        to: number;
    } | null;
    /** Focus the editor */
    focus(): void;
}
/**
 * Callback for input events.
 */
type InputEventCallback = (event: {
    type: 'click' | 'doubleClick' | 'tripleClick' | 'dragStart' | 'drag' | 'dragEnd';
    position?: number;
    from?: number;
    to?: number;
}) => void;
/**
 * Options for PointerEventHandler.
 */
interface PointerEventHandlerOptions {
    /** The editor to control */
    editor: EditorInterface;
    /** Position resolver for click mapping */
    positionResolver: ClickPositionResolver;
    /** Callback for input events */
    onInput?: InputEventCallback;
}
/**
 * PointerEventHandler handles all pointer input for the paged editor.
 * It provides:
 * - Single/double/triple click detection
 * - Drag selection with anchor tracking
 * - Coordinate normalization for zoom
 * - Focus management
 */
declare class PointerEventHandler {
    private editor;
    private positionResolver;
    private container;
    private onInput?;
    private dragging;
    private dragAnchor;
    private lastClickTime;
    private lastClickPos;
    private clickCount;
    static readonly MULTI_CLICK_DELAY = 500;
    private boundPointerDown;
    private boundPointerMove;
    private boundPointerUp;
    private boundPointerLeave;
    constructor(options: PointerEventHandlerOptions);
    /**
     * Attach event listeners to a container element.
     */
    attach(container: HTMLElement): void;
    /**
     * Detach event listeners from the container.
     */
    detach(): void;
    /**
     * Update the position resolver reference.
     */
    setPositionResolver(positionResolver: ClickPositionResolver): void;
    /**
     * Get position from client coordinates, accounting for zoom.
     */
    private getPositionFromCoords;
    /**
     * Handle pointer down - start selection or drag.
     */
    private onPointerDown;
    /**
     * Handle pointer move - extend drag selection.
     */
    private onPointerMove;
    /**
     * Handle pointer up - end drag.
     */
    private onPointerUp;
    /**
     * Handle pointer leave - end drag if leaving container.
     */
    private onPointerLeave;
    /**
     * Select the word at a position.
     */
    private selectWord;
    /**
     * Select the paragraph at a position.
     */
    private selectParagraph;
    /**
     * Emit an input event.
     */
    private emitEvent;
    /**
     * Get current drag state.
     */
    isDragging(): boolean;
    /**
     * Get drag anchor position.
     */
    getDragAnchor(): number | null;
    /**
     * Cancel any ongoing drag.
     */
    cancelDrag(): void;
}

export { type EditorInterface, type InputEventCallback, PointerEventHandler, type PointerEventHandlerOptions };
