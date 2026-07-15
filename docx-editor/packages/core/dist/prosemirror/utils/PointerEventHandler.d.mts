/**
 * Pointer Event Handler
 *
 * Centralized input handling for all pointer events.
 * Provides single source of truth for click, drag, and focus management.
 * @packageDocumentation
 * @public
 */
import { ClickPositionResolver } from './ClickPositionResolver.mjs';

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
    #private;
    static readonly MULTI_CLICK_DELAY = 500;
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
