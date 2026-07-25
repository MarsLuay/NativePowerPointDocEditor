/**
 * Layout Selection Gate
 *
 * Guards selection rendering until layout is up-to-date.
 * Uses sequenced versioning to prevent stale cursor positions.
 * @packageDocumentation
 * @public
 */
type RenderCallback = () => void;
/**
 * LayoutSelectionGate coordinates the timing between document edits and
 * layout reflow so that selection overlays are only painted against
 * current DOM geometry.
 *
 * Workflow:
 * 1. Document changes → setStateSeq(++seq)
 * 2. Layout starts → onLayoutStart()
 * 3. Layout completes → onLayoutComplete(seq)
 * 4. Selection update requested → requestRender()
 * 5. If safe → callback is called
 */
declare class LayoutSelectionGate {
    /** Current document state sequence */
    private stateSeq;
    /** Last painted layout sequence */
    private renderSeq;
    /** Whether layout is currently being computed/painted */
    private layoutUpdating;
    /** Pending render callback */
    private pendingRender;
    /** Registered render callbacks */
    private renderCallbacks;
    /**
     * Set the document state sequence (call when document changes).
     * This should be called on every ProseMirror transaction that changes the doc.
     */
    setStateSeq(seq: number): void;
    /**
     * Increment document state sequence (convenience method).
     * Returns the new sequence value.
     */
    incrementStateSeq(): number;
    /**
     * Get current document state sequence.
     */
    getStateSeq(): number;
    /**
     * Get current layout render sequence.
     */
    getRenderSeq(): number;
    /**
     * Called when layout computation starts.
     */
    onLayoutStart(): void;
    /**
     * Called when layout computation and DOM painting completes.
     * @param seq - The document state sequence that was just painted
     */
    onLayoutComplete(seq: number): void;
    /**
     * Check if it's safe to render selection.
     * Safe when: layout is not updating AND render sequence >= state sequence
     */
    isSafeToRender(): boolean;
    /**
     * Request a selection render. Will be executed when safe.
     * If already safe, executes immediately.
     */
    requestRender(): void;
    /**
     * Register a callback to be called on render events.
     */
    onRender(callback: RenderCallback): () => void;
    /**
     * Try to execute pending render if safe.
     */
    private tryRender;
    /**
     * Execute all registered render callbacks.
     */
    private executeRender;
    /**
     * Reset the gate state (useful for testing or document reload).
     */
    reset(): void;
    /**
     * Get debug info about current state.
     */
    getDebugInfo(): {
        stateSeq: number;
        renderSeq: number;
        layoutUpdating: boolean;
        hasPendingRender: boolean;
        isSafe: boolean;
    };
}

export { LayoutSelectionGate };
