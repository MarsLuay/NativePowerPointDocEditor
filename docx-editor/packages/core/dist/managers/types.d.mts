/**
 * Manager Types
 *
 * Framework-agnostic interfaces for the editor's manager classes.
 * @packageDocumentation
 * @public
 */
import { EditorView } from 'prosemirror-view';
import { Document } from '../types/document.mjs';
import '../colors-C3vA7HUU.mjs';
import '../formatting-DFtuRFQY.mjs';
import '../lists-CyGxd5Y2.mjs';
import '../content-BZ9rYecc.mjs';
import '../docx/wrapTypes.mjs';
import '../watermark-D90356ZM.mjs';
import '../styles-Diw0MASy.mjs';

/**
 * Manager Types
 *
 * Framework-agnostic interfaces for the editor's manager classes.
 * @packageDocumentation
 * @public
 */

/**
 * Framework-agnostic interface for an imperatively mounted editor instance.
 *
 * Returned by `renderAsync()` implementations (React, Vue, etc.).
 * Consumers use this to interact with the editor programmatically.
 */
interface EditorHandle {
    /** Save the document and return the DOCX as a Blob. */
    save(): Promise<Blob | null>;
    /** Get the current parsed document model. */
    getDocument(): Document | null;
    /** Focus the editor. */
    focus(): void;
    /** Unmount the editor and clean up. */
    destroy(): void;
}
/** Auto-save status */
type AutoSaveStatus = 'idle' | 'saving' | 'saved' | 'error';
/** Configuration for AutoSaveManager */
interface AutoSaveManagerOptions {
    /** Storage key for localStorage (default: 'docx-editor-autosave') */
    storageKey?: string;
    /** Save interval in milliseconds (default: 30000 - 30 seconds) */
    interval?: number;
    /** Maximum age of auto-save before it's considered stale (default: 24 hours) */
    maxAge?: number;
    /** Whether to save on document change with debounce (default: true) */
    saveOnChange?: boolean;
    /** Debounce delay for saveOnChange in milliseconds (default: 2000) */
    debounceDelay?: number;
    /** Callback when save succeeds */
    onSave?: (timestamp: Date) => void;
    /** Callback when save fails */
    onError?: (error: Error) => void;
    /** Callback when recovery data is found */
    onRecoveryAvailable?: (savedDocument: SavedDocumentData) => void;
}
/** Saved document data structure */
interface SavedDocumentData {
    /** The document JSON */
    ['document']: Document;
    /** When the document was saved */
    savedAt: string;
    /** Version for format compatibility */
    version: number;
    /** Optional document identifier */
    documentId?: string;
}
/** AutoSaveManager snapshot for UI consumption */
interface AutoSaveSnapshot {
    status: AutoSaveStatus;
    lastSaveTime: Date | null;
    hasRecoveryData: boolean;
    isEnabled: boolean;
}
/** Cell coordinates in a table */
interface CellCoordinates {
    tableIndex: number;
    rowIndex: number;
    columnIndex: number;
}
/** TableSelectionManager snapshot */
interface TableSelectionSnapshot {
    /** Currently selected cell, or null if no selection */
    selectedCell: CellCoordinates | null;
}
/** Error severity levels */
type ErrorSeverity = 'error' | 'warning' | 'info';
/** Error notification */
interface ErrorNotification {
    id: string;
    message: string;
    severity: ErrorSeverity;
    details?: string;
    timestamp: number;
    dismissed?: boolean;
}
/** ErrorManager snapshot */
interface ErrorManagerSnapshot {
    notifications: ErrorNotification[];
}
/** Plugin lifecycle configuration */
interface PluginLifecycleConfig {
    id: string;
    styles?: string;
    initialize?: (editorView: EditorView) => unknown;
    onStateChange?: (editorView: EditorView) => unknown;
    destroy?: () => void;
}
/** PluginLifecycleManager snapshot */
interface PluginLifecycleSnapshot {
    /** Map of plugin ID to plugin state */
    states: Map<string, unknown>;
    /** Version counter (incremented on any state change) */
    version: number;
}

export type { AutoSaveManagerOptions, AutoSaveSnapshot, AutoSaveStatus, CellCoordinates, EditorHandle, ErrorManagerSnapshot, ErrorNotification, ErrorSeverity, PluginLifecycleConfig, PluginLifecycleSnapshot, SavedDocumentData, TableSelectionSnapshot };
