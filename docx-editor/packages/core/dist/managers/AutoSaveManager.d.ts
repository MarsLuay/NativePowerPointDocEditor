/**
 * AutoSaveManager
 *
 * Framework-agnostic class for auto-saving documents to localStorage.
 * Extracted from the React `useAutoSave` hook.
 *
 * Usage with React:
 * ```ts
 * const snapshot = useSyncExternalStore(manager.subscribe, manager.getSnapshot);
 * ```
 * @packageDocumentation
 * @public
 */
import { S as Subscribable } from '../Subscribable-DOz6Ohoo.js';
import { AutoSaveSnapshot, AutoSaveManagerOptions, SavedDocumentData, AutoSaveStatus } from './types.js';
import { Document } from '../types/document.js';
import 'prosemirror-view';
import '../colors-C3vA7HUU.js';
import '../formatting-JhqWT_XM.js';
import '../lists-Bn29SzeS.js';
import '../content-B8ScSBzC.js';
import '../docx/wrapTypes.js';
import '../watermark-D90356ZM.js';
import '../styles-BGGjYqnf.js';

/**
 * AutoSaveManager
 *
 * Framework-agnostic class for auto-saving documents to localStorage.
 * Extracted from the React `useAutoSave` hook.
 *
 * Usage with React:
 * ```ts
 * const snapshot = useSyncExternalStore(manager.subscribe, manager.getSnapshot);
 * ```
 * @packageDocumentation
 * @public
 */

declare class AutoSaveManager extends Subscribable<AutoSaveSnapshot> {
    private storageKey;
    private interval;
    private maxAge;
    private saveOnChange;
    private debounceDelay;
    private onSaveCallback?;
    private onErrorCallback?;
    private onRecoveryAvailableCallback?;
    private storageAvailable;
    private currentDocument;
    private lastSavedJson;
    private intervalTimer;
    private debounceTimer;
    private status;
    private lastSaveTime;
    private _hasRecoveryData;
    private _isEnabled;
    constructor(options?: AutoSaveManagerOptions);
    /** Update the current document. Triggers debounced save if enabled. */
    onDocumentChanged(document: Document | null): void;
    /** Manually trigger a save. */
    save(): Promise<boolean>;
    /** Clear auto-saved data from storage. */
    clear(): void;
    /** Get recovery data from storage. */
    getRecoveryData(): SavedDocumentData | null;
    /** Accept recovery and return the document. */
    acceptRecovery(): Document | null;
    /** Dismiss recovery and clear saved data. */
    dismissRecovery(): void;
    /** Enable auto-save and start the interval timer. */
    enable(): void;
    /** Disable auto-save and stop all timers. */
    disable(): void;
    /** Start the interval timer. Call after enabling or on init. */
    startInterval(): void;
    /** Save synchronously on destroy (best-effort). */
    destroy(): void;
    private checkRecoveryData;
    private persistToStorage;
    private debounceSave;
    private stopTimers;
    private updateStatus;
    private emitSnapshot;
}
/** Format last save time for display */
declare function formatLastSaveTime(date: Date | null): string;
/** Get auto-save status label */
declare function getAutoSaveStatusLabel(status: AutoSaveStatus): string;
/** Get storage size used by auto-save */
declare function getAutoSaveStorageSize(storageKey?: string): number;
/** Format storage size for display */
declare function formatStorageSize(bytes: number): string;
/** Check if auto-save is supported */
declare function isAutoSaveSupported(): boolean;

export { AutoSaveManager, formatLastSaveTime, formatStorageSize, getAutoSaveStatusLabel, getAutoSaveStorageSize, isAutoSaveSupported };
