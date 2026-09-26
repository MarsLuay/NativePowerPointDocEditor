import type { TranslateFn } from '../i18n/translate';
import type { SaveState } from '../powerpoint/types';

export type { SaveState } from '../powerpoint/types';

/**
 * Runtime-agnostic semantics for the save-status vocabulary shared by the DOCX
 * (React) and PPTX (imperative) editors. `busy` states show a spinner and are
 * aria-busy; `interactive` states expose a click/keyboard retry affordance.
 */
export interface SaveStatusFlags {
  busy: boolean;
  interactive: boolean;
}

export function getSaveStatusFlags(state: SaveState): SaveStatusFlags {
  return {
    busy: state === 'saving',
    interactive: state === 'failed',
  };
}

const SAVE_STATE_LABEL_KEYS: Record<SaveState, string> = {
  idle: 'powerpoint:save.ready',
  dirty: 'powerpoint:save.unsaved',
  saving: 'powerpoint:save.saving',
  saved: 'powerpoint:save.saved',
  failed: 'powerpoint:save.failed',
  recovered: 'powerpoint:save.recovered',
  'view-only': 'powerpoint:save.viewOnly',
};

export function getSaveStatusLabel(t: TranslateFn, state: SaveState): string {
  return t(SAVE_STATE_LABEL_KEYS[state]);
}

export interface ApplySaveStatusOptions {
  state: SaveState;
  label: string;
  failedAriaLabel?: string;
  failedTitle?: string;
}

/**
 * Applies the shared status-presentation contract to an imperative status
 * element: `data-state`, `aria-busy`, and the failed retry affordance
 * (`is-clickable` + button role/tabindex/label/title). PPTX styles this element
 * via `[data-state]` and `.is-clickable`.
 */
export function applySaveStatusPresentation(el: HTMLElement, options: ApplySaveStatusOptions): void {
  const { state, label } = options;
  const { busy, interactive } = getSaveStatusFlags(state);

  el.setText(label);
  el.dataset.state = state;
  el.setAttribute('aria-busy', busy ? 'true' : 'false');
  el.toggleClass('is-clickable', interactive);

  if (interactive) {
    el.setAttribute('role', 'button');
    el.setAttribute('tabindex', '0');
    if (options.failedAriaLabel) {
      el.setAttribute('aria-label', options.failedAriaLabel);
    } else {
      el.removeAttribute('aria-label');
    }
    if (options.failedTitle) {
      el.title = options.failedTitle;
    } else {
      el.removeAttribute('title');
    }
  } else {
    el.setAttribute('role', 'status');
    el.removeAttribute('tabindex');
    el.removeAttribute('aria-label');
    el.removeAttribute('title');
  }
}

/**
 * DOCX tracks a reduced status vocabulary in local React state. Map it onto the
 * canonical {@link SaveState} so both editors share one set of semantics.
 */
export type DocxSaveStatus = 'saved' | 'saving' | 'unsaved' | 'failed';

export const DOCX_SAVE_STATUS_TO_STATE: Record<DocxSaveStatus, SaveState> = {
  saved: 'saved',
  saving: 'saving',
  unsaved: 'dirty',
  failed: 'failed',
};
