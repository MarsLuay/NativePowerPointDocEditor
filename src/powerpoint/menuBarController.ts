import { Component } from 'obsidian';

import { isNode } from '../domGuards';
import {
	PPTX_EDITOR_CHROME_MENUBAR_CLASS,
} from '../editorChromeRegions';
import { debugLog } from '../logger';
import { createMenuItem, createMenuSection, createPopoverShell, positionPopoverBelow } from '../menuControls';
import type { MenuDropdownEntry } from './types';

export type MenuBarTab =
  | { kind: 'dropdown'; label: string; getItems: () => MenuDropdownEntry[] }
  | {
    kind: 'search';
    label: string;
    inputAriaLabel: string;
    placeholder: string;
    emptyLabel: string;
    getItems: () => MenuSearchEntry[];
  }
  | { kind: 'action'; label: string; action: () => void };

export interface MenuSearchEntry {
  id: string;
  label: string;
  keywords?: readonly string[];
  icon?: string;
  disabled?: boolean;
  onClick: () => void;
}

/**
 * The Google-style menu bar: a row of tabs where "dropdown" tabs reveal a
 * floating list of {@link MenuDropdownEntry} items on click/hover, and "action"
 * tabs fire a callback. Extracted verbatim from `NativePowerPointView`.
 *
 * It is parameterized by callbacks ({@link MenuBarTab}), so it owns only the
 * open/close state machine and is decoupled from the actual menu contents and
 * commands. Registered as a child `Component` of the view so its global
 * listeners are cleaned up with the view.
 */
export class MenuBarController extends Component {
  private activeTab: HTMLElement | null = null;
  private activeDropdown: HTMLElement | null = null;
  private closeTimer: number | null = null;

  build(root: HTMLElement, tabs: MenuBarTab[]): void {
    const bar = root.createDiv({ cls: PPTX_EDITOR_CHROME_MENUBAR_CLASS });
    for (const tab of tabs) {
      if (tab.kind === 'dropdown') {
        this.createDropdownTab(bar, tab.label, tab.getItems);
      } else if (tab.kind === 'search') {
        this.createSearchTab(bar, tab);
      } else {
        this.createActionTab(bar, tab.label, tab.action);
      }
    }

    this.registerDomEvent(
      activeDocument,
      'pointerdown',
      (event) => {
        const target = isNode(event.target) ? event.target : null;
        if (target && bar.contains(target)) return;
        if (target && this.activeDropdown?.contains(target)) return;
        this.closeDropdown();
      },
      true
    );
    this.register(() => this.closeDropdown());
  }

  private createDropdownTab(
    bar: HTMLElement,
    label: string,
    getItems: () => MenuDropdownEntry[]
  ): HTMLButtonElement {
    const button = bar.createEl('button', {
      cls: 'native-powerpoint-menubar-item',
      text: label
    });
    button.type = 'button';
    button.addEventListener('click', () => {
      if (this.activeTab === button) {
        this.closeDropdown();
      } else {
        this.openDropdown(button, getItems);
      }
    });
    // Google-style menu bar: hovering a tab reveals its options.
    button.addEventListener('mouseenter', () => this.openDropdown(button, getItems));
    button.addEventListener('mouseleave', () => this.scheduleClose());
    return button;
  }

  private createSearchTab(
    bar: HTMLElement,
    tab: Extract<MenuBarTab, { kind: 'search' }>,
  ): HTMLButtonElement {
    const button = bar.createEl('button', {
      cls: 'native-powerpoint-menubar-item',
      text: tab.label,
    });
    button.type = 'button';
    button.setAttribute('aria-label', tab.inputAriaLabel);
    button.addEventListener('click', () => {
      if (this.activeTab === button) {
        this.closeDropdown();
      } else {
        this.openSearch(button, tab);
      }
    });
    button.addEventListener('mouseenter', () => this.closeDropdown());
    return button;
  }

  private createActionTab(bar: HTMLElement, label: string, action: () => void): HTMLButtonElement {
    const button = bar.createEl('button', {
      cls: 'native-powerpoint-menubar-item',
      text: label
    });
    button.type = 'button';
    button.addEventListener('click', () => {
      this.closeDropdown();
      debugLog('menu', 'Dispatching PowerPoint menu command', { commandId: label });
      action();
    });
    // Moving onto a no-dropdown tab dismisses any open menu, like Google's bar.
    button.addEventListener('mouseenter', () => this.closeDropdown());
    return button;
  }

  private openDropdown(tab: HTMLElement, getItems: () => MenuDropdownEntry[]): void {
    this.cancelCloseTimer();
    if (this.activeTab === tab && this.activeDropdown) return;
    this.closeDropdown();

    const dropdown = createPopoverShell(activeDocument.body, {
      className: 'native-powerpoint-menubar-dropdown native-powerpoint-light-surface'
    });
    for (const entry of getItems()) {
      if (entry === 'separator') {
        createMenuSection(dropdown, {
          className: 'native-powerpoint-menubar-dropdown-sep',
          role: 'separator'
        });
        continue;
      }
      createMenuItem(dropdown, {
        className: 'native-powerpoint-menubar-dropdown-item',
        text: entry.label,
        icon: entry.icon ? { name: entry.icon, className: 'native-powerpoint-menubar-dropdown-icon' } : undefined,
        labelClassName: 'native-powerpoint-menubar-dropdown-label',
        disabled: entry.disabled,
        onClick: entry.disabled ? undefined : () => {
          this.closeDropdown();
          debugLog('menu', 'Dispatching PowerPoint menu command', { commandId: entry.label });
          entry.onClick();
        }
      });
    }

    dropdown.addEventListener('mouseenter', () => this.cancelCloseTimer());
    dropdown.addEventListener('mouseleave', () => this.scheduleClose());

    positionPopoverBelow(dropdown, tab, 2);

    tab.addClass('is-active');
    this.activeTab = tab;
    this.activeDropdown = dropdown;
  }

  private openSearch(
    tab: HTMLElement,
    searchTab: Extract<MenuBarTab, { kind: 'search' }>,
  ): void {
    this.cancelCloseTimer();
    this.closeDropdown();

    const dropdown = createPopoverShell(activeDocument.body, {
      className: 'native-powerpoint-menubar-search-menu native-powerpoint-light-surface',
      role: 'dialog',
    });
    const input = dropdown.createEl('input', {
      cls: 'native-powerpoint-menubar-search-input',
      type: 'search',
      attr: {
        'aria-label': searchTab.inputAriaLabel,
        placeholder: searchTab.placeholder,
      },
    });
    const results = dropdown.createDiv({ cls: 'native-powerpoint-menubar-search-results', attr: { role: 'listbox' } });
    let activeIndex = 0;

    const getMatches = (): MenuSearchEntry[] => {
      const query = input.value.trim().toLocaleLowerCase();
      const entries = searchTab.getItems();
      if (!query) return entries;
      return entries.filter((entry) => [entry.label, ...(entry.keywords ?? [])]
        .join(' ')
        .toLocaleLowerCase()
        .includes(query));
    };

    const choose = (entry: MenuSearchEntry): void => {
      this.closeDropdown();
      debugLog('search', 'Dispatching PowerPoint option-search command', { commandId: entry.id });
      entry.onClick();
    };

    const renderResults = (): void => {
      const matches = getMatches();
      activeIndex = Math.max(0, Math.min(activeIndex, matches.length - 1));
      results.empty();

      if (matches.length === 0) {
        results.createDiv({ cls: 'native-powerpoint-menubar-search-empty', text: searchTab.emptyLabel });
        return;
      }

      matches.forEach((entry, index) => {
        createMenuItem(results, {
          className: 'native-powerpoint-menubar-dropdown-item native-powerpoint-menubar-search-result',
          text: entry.label,
          icon: entry.icon ? { name: entry.icon, className: 'native-powerpoint-menubar-dropdown-icon' } : undefined,
          labelClassName: 'native-powerpoint-menubar-dropdown-label',
          role: 'option',
          selected: index === activeIndex,
          disabled: entry.disabled,
          onMouseEnter: () => {
            activeIndex = index;
            renderResults();
          },
          onClick: entry.disabled ? undefined : () => choose(entry),
        });
      });
    };

    input.addEventListener('input', () => {
      activeIndex = 0;
      renderResults();
    });
    input.addEventListener('keydown', (event) => {
      const matches = getMatches();
      if (event.key === 'Escape') {
        event.preventDefault();
        this.closeDropdown();
      } else if (event.key === 'ArrowDown') {
        event.preventDefault();
        activeIndex = matches.length > 0 ? (activeIndex + 1) % matches.length : 0;
        renderResults();
      } else if (event.key === 'ArrowUp') {
        event.preventDefault();
        activeIndex = matches.length > 0 ? (activeIndex - 1 + matches.length) % matches.length : 0;
        renderResults();
      } else if (event.key === 'Enter') {
        event.preventDefault();
        const entry = matches[activeIndex];
        if (entry && !entry.disabled) choose(entry);
      }
    });

    positionPopoverBelow(dropdown, tab, 2);
    tab.addClass('is-active');
    this.activeTab = tab;
    this.activeDropdown = dropdown;
    debugLog('search', 'PowerPoint option search opened', { op: 'open', resultCount: getMatches().length });
    renderResults();
    window.setTimeout(() => input.focus());
  }

  private scheduleClose(): void {
    this.cancelCloseTimer();
    this.closeTimer = window.setTimeout(() => {
      this.closeTimer = null;
      this.closeDropdown();
    }, 160);
  }

  private cancelCloseTimer(): void {
    if (this.closeTimer !== null) {
      window.clearTimeout(this.closeTimer);
      this.closeTimer = null;
    }
  }

  private closeDropdown(): void {
    this.cancelCloseTimer();
    this.activeDropdown?.remove();
    this.activeDropdown = null;
    this.activeTab?.removeClass('is-active');
    this.activeTab = null;
  }
}
