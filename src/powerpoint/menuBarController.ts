import { Component, setIcon } from 'obsidian';

import { isNode } from '../domGuards';
import type { MenuDropdownEntry } from './types';

export type MenuBarTab =
  | { kind: 'dropdown'; label: string; getItems: () => MenuDropdownEntry[] }
  | { kind: 'action'; label: string; action: () => void };

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
    const bar = root.createDiv({ cls: 'native-powerpoint-menubar' });
    for (const tab of tabs) {
      if (tab.kind === 'dropdown') {
        this.createDropdownTab(bar, tab.label, tab.getItems);
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

  private createActionTab(bar: HTMLElement, label: string, action: () => void): HTMLButtonElement {
    const button = bar.createEl('button', {
      cls: 'native-powerpoint-menubar-item',
      text: label
    });
    button.type = 'button';
    button.addEventListener('click', () => {
      this.closeDropdown();
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

    const dropdown = activeDocument.body.createDiv({
      cls: 'native-powerpoint-menubar-dropdown native-powerpoint-light-surface'
    });
    for (const entry of getItems()) {
      if (entry === 'separator') {
        dropdown.createDiv({ cls: 'native-powerpoint-menubar-dropdown-sep' });
        continue;
      }
      const item = dropdown.createEl('button', { cls: 'native-powerpoint-menubar-dropdown-item' });
      item.type = 'button';
      if (entry.icon) {
        setIcon(item.createSpan({ cls: 'native-powerpoint-menubar-dropdown-icon' }), entry.icon);
      }
      item.createSpan({ cls: 'native-powerpoint-menubar-dropdown-label', text: entry.label });
      if (entry.disabled) {
        item.disabled = true;
      } else {
        item.addEventListener('click', () => {
          this.closeDropdown();
          entry.onClick();
        });
      }
    }

    dropdown.addEventListener('mouseenter', () => this.cancelCloseTimer());
    dropdown.addEventListener('mouseleave', () => this.scheduleClose());

    const rect = tab.getBoundingClientRect();
    dropdown.style.left = `${rect.left}px`;
    dropdown.style.top = `${rect.bottom + 2}px`;

    tab.addClass('is-active');
    this.activeTab = tab;
    this.activeDropdown = dropdown;
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
