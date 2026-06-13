import { Component } from 'obsidian';

import { isElement, isNode } from '../domGuards';
import { getToolbarTooltipText, positionToolbarTooltip } from './textUtils';

/**
 * Hover tooltips for toolbar / text-toolbar / find-panel buttons. A small,
 * self-contained state machine: it watches pointer movement within a root
 * element and shows a delayed tooltip for any button carrying tooltip text.
 * Extracted verbatim from `NativePowerPointView`.
 *
 * Registered as a child `Component` of the view so its global listeners are
 * cleaned up with the view.
 */
export class ToolbarTooltipController extends Component {
  attach(root: HTMLElement): void {
    const TOOLBAR_TOOLTIP_DELAY_MS = 450;
    let activeTarget: HTMLElement | null = null;
    let tooltipEl: HTMLDivElement | null = null;
    let tooltipTimer: number | null = null;

    const clearTooltipTimer = (): void => {
      if (tooltipTimer !== null) {
        window.clearTimeout(tooltipTimer);
        tooltipTimer = null;
      }
    };

    const removeTooltip = (): void => {
      tooltipEl?.remove();
      tooltipEl = null;
    };

    const hideTooltip = (): void => {
      clearTooltipTimer();
      removeTooltip();
      activeTarget = null;
    };

    const getTooltipTarget = (target: EventTarget | null): HTMLElement | null => {
      if (!isNode(target) || !isElement(target)) return null;
      const candidate = target.closest<HTMLElement>(
        '.native-powerpoint-toolbar button, .native-powerpoint-text-toolbar button, .native-powerpoint-find-panel button, .native-powerpoint-rotate-handle'
      );
      if (!candidate || !root.contains(candidate)) return null;
      return candidate;
    };

    const showTooltip = (target: HTMLElement): void => {
      const label = getToolbarTooltipText(target);
      if (!label || !target.isConnected) return;

      removeTooltip();
      const tooltip = activeDocument.body.createDiv({ cls: 'docxidian-toolbar-tooltip', text: label });
      tooltipEl = tooltip;
      positionToolbarTooltip(target, tooltip);
    };

    const scheduleTooltip = (target: HTMLElement): void => {
      if (target === activeTarget) return;
      hideTooltip();
      activeTarget = target;
      tooltipTimer = window.setTimeout(() => {
        tooltipTimer = null;
        if (activeTarget === target) {
          showTooltip(target);
        }
      }, TOOLBAR_TOOLTIP_DELAY_MS);
    };

    const handlePointerOver = (event: PointerEvent): void => {
      const target = getTooltipTarget(event.target);
      if (target) {
        scheduleTooltip(target);
      }
    };

    const handlePointerOut = (event: PointerEvent): void => {
      if (!activeTarget || (isNode(event.relatedTarget) && activeTarget.contains(event.relatedTarget))) {
        return;
      }
      hideTooltip();
    };

    this.registerDomEvent(root, 'pointerover', handlePointerOver);
    this.registerDomEvent(root, 'pointerout', handlePointerOut);
    this.registerDomEvent(window, 'scroll', hideTooltip, true);
    this.registerDomEvent(window, 'resize', hideTooltip);
    this.register(hideTooltip);
  }
}
