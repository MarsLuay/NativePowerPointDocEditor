import * as React from 'react';

interface TooltipProps {
  content: React.ReactNode;
  children: React.ReactElement;
  side?: 'top' | 'bottom' | 'left' | 'right';
  delayMs?: number;
}

/**
 * Pass-through stub. Obsidian host owns hover labels (and suppresses chrome
 * tooltips on the formatting rail). Re-enabling floating tips here fought
 * that policy and resurfaced strings like "Insert link (Ctrl+K)".
 */
export function Tooltip({ children }: TooltipProps) {
  return children;
}
