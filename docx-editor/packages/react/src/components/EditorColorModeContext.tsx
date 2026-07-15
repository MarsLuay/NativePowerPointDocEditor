/**
 * Editor chrome light/dark flag for portaled UI (Radix Select, etc.).
 *
 * Portals render outside `.docx-editor-root.dark`, so they must not snapshot
 * theme via `document.querySelector` during render — that reads the previous
 * commit and leaves empty portal shells stuck on `.dark` until remount.
 */

import * as React from 'react';

const EditorIsDarkContext = React.createContext<boolean | null>(null);

export function EditorIsDarkProvider({
  isDark,
  children,
}: {
  isDark: boolean;
  children: React.ReactNode;
}) {
  return <EditorIsDarkContext.Provider value={isDark}>{children}</EditorIsDarkContext.Provider>;
}

/** Prefer context from DocxEditorShell; null means no provider (standalone/tests). */
export function useEditorIsDark(): boolean | null {
  return React.useContext(EditorIsDarkContext);
}
