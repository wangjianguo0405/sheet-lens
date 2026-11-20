import { useEffect, useRef, type RefObject } from "react";

import { toggleShapeDebug } from "@/lib/debugPanels";
import type { ActiveCellStyle, CellFormat } from "@/lib/formats";
import type { Tab } from "@/lib/tabs";

export interface ShortcutActions {
  tabs: Tab[];
  activePath: string | null;
  /** The active cell's style, so the format keys can flip it. */
  activeStyle: ActiveCellStyle | null;
  applyFormat: (patch: CellFormat | null) => void;
  toggleFilter: () => Promise<void>;
  handleOpen: () => void;
  closeTab: (path: string) => void;
  selectTab: (path: string) => void;
  reloadTab: (path: string) => void | Promise<void>;
  undo: () => void;
  redo: () => void;
  setGotoOpen: (open: boolean) => void;
  setSearchOpen: (open: boolean) => void;
  /** Clears the value the panel was asked to look for. */
  setSearchSeed: (seed: null) => void;
  /** Brings the keyboard back to the panel's field. */
  setSearchFocus: (update: (n: number) => number) => void;
  /** The search panel's box, so Escape only closes it from inside. */
  searchPanelRef: RefObject<HTMLDivElement | null>;
}

/**
 * The window's shortcuts, all of them in one listener: Excel's keys for the
 * file and the sheet, this app's own for reloading and for undo, and the F5
 * that would otherwise reload the webview.
 *
 * The handlers are read through a ref rather than bound as the effect's
 * dependencies: they are rebuilt with the selection they act on, and the
 * window listener should not have to be taken off and put back on every time
 * the cursor moves.
 */
export function useShortcuts(actions: ShortcutActions) {
  const latest = useRef(actions);
  latest.current = actions;

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const {
        tabs,
        activePath,
        activeStyle,
        applyFormat,
        toggleFilter,
        handleOpen,
        closeTab,
        selectTab,
        reloadTab,
        undo,
        redo,
        setGotoOpen,
        setSearchOpen,
        setSearchSeed,
        setSearchFocus,
        searchPanelRef,
      } = latest.current;
      if (!event.ctrlKey && !event.metaKey) {
        // Excel's F5 opens Go To. Without this the webview would reload the
        // page, which for a viewer means losing every open workbook's view.
        if (event.key === "F5") {
          event.preventDefault();
          if (activePath) setGotoOpen(true);
          return;
        }
        if (event.key === "Escape") {
          // Escape puts the panel away only from the panel's own row — the
          // search field is where it is asked for. Everywhere else the key is
          // the sheet's or a field's: it takes the marquee off a copy, closes
          // a cell's text, and none of that is meant to take the panel with
          // it. Ctrl+F is the shortcut that closes it from anywhere.
          const target = event.target;
          if (target instanceof Node && searchPanelRef.current?.contains(target)) {
            setSearchOpen(false);
          }
        }
        return;
      }
      const key = event.key.toLowerCase();
      // Ctrl+Tab walks the open workbooks, the way it walks Excel's windows;
      // the sheet's own Ctrl+PageUp/Down belong to the grid.
      if (key === "tab") {
        event.preventDefault();
        if (tabs.length < 2) return;
        const at = tabs.findIndex((tab) => tab.info.path === activePath);
        const forward = !event.shiftKey;
        const next =
          at < 0
            ? forward
              ? tabs[0]
              : tabs[tabs.length - 1]
            : tabs[(at + (forward ? 1 : -1) + tabs.length) % tabs.length];
        if (next) selectTab(next.info.path);
        return;
      }
      // Shift makes a different shortcut rather than the same one louder.
      // Without this, Ctrl+Shift+G answered as Ctrl+G and opened the ordinary
      // Go To, and every other letter did the same thing. Only Ctrl+Shift+D has
      // a meaning of its own: the hidden readout, dev builds only, like the
      // panel itself.
      if (event.shiftKey) {
        if (key === "z") {
          const target = event.target;
          if (target instanceof Element && target.closest("input, textarea, [contenteditable]"))
            return;
          event.preventDefault();
          redo();
        } else if (key === "d" && import.meta.env.DEV) {
          event.preventDefault();
          toggleShapeDebug();
        } else if (key === "l") {
          // Excel's filter switch. It is the ribbon's own button: the key
          // belongs to the sheet, not the cursor.
          if (!activePath) return;
          event.preventDefault();
          void toggleFilter();
        }
        return;
      }
      switch (key) {
        case "o":
          event.preventDefault();
          void handleOpen();
          break;
        case "f":
          if (!activePath) return;
          event.preventDefault();
          setSearchOpen(true);
          setSearchSeed(null);
          setSearchFocus((n) => n + 1);
          break;
        case "g":
          if (!activePath) return;
          event.preventDefault();
          setGotoOpen(true);
          break;
        case "w":
          if (!activePath) return;
          event.preventDefault();
          closeTab(activePath);
          break;
        // The webview's own Ctrl+R would reload the page and take every open
        // workbook's view with it. The key reads the file again instead — the
        // tab menu's R — and is held back even with nothing open, so it never
        // lands as a page reload.
        case "r": {
          event.preventDefault();
          if (!activePath) return;
          const target = event.target;
          if (target instanceof Element && target.closest("input, textarea, [contenteditable]"))
            return;
          void reloadTab(activePath);
          break;
        }
        // Ctrl+Z takes back the last session edit; Ctrl+Y puts it back. A field
        // being typed into keeps both keys: that undo is the field's own.
        case "z":
        case "y": {
          const target = event.target;
          if (target instanceof Element && target.closest("input, textarea, [contenteditable]"))
            return;
          event.preventDefault();
          if (key === "z") undo();
          else redo();
          break;
        }
        // Excel's typing keys, applied to what is selected. The formats are
        // the session's own — the ribbon writes the same ones. A field being
        // typed into keeps them: that text is not the sheet's.
        case "b":
        case "i":
        case "u": {
          if (!activePath) return;
          const target = event.target as HTMLElement | null;
          if (target?.closest("input, textarea, [contenteditable]")) return;
          event.preventDefault();
          const cell = activeStyle?.style ?? null;
          if (key === "b") applyFormat({ bold: !(cell?.bold ?? false) });
          else if (key === "i") applyFormat({ italic: !(cell?.italic ?? false) });
          else applyFormat({ underline: cell?.underline ? null : "single" });
          break;
        }
        default:
          break;
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);
}
