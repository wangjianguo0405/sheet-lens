import { useCallback, useRef, useState } from "react";

/**
 * The search panel's own little state: whether it shows, the value it was asked
 * to look for, and the signal that brings the keyboard back to its field.
 */
export function useSearchPanel() {
  const [open, setOpen] = useState(false);
  /**
   * A value the cell's menu asked the search panel to look for. The signal is
   * bumped on every ask, so asking for the same text twice still runs it.
   */
  const [seed, setSeed] = useState<{ text: string; signal: number } | null>(null);
  /**
   * Bumped by Ctrl+F. The panel focuses its input when it opens; this is what
   * brings the keyboard back to it when it is already open, where pressing the
   * shortcut used to put the whole panel away.
   */
  const [focus, setFocus] = useState(0);
  /** The panel's box, so Escape can tell whose key it was. */
  const panelRef = useRef<HTMLDivElement>(null);

  /** Looks for a value the cell's menu asked this panel to find. */
  const searchFor = useCallback((text: string) => {
    if (!text) return;
    setOpen(true);
    setSeed((prev) => ({ text, signal: (prev?.signal ?? 0) + 1 }));
  }, []);

  return {
    searchOpen: open,
    setSearchOpen: setOpen,
    searchSeed: seed,
    setSearchSeed: setSeed,
    searchFocus: focus,
    setSearchFocus: setFocus,
    searchFor,
    searchPanelRef: panelRef,
  };
}
