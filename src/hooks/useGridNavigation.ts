import { useCallback, type Dispatch, type SetStateAction } from "react";

import type { SearchHit } from "@/lib/excel";
import type { Target } from "@/lib/column";
import type { Selection } from "@/lib/selection";
import type { Tab } from "@/lib/tabs";
import type { ViewMode } from "@/lib/viewMode";

export interface UseGridNavigationOptions {
  active: Tab | null;
  activePath: string | null;
  setTabs: Dispatch<SetStateAction<Tab[]>>;
  updateActive: (patch: Partial<Omit<Tab, "info">>) => void;
  /** Asks the grid to bring the cursor into sight. */
  setReveal: Dispatch<SetStateAction<number>>;
  /** The grid's scroller, so the name box can hand the keyboard back to it. */
  viewport: HTMLDivElement | null;
  setExtent: Dispatch<SetStateAction<{ rows: number; cols: number } | null>>;
}

/**
 * Moving the cursor and the view: the selection the grid reports, where Go To,
 * a search hit and the name box put it, and which view each sheet was left in.
 * Nothing here reaches the file — the cursors and views live on the tab.
 */
export function useGridNavigation({
  active,
  activePath,
  setTabs,
  updateActive,
  setReveal,
  viewport,
  setExtent,
}: UseGridNavigationOptions) {
  /** Remembers where the cursor was, keyed by sheet. */
  const putCursor = useCallback(
    (sheetName: string, next: Selection) => {
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === activePath
            ? { ...tab, cursors: { ...tab.cursors, [sheetName]: next } }
            : tab,
        ),
      );
    },
    [activePath, setTabs],
  );

  const handleSelect = useCallback(
    (next: Selection) => {
      if (!active) return;
      putCursor(active.sheet, next);
    },
    [active, putCursor],
  );

  /**
   * Remembers which view a sheet was switched to. Keyed by sheet, like the
   * cursors: switching away and back returns the sheet to the view it was left
   * in, rather than to the one the workbook opened with.
   */
  const putView = useCallback(
    (sheetName: string, next: ViewMode) => {
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === activePath
            ? { ...tab, views: { ...tab.views, [sheetName]: next } }
            : tab,
        ),
      );
    },
    [activePath, setTabs],
  );

  /** Which of Excel's two views the sheet on screen is being read in. */
  const handleView = useCallback(
    (next: ViewMode) => {
      if (!active) return;
      putView(active.sheet, next);
    },
    [active, putView],
  );

  const handleJump = useCallback(
    (hit: SearchHit) => {
      updateActive({ sheet: hit.sheet });
      putCursor(hit.sheet, { kind: "cell", row: hit.row, col: hit.col });
      setReveal((n) => n + 1);
    },
    [updateActive, putCursor, setReveal],
  );

  const handleGo = useCallback(
    (sheet: string, row: number, col: number) => {
      updateActive({ sheet });
      putCursor(sheet, { kind: "cell", row, col });
      setReveal((n) => n + 1);
    },
    [updateActive, putCursor, setReveal],
  );

  /**
   * A reference typed into the name box goes where Go To goes. A sheet name that
   * is not in the workbook is not an error here — there is no dialog to report
   * it in — so the selection stays where it was and the box goes back to saying
   * where that is.
   */
  const handleNameBoxGo = useCallback(
    (target: Target) => {
      if (!active) return;
      const sheet = target.sheet
        ? active.info.sheets.find((sheet) => sheet.name === target.sheet)?.name
        : active.sheet;
      if (!sheet) return;
      handleGo(sheet, target.row, target.col);
      // Back to the grid, or the arrow keys would be typing into the name box.
      viewport?.focus({ preventScroll: true });
    },
    [active, handleGo, viewport],
  );

  const handleExtent = useCallback(
    (rows: number, cols: number) => {
      setExtent({ rows, cols });
    },
    [setExtent],
  );

  return { handleSelect, handleView, handleJump, handleGo, handleNameBoxGo, handleExtent };
}
