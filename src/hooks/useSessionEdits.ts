import { useCallback, useEffect, useState, type Dispatch, type SetStateAction } from "react";

import { currentRegion, type AutoFilter, type SheetInfo } from "@/lib/excel";
import {
  areaRect,
  colorRange,
  formatKey,
  formatSelection,
  type CellFormat,
  type FilterPane,
} from "@/lib/formats";
import { initialFontColor, storeFontColor } from "@/lib/fontColor";
import { initialMarkColor, markSelection, storeMarkColor } from "@/lib/marks";
import type { SessionOutline } from "@/lib/outline";
import { activeArea, areasOf, selectionAnchor, ORIGIN } from "@/lib/selection";
import type { Tab } from "@/lib/tabs";

export interface UseSessionEditsOptions {
  active: Tab | null;
  activeSheet: SheetInfo | null;
  /** What the grid actually shows; the mark and format extents use it. */
  extent: { rows: number; cols: number } | null;
  frozenBands: { rows: number; cols: number };
  setFrozenBands: Dispatch<SetStateAction<{ rows: number; cols: number }>>;
  filterOn: boolean;
  setFilterOn: Dispatch<SetStateAction<boolean>>;
  setError: Dispatch<SetStateAction<string | null>>;
  setTabs: Dispatch<SetStateAction<Tab[]>>;
  /** Takes the undo step an edit is about to need. */
  recordEdit: (path: string, sheet: string) => void;
}

/**
 * Everything the reader lays over a sheet from the ribbon and the cell's menu:
 * marks, fonts and colours, the tab's own colour, and the view decisions the
 * ribbon carries — freeze, split, filters and outline groups. None of it is
 * written back; all of it lasts as long as the tab, and the edits proper take
 * an undo step each.
 */
export function useSessionEdits({
  active,
  activeSheet,
  extent,
  frozenBands,
  setFrozenBands,
  filterOn,
  setFilterOn,
  setError,
  setTabs,
  recordEdit,
}: UseSessionEditsOptions) {
  /** The colour the toolbar's paint bucket is holding, kept between sessions. */
  const [markColor, setMarkColor] = useState(initialMarkColor);
  /** The colour it paints text with, kept the same way and for the same reason. */
  const [fontColor, setFontColor] = useState(initialFontColor);

  useEffect(() => {
    storeMarkColor(markColor);
  }, [markColor]);

  useEffect(() => {
    storeFontColor(fontColor);
  }, [fontColor]);

  /**
   * Shades the current selection in a review colour, or lifts the shading off
   * it when `fill` is null. The marks belong to the tab, so they last only as
   * long as the workbook is open and the file is never touched.
   */
  const applyMark = useCallback(
    (fill: string | null) => {
      if (!active || !activeSheet) return;
      recordEdit(active.info.path, active.sheet);
      // The bucket then holds what was just used, Excel's own behaviour: the
      // next press of the bucket is the same colour without opening anything.
      if (fill !== null) setMarkColor(fill);
      // The extent the grid is drawing, not `activeSheet`'s, which is calamine's
      // range: that spans only cells holding a value, so a column kept alive by
      // a border or a merge falls outside it and the fill stopped short of what
      // is on screen. The grid reports its own extent; before the first report
      // the parser's range is all there is.
      const bounds = extent ?? { rows: activeSheet.rows, cols: activeSheet.cols };
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === active.info.path
            ? {
                ...tab,
                marks: markSelection(
                  tab.marks,
                  tab.sheet,
                  tab.cursors[tab.sheet] ?? ORIGIN,
                  bounds,
                  fill,
                ),
              }
            : tab,
        ),
      );
    },
    [active, activeSheet, extent, setTabs, recordEdit],
  );

  /**
   * Lays a font or alignment choice over the selection, or takes the overlay
   * off it when `patch` is null. It belongs to the tab and never reaches the
   * file, exactly as a mark does.
   */
  const applyFormat = useCallback(
    (patch: CellFormat | null) => {
      if (!active || !activeSheet) return;
      recordEdit(active.info.path, active.sheet);
      const bounds = extent ?? { rows: activeSheet.rows, cols: activeSheet.cols };
      const cursor = active.cursors[active.sheet] ?? ORIGIN;
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === active.info.path
            ? {
                ...tab,
                formats: formatSelection(tab.formats, active.sheet, cursor, bounds, patch),
              }
            : tab,
        ),
      );
    },
    [active, activeSheet, extent, setTabs, recordEdit],
  );

  /**
   * Colours a stretch of one cell's text, or takes the colour back off it when
   * `color` is null. The stretch is character offsets into the value — what
   * the field over the cell reports for whatever it has selected.
   */
  const applyTextColor = useCallback(
    (at: { row: number; col: number; from: number; to: number }, color: string | null) => {
      if (!active) return;
      recordEdit(active.info.path, active.sheet);
      const key = formatKey(active.sheet, at.row, at.col);
      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const next = new Map(tab.textColors);
          const runs = colorRange(next.get(key) ?? [], at.from, at.to, color);
          if (runs.length === 0) next.delete(key);
          else next.set(key, runs);
          return { ...tab, textColors: next };
        }),
      );
    },
    [active, setTabs, recordEdit],
  );

  /**
   * Colours a sheet's tab, or takes the reader's colour off it when `color` is
   * null — which hands the tab back to the colour the workbook gives it, not to
   * no colour at all. Like a mark, it belongs to the tab and the file is never
   * touched.
   */
  const applyTabColor = useCallback(
    (sheet: string, color: string | null) => {
      if (!active) return;
      recordEdit(active.info.path, sheet);
      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const tabColors = new Map(tab.tabColors);
          if (color === null) tabColors.delete(sheet);
          else tabColors.set(sheet, color);
          return { ...tab, tabColors };
        }),
      );
    },
    [active, setTabs, recordEdit],
  );

  /**
   * Freezes the sheet above and to the left of the active cell, or clears
   * whatever freeze is in force — the file's included, which is why clearing
   * is a zero-sized override rather than a return to null. Nothing reaches the
   * workbook; the tab is what remembers it.
   */
  const applyFreeze = useCallback(
    (next: { rows: number; cols: number }) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === active.info.path
            ? { ...tab, frozen: { ...tab.frozen, [tab.sheet]: next } }
            : tab,
        ),
      );
    },
    [active, setTabs],
  );

  /**
   * Splits the sheet at the active cell, or closes the split. Like the freeze
   * this is a view decision kept on the tab and never written to the file.
   */
  const applySplit = useCallback(
    (next: { row: number; col: number } | null) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const splitAt = { ...tab.splitAt };
          if (next) splitAt[tab.sheet] = next;
          else delete splitAt[tab.sheet];
          return { ...tab, splitAt };
        }),
      );
    },
    [active, setTabs],
  );

  const toggleSplit = useCallback(() => {
    if (!active) return;
    if (active.splitAt[active.sheet]) {
      applySplit(null);
      return;
    }
    const at = selectionAnchor(active.cursors[active.sheet] ?? ORIGIN);
    // The lines cross at the active cell's top-left corner. A freeze in the
    // file is released first — it is not where the split goes.
    //
    // A cell in the first row or the first column splits only the other axis,
    // as Excel does: there is nothing above row 1 or left of column A to put in
    // a pane (measured: Split at A5 is SplitRow=4/SplitColumn=0, two panes). At
    // A1 itself neither axis has anywhere to go and Excel halves the window
    // instead, which the grid turns the zero/zero into.
    applySplit({ row: at.row, col: at.col });
    applyFreeze({ rows: 0, cols: 0 });
    setFrozenBands({ rows: 0, cols: 0 });
  }, [active, applySplit, applyFreeze, setFrozenBands]);

  const toggleFreeze = useCallback(() => {
    if (!active) return;
    if (frozenBands.rows > 0 || frozenBands.cols > 0) {
      applyFreeze({ rows: 0, cols: 0 });
      return;
    }
    const at = selectionAnchor(active.cursors[active.sheet] ?? ORIGIN);
    applyFreeze({ rows: at.row, cols: at.col });
    // A frozen pane and a split are two answers to the same question, and
    // Excel lets only one of them stand.
    applySplit(null);
  }, [active, frozenBands, applyFreeze, applySplit]);

  /**
   * Adds a filter over the current region, or takes away whichever filter is
   * in force. Like the freeze, this is a decision kept on the tab.
   */
  const setFilterDecision = useCallback(
    (pane: FilterPane) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === active.info.path
            ? { ...tab, filterPane: { ...tab.filterPane, [tab.sheet]: pane } }
            : tab,
        ),
      );
    },
    [active, setTabs],
  );

  const toggleFilter = useCallback(async () => {
    if (!active) return;
    if (filterOn) {
      setFilterDecision({ area: null, off: true });
      setFilterOn(false);
      return;
    }
    const at = selectionAnchor(active.cursors[active.sheet] ?? ORIGIN);
    try {
      const block = await currentRegion(active.info.path, active.sheet, at.row, at.col);
      const area: AutoFilter = {
        row: block.row,
        col: block.col,
        rows: block.rows,
        cols: block.cols,
        applied: {},
        custom: [],
      };
      setFilterDecision({ area, off: false });
      setFilterOn(true);
    } catch (reason: unknown) {
      setError(String(reason));
    }
  }, [active, filterOn, setFilterDecision, setFilterOn, setError]);

  /**
   * Groups or ungroups the rows or columns the selection covers. Excel asks
   * which axis from the Group dialog; the ribbon's little menu asks instead.
   * Nothing is written back — the grid lays these over the file's own levels.
   */
  const applyOutline = useCallback(
    (action: "rows" | "cols" | "ungroup") => {
      if (!active || !activeSheet) return;
      const bounds = extent ?? { rows: activeSheet.rows, cols: activeSheet.cols };
      const cursor = active.cursors[active.sheet] ?? ORIGIN;
      const rects = areasOf(cursor).map((area) => areaRect(area, bounds));
      if (rects.length === 0) return;
      const rowFrom = Math.min(...rects.map((rect) => rect.row));
      const rowTo = Math.max(...rects.map((rect) => rect.row + rect.rows - 1));
      const colFrom = Math.min(...rects.map((rect) => rect.col));
      const colTo = Math.max(...rects.map((rect) => rect.col + rect.cols - 1));

      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const current: SessionOutline = tab.sessionOutline[tab.sheet] ?? {
            added: { rows: [], cols: [] },
            removed: { rows: [], cols: [] },
          };
          const next: SessionOutline = {
            added: {
              rows: [...current.added.rows],
              cols: [...current.added.cols],
            },
            removed: {
              rows: [...current.removed.rows],
              cols: [...current.removed.cols],
            },
          };
          if (action === "rows") {
            next.added.rows.push([rowFrom, rowTo]);
          } else if (action === "cols") {
            next.added.cols.push([colFrom, colTo]);
          } else {
            // Ungroup follows the shape of what is selected, as Excel does: a
            // run of whole rows is a row group, a run of columns a column one,
            // and anything else is taken as rows — the axis a cell selection
            // groups by default.
            const axis = activeArea(cursor).kind === "column" ? "cols" : "rows";
            const [from, to] = axis === "rows" ? [rowFrom, rowTo] : [colFrom, colTo];
            next.removed[axis].push([from, to]);
            next.added[axis] = next.added[axis].filter(([start, end]) => end < from || start > to);
          }
          return {
            ...tab,
            sessionOutline: { ...tab.sessionOutline, [tab.sheet]: next },
          };
        }),
      );
    },
    [active, activeSheet, extent, setTabs],
  );

  return {
    markColor,
    fontColor,
    setFontColor,
    applyMark,
    applyFormat,
    applyTextColor,
    applyTabColor,
    applyFreeze,
    applySplit,
    toggleSplit,
    toggleFreeze,
    setFilterDecision,
    toggleFilter,
    applyOutline,
  };
}
