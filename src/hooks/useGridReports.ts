import { useState } from "react";

import type { Cell } from "@/lib/excel";
import type { ActiveCellStyle } from "@/lib/formats";

/**
 * What the grid reports about itself for everything outside it to draw: the
 * cursor's cell, the style the ribbon shows for it, the freeze and filter in
 * force, the extent on screen, and how many rows the filter is hiding. The
 * grid owns all of it; these are its answers, kept here so the toolbar, the
 * ribbon and the status bar read the same ones.
 */
export function useGridReports() {
  const [detail, setDetail] = useState<Cell | null>(null);
  /** The active cell's style as the ribbon shows it, overlays included. */
  const [activeStyle, setActiveStyle] = useState<ActiveCellStyle | null>(null);
  /** The frozen bands in force, and whether a filter is, as the grid reports them. */
  const [frozenBands, setFrozenBands] = useState({ rows: 0, cols: 0 });
  const [filterOn, setFilterOn] = useState(false);
  /** What the grid actually shows, which can exceed calamine's range. */
  const [extent, setExtent] = useState<{ rows: number; cols: number } | null>(null);
  /**
   * How many rows the active sheet's filter is hiding, or null when there is no
   * filter. The grid owns the filter, so it reports this up for the status bar
   * to say — the same footing the extent is on.
   */
  const [filtered, setFiltered] = useState<{
    shown: number;
    total: number;
  } | null>(null);

  return {
    detail,
    setDetail,
    activeStyle,
    setActiveStyle,
    frozenBands,
    setFrozenBands,
    filterOn,
    setFilterOn,
    extent,
    setExtent,
    filtered,
    setFiltered,
  };
}
