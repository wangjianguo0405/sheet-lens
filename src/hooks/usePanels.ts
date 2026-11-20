import { useCallback, useEffect, useRef, useState } from "react";

import { initialFocusCell, storeFocusCell } from "@/lib/focusCell";
import { initialSheetPanels, storeSheetPanels, type SheetPanels } from "@/lib/sheetPanels";
import { readTreeShown, storeTreeShown } from "@/lib/fileTree";
import { readWelcomeColumns, storeWelcomeColumns } from "@/lib/welcomeColumns";

const DEFAULT_SHEET_LIST_W = 240;
const MIN_SHEET_LIST_W = 160;
const MAX_SHEET_LIST_W = 480;

const DEFAULT_TREE_W = 240;
const MIN_TREE_W = 160;
const MAX_TREE_W = 420;

const DEFAULT_SEARCH_W = 288;
const MIN_SEARCH_W = 200;
const MAX_SEARCH_W = 640;

/**
 * The sheet's scrollbar takes a third of the row until it is dragged, and there
 * is no configured ceiling on it afterwards: what it may take is whatever the
 * tab strip beside it can spare, measured during the drag. A fixed one was wrong
 * twice over — on a narrow window the strip's own floor decides first, and the
 * default share can start out *wider* than the number, so the first touch of the
 * handle would snap the bar shorter instead of moving it.
 */
const MIN_SHEET_BAR_W = 120;

/**
 * The things a drag can size: the file tree, the sheet list, the search panel,
 * and the sheet's scrollbar at the end of the tab row.
 */
type DraggablePanel = "fileTree" | "sheetList" | "search" | "sheetBar";

const PANEL_LIMITS: Record<DraggablePanel, [number, number]> = {
  fileTree: [MIN_TREE_W, MAX_TREE_W],
  sheetList: [MIN_SHEET_LIST_W, MAX_SHEET_LIST_W],
  search: [MIN_SEARCH_W, MAX_SEARCH_W],
  sheetBar: [MIN_SHEET_BAR_W, Number.POSITIVE_INFINITY],
};

/**
 * The window's own furniture: which panels and start-page columns show, how
 * wide the draggable ones are, and the drag that resizes them. Everything here
 * that outlives a run is kept in storage by the effects below, so the window
 * opens the way it was left; the widths themselves are not, and neither is a
 * drag in progress.
 */
export function usePanels() {
  /** Which of the two sheet pickers are showing; they are independent. */
  const [panels, setPanels] = useState<SheetPanels>(initialSheetPanels);
  /** Whether the cursor's row and column are washed, kept between sessions. */
  const [focusCell, setFocusCell] = useState(initialFocusCell);
  /** Whether the reader's folder tree sits beside the sheet list, kept between sessions. */
  const [treeShown, setTreeShown] = useState(readTreeShown);
  /** The start page's two lists; the tree beside them is `treeShown`. */
  const [welcomeColumns, setWelcomeColumns] = useState(readWelcomeColumns);
  const [sheetListWidth, setSheetListWidth] = useState(DEFAULT_SHEET_LIST_W);
  const [treeWidth, setTreeWidth] = useState(DEFAULT_TREE_W);
  const [searchWidth, setSearchWidth] = useState(DEFAULT_SEARCH_W);
  const [sheetBarWidth, setSheetBarWidth] = useState<number | null>(null);
  /** The row's own box, so a drag can start from the length it has now. */
  const sheetBarRef = useRef<HTMLDivElement>(null);
  /**
   * Which side panel a drag is sizing, so one set of listeners serves all of
   * them — they differ only in which way the pointer's delta counts, how far
   * they may be dragged, and where the result is stored.
   */
  const [draggingPanel, setDraggingPanel] = useState<DraggablePanel | null>(null);
  /** Where the drag started, so the width follows the pointer's delta. */
  const panelDrag = useRef<{ x: number; width: number } | null>(null);

  /**
   * Each switch flips only its own half. Taking the whole pair from the render
   * this handler belongs to would let two presses before the next render
   * overwrite one another — and the second press would restore what the first
   * one turned off.
   */
  const togglePanel = useCallback((which: keyof SheetPanels) => {
    setPanels((prev) => ({ ...prev, [which]: !prev[which] }));
  }, []);

  useEffect(() => {
    storeSheetPanels(panels);
  }, [panels]);

  useEffect(() => {
    storeTreeShown(treeShown);
  }, [treeShown]);

  useEffect(() => {
    storeWelcomeColumns(welcomeColumns);
  }, [welcomeColumns]);

  useEffect(() => {
    storeFocusCell(focusCell);
  }, [focusCell]);

  const startPanelDrag = useCallback(
    (which: "fileTree" | "sheetList" | "search", event: React.PointerEvent<HTMLDivElement>) => {
      panelDrag.current = {
        x: event.clientX,
        width:
          which === "fileTree" ? treeWidth : which === "sheetList" ? sheetListWidth : searchWidth,
      };
      setDraggingPanel(which);
    },
    [treeWidth, sheetListWidth, searchWidth],
  );

  /**
   * The sheet's scrollbar is a share of the row until it is dragged, so the drag
   * starts from whatever that measures now rather than from a number the row has
   * since outgrown.
   */
  const startSheetBarDrag = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const el = sheetBarRef.current;
    if (!el) return;
    panelDrag.current = {
      x: event.clientX,
      width: Math.round(el.getBoundingClientRect().width),
    };
    setDraggingPanel("sheetBar");
  }, []);

  /**
   * The widest the sheet's scrollbar can be before the tab strip beside it would
   * have to give up more room than it has — which is the only ceiling there is,
   * and so has to be measured rather than written down: the dots come and go,
   * and the floor the tabs keep is the stylesheet's business, not this file's.
   *
   * Asking for more than this is not merely pointless, it is what made the drag
   * go dead: the layout clamps while the number goes on climbing, so the handle
   * stops following the pointer and the first stretch of a drag back does
   * nothing at all.
   */
  const sheetBarRoom = useCallback(() => {
    const bar = sheetBarRef.current;
    const strip = bar?.closest("[data-sheet-strip]")?.querySelector("[data-sheet]")?.parentElement;
    // Neither can be missing while a drag is running — the handle that starts it
    // is in the same row — so the fallback is only here to keep the type honest.
    if (!bar || !strip) return Number.POSITIVE_INFINITY;
    const floor = parseFloat(getComputedStyle(strip).minWidth) || 0;
    const spare = strip.getBoundingClientRect().width - floor;
    return bar.getBoundingClientRect().width + Math.max(0, spare);
  }, []);

  /**
   * The listeners live on the window, not the handle: a drag leaves the four
   * pixels of it immediately, and pointer capture would keep the cursor from
   * reading as a resize once it is outside.
   */
  useEffect(() => {
    if (!draggingPanel) return;
    const [min, max] = PANEL_LIMITS[draggingPanel];
    const resize = {
      fileTree: setTreeWidth,
      sheetList: setSheetListWidth,
      search: setSearchWidth,
      sheetBar: setSheetBarWidth,
    }[draggingPanel];
    // A panel on the left grows as the pointer moves right; the two on the
    // right grow as it moves left.
    const direction = draggingPanel === "sheetList" || draggingPanel === "fileTree" ? 1 : -1;

    const onMove = (event: PointerEvent) => {
      const drag = panelDrag.current;
      if (!drag) return;
      const delta = (event.clientX - drag.x) * direction;
      // The scrollbar's ceiling is whichever comes first: the width the row has
      // to give it, or the one configured for it.
      const ceiling = draggingPanel === "sheetBar" ? Math.min(max, sheetBarRoom()) : max;
      resize(Math.min(ceiling, Math.max(min, drag.width + delta)));
    };
    const onUp = () => setDraggingPanel(null);

    const { cursor, userSelect } = document.body.style;
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
    return () => {
      document.body.style.cursor = cursor;
      document.body.style.userSelect = userSelect;
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
    };
  }, [draggingPanel, sheetBarRoom]);

  return {
    panels,
    togglePanel,
    focusCell,
    setFocusCell,
    treeShown,
    setTreeShown,
    welcomeColumns,
    setWelcomeColumns,
    sheetListWidth,
    treeWidth,
    searchWidth,
    sheetBarWidth,
    sheetBarRef,
    draggingPanel,
    startPanelDrag,
    startSheetBarDrag,
    sheetBarRoom,
  };
}
