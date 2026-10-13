import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type ReactNode,
} from "react";
import { ListFilter } from "lucide-react";

import { Button } from "@/components/ui/button";
import { FilterButton } from "@/components/FilterButton";
import { ShapeLayer, type ShapeDebugFacts } from "@/components/ShapeLayer";
import { canvasTheme } from "@/lib/canvasTheme";
import { paintBase, paintOverlay, type GridView, type PaintStats } from "@/lib/canvasPaint";
import {
  cellFontStack,
  clearTextCaches,
  fontString,
  layoutLines,
  textWidth,
} from "@/lib/canvasText";
import { copyText } from "@/lib/clipboard";
import { cellRef } from "@/lib/column";
import { shapeDebugEnabled, subscribeShapeDebug } from "@/lib/debugPanels";
import {
  axisValues,
  currentRegion,
  edgeCell,
  filteredRows,
  getCells,
  sheetLayout,
  type AutoFilter,
  type BorderEdge,
  type Cell,
  type CellStyle,
  type FilterCriterion,
  type MergeRect,
  type SheetInfo,
  type SheetLayout,
} from "@/lib/excel";
import {
  BLOCK_COLS,
  BLOCK_ROWS,
  CELL_PAD_LEFT,
  CELL_PAD_RIGHT,
  CJK_FALLBACK,
  COL_HEADER_H,
  LINE_HEIGHT,
  MAX_COL_W,
  MIN_COL_W,
  MIN_ROW_H,
  OVERSCAN_COLS,
  OVERSCAN_ROWS,
  PRINT_AREA_INSIDE,
  PRINT_AREA_INK,
  PRINT_AREA_LINE,
  RESIZE_HANDLE,
  ROW_HEADER_W,
  buildMetrics,
  cellKey,
  charsToPx,
  digitWidthPx,
  estimateWidth,
  indexAt,
  spanAt,
  styleOfColumn,
  styleOfRow,
  toFilters,
  toOverrides,
} from "@/lib/gridGeometry";
import { stepZoom } from "@/lib/zoom";
import {
  colsBetween,
  rangeBetween,
  rangeOver,
  rowsBetween,
  selectionAnchor,
  type Selection,
} from "@/lib/selection";

/** Where a pointer landed: a cell, or one of the two gutter strips. */
type Hit =
  | { kind: "cell"; row: number; col: number }
  | { kind: "rowHeader"; row: number }
  | { kind: "colHeader"; col: number }
  | null;

/** What a drag-select is anchored on, and whether it has moved off it yet. */
type DragAnchor =
  | { kind: "cell"; row: number; col: number; dragged: boolean }
  | { kind: "row"; row: number; dragged: boolean }
  | { kind: "column"; col: number; dragged: boolean };

/** Where a sheet's view sits, in sheet-space pixels at 100% zoom. */
export interface SheetView {
  top: number;
  left: number;
  /**
   * The filter this sheet was left with, by column. Excel holds one per sheet
   * for as long as the workbook is open, and the column's button goes on
   * showing it, so it belongs here with the scroll position rather than with
   * the rows it happens to hide. Absent means nobody has filtered this sheet
   * yet, which is not the same as having cleared it.
   */
  filters?: Map<number, string[]>;
}

/** Those positions, by workbook path and then by sheet name. */
export type SheetViews = Map<string, Map<string, SheetView>>;

interface GridProps {
  path: string;
  sheet: SheetInfo;
  selection: Selection;
  /**
   * Where each sheet of each open workbook was left scrolled, keyed by
   * workbook and then by sheet. The app owns it rather than the grid, so that
   * closing a tab forgets its sheets' positions — the same way it forgets
   * their cursors.
   */
  views: SheetViews;
  onSelect: (selection: Selection) => void;
  onDetail?: (cell: Cell | null) => void;
  /** The extent actually on screen, which can exceed calamine's range. */
  onExtent?: (rows: number, cols: number) => void;
  /** Steps to the sheet `delta` along in the workbook, for Ctrl+PageUp/Down. */
  onSheetStep?: (delta: number) => void;
  /**
   * Bumped when the caller places the cursor somewhere it has to be *seen* — Go
   * To, or a search hit. Arriving on a sheet is not one of those: Excel puts
   * back the view that sheet was left at, off-screen cursor and all, so a jump
   * is the only thing allowed to drag it.
   */
  reveal?: number;
  /** View scale as a percentage, the way Excel names it. */
  zoom: number;
  onZoom: (percent: number) => void;
  /**
   * Sits between the filter banner and the grid. Excel stacks them filter →
   * formula bar → cells, so the bar cannot live above the whole view.
   */
  bar?: ReactNode;
  /**
   * Hides the scroller's own horizontal scrollbar, for a caller that draws the
   * sheet's scrollbar itself — level with the sheet tabs, where Excel has it.
   * The sheet gets one scrollbar either way, never two.
   */
  hideHScroll?: boolean;
  /**
   * The element the sheet scrolls within, handed out so a scrollbar drawn
   * elsewhere is a view of it rather than a second scroller.
   */
  onViewport?: (element: HTMLDivElement | null) => void;
}

interface ResizeDrag {
  kind: "col" | "row";
  index: number;
  start: number;
  size: number;
}

async function computeVisibleRows(
  path: string,
  sheet: string,
  filter: AutoFilter | null,
  filters: Map<number, string[]>,
): Promise<number[] | null> {
  if (!filter || filters.size === 0) return null;
  const criteria: FilterCriterion[] = [...filters].map(([col, values]) => ({
    col,
    values,
  }));
  return filteredRows(
    path,
    sheet,
    filter.row + 1,
    filter.row + filter.rows - 1,
    criteria,
  );
}

/**
 * A canvas-rendered spreadsheet grid.
 *
 * The DOM grid this replaces drew a div per visible cell and relied on React to
 * reconcile them; moving the selection by one cell rebuilt about a thousand
 * elements, which is what made clicking a cell lag. Here the cells are painted,
 * on two stacked canvases:
 *
 *   base    the cells, their text, the gridlines and the gutter bands — repaint
 *           only when the scroll, the zoom, the size or the data changes.
 *   overlay the selection, and the gutter labels — repaint on a selection
 *           change, which is the common interaction, without touching the layer
 *           that holds every string on screen.
 *
 * The scroller stays a real `overflow: auto` box with a spacer inside it, which
 * keeps native scrollbars, native wheel physics and the existing hit-testing
 * arithmetic; only its content is canvas instead of elements.
 */
/**
 * A rolling window of base passes.
 *
 * One frame's `performance.now()` swings by a quarter either way — a dev build,
 * a JIT that has not warmed up, a GC that fired mid-pass — so a single sample
 * says nothing about whether a change helped. The readout reports the window's
 * median instead, which is stable enough to compare two builds by.
 */
interface PaintTimings {
  overlayMs: number;
  layers: number;
  /** The last base pass's own counts. An overlay-only frame keeps them. */
  stats: PaintStats | null;
  base: number[];
  surface: number[];
  text: number[];
  /** Cells the pass walked. Recorded alongside the timings so a median cell
   *  count can be quoted with the median time — the last pass's own count can
   *  belong to a different scroll position than the window's timings do. */
  cells: number[];
}

const PERF_SAMPLES = 60;

/**
 * The copy marquee's crawl: one dash-width every so often, which is about the
 * speed Excel's own ants move at. Slow enough that the overlay pass is not
 * repainted for nothing, quick enough to read as movement rather than a dash.
 */
const ANTS_MS = 60;
const ANTS_STEP = 2;

/**
 * The blank rows and columns the scroller holds past the sheet's own extent.
 *
 * Excel's sheets are 1,048,576 rows by 16,384 columns and its view runs on well
 * past anything the file holds — a grid that stopped dead at the last cell reads
 * as broken beside it, and there is nowhere to put a note beside the data.
 *
 * It is grown as the view reaches the end and given back when the view comes
 * away from it, rather than laid out once at the sheet's real size: a million
 * rows of offsets is 8 MB of `Float64Array` rebuilt on every sheet switch, and
 * there is nothing to see in the empty millionth row. The floor is what makes
 * the first wheel past the data do anything at all on a sheet that fits its
 * window.
 *
 * A top-up adds half of what the view shows and the giving-back leaves a whole
 * screen behind — both measured in the sheet's own row height and column width,
 * since one screen of a narrow-columned sheet is a very different count from one
 * screen of a wide one. `VOID_MIN_ROWS` / `VOID_MIN_COLS` are the floor, there
 * so that a sheet shorter than its window can be scrolled at all.
 */
const VOID_MIN_ROWS = 64;
const VOID_MIN_COLS = 16;

function record(samples: number[], value: number): void {
  samples.push(value);
  if (samples.length > PERF_SAMPLES) samples.shift();
}

/** `fraction` of 0 gives the minimum, 0.5 the median, 1 the maximum. */
function percentile(samples: number[], fraction: number): number {
  if (samples.length === 0) return 0;
  const sorted = [...samples].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.floor(fraction * sorted.length))];
}

function describePaint(timing: PaintTimings): string {
  if (timing.layers === 0) return "nothing painted yet";

  const lines = [
    // An overlay-only frame never runs the base pass, so it has no base time of
    // its own. Saying so matters: the 0.0ms reads as "free" otherwise.
    timing.layers === 2
      ? `paint 2 layers — overlay ${timing.overlayMs.toFixed(1)}ms`
      : `paint overlay only — ${timing.overlayMs.toFixed(1)}ms; the base pass did not run this frame`,
  ];

  // The window is reported whatever the last frame was. A click paints only the
  // overlay, and a click is exactly when this is recomputed — skipping the
  // window there hid the scroll that had just been measured.
  const n = timing.base.length;
  if (n === 0) {
    lines.push("  no base pass yet — scroll the sheet, then read again");
    return lines.join("\n");
  }

  const { stats } = timing;
  lines.push(
    `  base, last ${n} pass${n === 1 ? "" : "es"}: median ${percentile(timing.base, 0.5).toFixed(1)}ms` +
      `   min ${percentile(timing.base, 0).toFixed(1)}ms` +
      `   p95 ${percentile(timing.base, 0.95).toFixed(1)}ms`,
    // Medians here rather than the last pass's split: one frame's split swings
    // as much as the total does.
    `  median split: surface ${percentile(timing.surface, 0.5).toFixed(1)}ms / text ${percentile(timing.text, 0.5).toFixed(1)}ms` +
      `   over a median ${percentile(timing.cells, 0.5)} cells walked`,
  );
  if (stats) {
    const cells = stats.rows * stats.cols;
    lines.push(
      `  the pass just measured: ${stats.rows} rows × ${stats.cols} cols = ${cells} cells, ` +
        `${stats.texts} of them holding text; ${stats.spans} merges and runs drawn as one box`,
      `  its style lookups: ${stats.styleResolved} resolved for real, ${stats.styleMemoHits} from the memo` +
        ` — ${((stats.styleResolved + stats.styleMemoHits) / Math.max(1, cells)).toFixed(1)} per cell`,
      // The two halves of the surface pass, side by side: how many rectangles it
      // drew, and how long the whole pass was. Whichever of the two the time
      // tracks is the one worth attacking.
      `  its drawing: ${stats.surfaceFills} rectangles filled (${(stats.surfaceFills / Math.max(1, cells)).toFixed(1)} per cell)`,
    );
  }
  return lines.join("\n");
}

export function CanvasGrid({
  path,
  sheet,
  selection,
  views,
  onSelect,
  onDetail,
  onExtent,
  onSheetStep,
  reveal,
  zoom,
  onZoom,
  bar,
  hideHScroll,
  onViewport,
}: GridProps) {
  const viewportRef = useRef<HTMLDivElement>(null);
  /**
   * How tall the scroller's own horizontal scrollbar is, and 0 when that bar is
   * wanted. A bar cannot be hidden one axis at a time — `overflow-x: hidden`
   * would stop the wheel scrolling the sheet sideways, and a 0-height
   * `::-webkit-scrollbar` takes the vertical bar with it — so when the caller
   * draws the sheet's scrollbar itself, the scroller is made exactly that much
   * taller than the box that clips it and its own bar falls outside. Nothing
   * else about the element changes: same scrolling, same measurements.
   */
  const [hBarHeight, setHBarHeight] = useState(0);
  const baseRef = useRef<HTMLCanvasElement>(null);
  const overlayRef = useRef<HTMLCanvasElement>(null);
  /** The clipped strip the filter buttons live in, over the cell area. */
  const overlayAreaRef = useRef<HTMLDivElement>(null);
  /** One positioned wrapper per filter-header column, keyed by column index. */
  const filterSlots = useRef(new Map<number, HTMLDivElement>());
  /** The clipped cell area the shapes and the print mark live in. */
  const pageLayerRef = useRef<HTMLDivElement>(null);
  const shapesInnerRef = useRef<HTMLDivElement>(null);
  const markRef = useRef<HTMLDivElement>(null);
  /**
   * How long the last frame took, per layer, for the hidden readout. A
   * synthetic sheet cannot reproduce what a real design document costs, so
   * these numbers have to come from the app itself.
   */
  const perf = useRef<PaintTimings>({
    overlayMs: 0,
    layers: 0,
    stats: null,
    base: [],
    surface: [],
    text: [],
    cells: [],
  });

  // A window holding two sheets would describe neither, so a switch starts over
  // and the readout says so until the first pass lands.
  useEffect(() => {
    perf.current.base.length = 0;
    perf.current.surface.length = 0;
    perf.current.text.length = 0;
    perf.current.cells.length = 0;
  }, [path, sheet.name]);

  /**
   * Bumped to make the readout recompute, so it tracks a scroll instead of
   * waiting for the next selection change. The ref mirrors `readoutOn` because
   * that flag is read far below this point.
   */
  const [readoutTick, setReadoutTick] = useState(0);
  const lastReadoutTick = useRef(0);
  const readoutOnRef = useRef(false);

  /**
   * A sheet change hands the keyboard to the grid.
   *
   * Switching sheets — or opening a workbook, which amounts to the same — left
   * focus on whatever had been clicked, so PageUp, Home and the rest did
   * nothing until the grid was clicked first. `preventScroll`, because focusing
   * something that fills the view can otherwise jog the scroller.
   */
  useEffect(() => {
    viewportRef.current?.focus({ preventScroll: true });
    window.getSelection()?.removeAllRanges();
  }, [path, sheet.name]);

  const cache = useRef(new Map<string, Cell[][]>());
  const pending = useRef(new Set<string>());
  const drag = useRef<ResizeDrag | null>(null);
  /**
   * A drag-select in progress, and whether it has moved.
   *
   * The anchor is a cell, a whole row or a whole column, because a drag that
   * starts on a gutter picks rows or columns — dragging one row header down is
   * how a run like `3:7` is made by mouse, and only that end moves.
   */
  const rangeDrag = useRef<DragAnchor | null>(null);
  /** Where the pointer last was, which a drag keeps reading after it leaves. */
  const dragPointer = useRef<{ x: number; y: number } | null>(null);
  /** The frame that carries a drag on past the edge of the viewport. */
  const autoScroll = useRef(0);
  /** Set when such a drag ends, so its closing click cannot collapse the block. */
  const suppressClick = useRef(false);

  /** Read inside the paint pass, which runs outside React's render. */
  const selectionRef = useRef(selection);
  selectionRef.current = selection;
  const zoomRef = useRef(zoom / 100);
  zoomRef.current = zoom / 100;
  const scrollRef = useRef<SheetView>({ top: 0, left: 0 });
  /**
   * The cursor the view has already been moved to.
   *
   * The pass follows a cursor the user moved — an arrow key, a search hit, Go To
   * — and nothing else. A sheet arriving is not that: its view was just put back
   * where it was left, and the cursor it carries is the one from last time,
   * which is off screen only because it was scrolled out of the way. Excel keeps
   * that state rather than undoing it, so arriving must not count as a move.
   * The cursor is recorded when the sheet changes, since the pass itself runs
   * again for everything that lands with the arrival — the layout, the column
   * metrics — and every one of those runs has to leave the view alone.
   */
  const followedRef = useRef<Selection | null>(null);
  /** The last jump token the pass acted on, so only a new one may drag the view. */
  const revealRef = useRef(reveal ?? 0);
  const sizeRef = useRef({ width: 0, height: 0 });
  const dprRef = useRef(1);

  const [baseColW, setBaseColW] = useState(96);
  const [baseRowH, setBaseRowH] = useState(24);
  const [colOverrides, setColOverrides] = useState<Map<number, number>>(() => new Map());
  const [rowOverrides, setRowOverrides] = useState<Map<number, number>>(() => new Map());
  const [layout, setLayout] = useState<SheetLayout | null>(null);
  const [layoutSheet, setLayoutSheet] = useState<string | null>(null);
  const [fontFamily, setFontFamily] = useState<string | null>(null);
  const [fontSize, setFontSize] = useState(12);
  const [showGrid, setShowGrid] = useState(true);
  const [filters, setFilters] = useState<Map<number, string[]>>(() => new Map());
  const [visibleRows, setVisibleRows] = useState<number[] | null>(null);
  /**
   * How much blank area the scroller holds past the sheet's own extent, in rows
   * and columns — as much of Excel's own empty thousands as the view is about
   * to want, grown as it goes and given back when it comes away.
   */
  const [voidRows, setVoidRows] = useState(VOID_MIN_ROWS);
  const [voidCols, setVoidCols] = useState(VOID_MIN_COLS);
  /** Bumped when a block lands, so the paint pass knows to run again. */
  const [version, setVersion] = useState(0);
  /** The hover readout, and the cell it belongs to so a move within it is free. */
  const [tip, setTip] = useState<{ text: string; x: number; y: number } | null>(null);
  const hoverKey = useRef<string | null>(null);
  /**
   * What Ctrl+C took, and which sheet it came from.
   *
   * Excel's copy is otherwise silent — no dialog, no flash — so this marquee is
   * the only sign it worked, and it stays until Escape or the next copy, on the
   * sheet it came from rather than on whichever sheet is showing.
   */
  const copiedRef = useRef<{ path: string; sheet: string; rect: MergeRect } | null>(null);
  /** How far the marquee's dashes have crawled, walked by the timer below. */
  const antsRef = useRef(0);
  const antsTimer = useRef(0);

  const layoutReady = layout !== null && layoutSheet === sheet.name;

  /**
   * Ctrl+A, in Excel's two stages: the cell's current region, then the whole
   * sheet.
   *
   * The region is what the first press has to mean — not the sheet's used
   * extent, which is what it was, and which is the whole sheet whenever the
   * document fills it. The two are indistinguishable on a sheet that is one
   * table, which is most of them, but a document with two tables and a blank
   * row between can tell them apart.
   */
  const selectAll = useCallback(async () => {
    const sheetRows = Math.max(sheet.rows, layout?.rows ?? 0);
    const active = selectionAnchor(selection);
    let region;
    try {
      region = await currentRegion(path, sheet.name, active.row, active.col);
    } catch {
      // The backend is gone; everything is a more honest answer than nothing.
      onSelect(rowsBetween(0, sheetRows - 1));
      return;
    }
    const onRegion =
      selection.kind === "range" &&
      selection.row === region.row &&
      selection.col === region.col &&
      selection.rows === region.rows &&
      selection.cols === region.cols;
    // One cell is what a blank anchor gives — nothing surrounds it, so there is
    // no region to select. Excel falls back to the whole sheet there, and it has
    // to: a block of one cell looks exactly like the cell that was already
    // selected, which would leave the first press appearing to do nothing and
    // the second one doing everything.
    const alone = region.rows === 1 && region.cols === 1;
    onSelect(
      alone || onRegion
        ? rowsBetween(0, sheetRows - 1)
        : rangeOver(region.row, region.col, region.rows, region.cols),
    );
  }, [path, sheet.name, selection, onSelect, sheet.rows, layout]);

  /**
   * Taken on `window`, not in the viewport's own handler.
   *
   * A press that follows a click on the toolbar or on the start page lands on the
   * body, where the browser's own select-all highlights the page's DOM, leaves
   * the sheet unselected, and leaves that highlight behind afterwards. A field
   * keeps its own Ctrl+A: there it is the user selecting the text they typed.
   */
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== "a") return;
      const target = event.target as HTMLElement | null;
      if (target?.closest("input, textarea, [contenteditable]")) return;
      event.preventDefault();
      void selectAll();
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [selectAll]);

  /**
   * Merge rectangles clipped to the cells the parser actually reported. A file
   * can declare regions beyond the used range, and those have no offsets to
   * render at.
   */
  const sheetRowsGuess = Math.max(sheet.rows, layout?.rows ?? 0);
  const sheetColsGuess = Math.max(sheet.cols, layout?.cols ?? 0);

  const merges = useMemo(() => {
    const raw = layout?.merges ?? [];
    const out: MergeRect[] = [];
    for (const merge of raw) {
      if (merge.row >= sheetRowsGuess || merge.col >= sheetColsGuess) continue;
      out.push({
        row: merge.row,
        col: merge.col,
        rows: Math.min(merge.rows, sheetRowsGuess - merge.row),
        cols: Math.min(merge.cols, sheetColsGuess - merge.col),
      });
    }
    return out;
  }, [layout, sheetRowsGuess, sheetColsGuess]);

  const autoFilter = layout?.autoFilter ?? null;
  const printArea = layout?.printArea ?? null;

  /**
   * The columns carrying a filter button. All of them are mounted and moved
   * imperatively rather than being mounted per visible range: a table is a few
   * dozen columns wide at most, and the overlay clips whatever has scrolled
   * out, so tracking the range in React would buy nothing.
   */
  const filterCols = useMemo(() => {
    if (!autoFilter) return [] as number[];
    const out: number[] = [];
    for (let c = autoFilter.col; c < autoFilter.col + autoFilter.cols; c += 1) {
      if (c < sheetColsGuess) out.push(c);
    }
    return out;
  }, [autoFilter, sheetColsGuess]);

  const outsidePrint = useCallback(
    (row: number, col: number): boolean => {
      if (!printArea) return false;
      return (
        row < printArea.row ||
        row >= printArea.row + printArea.rows ||
        col < printArea.col ||
        col >= printArea.col + printArea.cols
      );
    },
    [printArea],
  );

  const displayRows = useMemo(() => {
    if (!visibleRows) return null;
    const headerRow = autoFilter?.row ?? -1;
    if (headerRow < 0) return visibleRows;
    const above = Array.from({ length: headerRow + 1 }, (_, index) => index);
    return [...above, ...visibleRows.filter((row) => row > headerRow)];
  }, [visibleRows, autoFilter]);

  const rowView = useMemo(() => {
    if (!displayRows) return null;
    const index = new Map<number, number>();
    displayRows.forEach((sheetRow, display) => index.set(sheetRow, display));
    return { rows: displayRows, index };
  }, [displayRows]);

  /** How many rows the view shows — a filtered view shows fewer. */
  const displayRowCount = rowView ? rowView.rows.length : sheetRowsGuess;

  /**
   * A display row's place on the sheet.
   *
   * Past the rows this view has, the rest are the blank ones the scroller moves
   * through, numbered on from the sheet's own extent — a filtered view has no
   * room to put them among its own rows, and out there is where they would be
   * hidden anyway. Answering -1 instead left that whole area unpainted.
   */
  const sheetRowOf = useCallback(
    (display: number): number =>
      rowView ? (rowView.rows[display] ?? sheetRowsGuess + display - displayRowCount) : display,
    [rowView, sheetRowsGuess, displayRowCount],
  );
  const displayOfRow = useCallback(
    (sheetRow: number): number => (rowView ? (rowView.index.get(sheetRow) ?? -1) : sheetRow),
    [rowView],
  );

  /**
   * What the scroller holds in all: the sheet's own rows and columns, then the
   * blank area. Everything that has to reach the far edge of the scroll — the
   * extent, the hit test — is measured against these rather than the sheet's.
   */
  const rowCount = displayRowCount + voidRows;
  const colCount = sheetColsGuess + voidCols;

  const cols = useMemo(
    () =>
      buildMetrics(
        colCount,
        // `<col>` describes columns a sheet may hold nothing in, and those runs
        // reach past the last cell; outside the extent the sheet's own default
        // width is the honest answer, not the file's.
        (index) => (index < sheetColsGuess ? (colOverrides.get(index) ?? baseColW) : baseColW),
      ),
    [colCount, sheetColsGuess, colOverrides, baseColW],
  );

  const rows = useMemo(
    () =>
      buildMetrics(rowCount, (display) => {
        const row = sheetRowOf(display);
        return row >= 0 && row < sheetRowsGuess ? (rowOverrides.get(row) ?? baseRowH) : baseRowH;
      }),
    [rowView, rowCount, sheetRowsGuess, rowOverrides, baseRowH, sheetRowOf],
  );

  const minRowH = Math.max(MIN_ROW_H, Math.ceil(fontSize * LINE_HEIGHT));

  /**
   * Where the sheet's own extent ends, in sheet pixels.
   *
   * The blank area is a prefix on top of this and never moves it, so this is
   * both the boundary the file's formatting stops at and what the shapes and
   * the print mark are placed against — `cols.total` and `rows.total` now run
   * on past it, and measuring against those put them out in the empty area.
   */
  const usedRowsTotal = rows.offsets[displayRowCount] ?? rows.total;
  const usedColsTotal = cols.offsets[sheetColsGuess] ?? cols.total;

  /**
   * Tops the blank area up when the view has run out of it, and gives back what
   * the view has come away from.
   *
   * The top-up waits until the view actually reaches the end of what is held.
   * Measuring it against the *sheet's* end instead — topping up whenever the
   * view came within a screen of the data — moved the end of the scrollbar
   * along with the thumb for the whole length of the sheet: it grew from the
   * middle outwards and the thumb never approached the right-hand side, which
   * is not how Excel's reads. So the range stands still while there is still
   * somewhere to go, and it moves half a screen at a time — which is also how
   * often either direction costs a state update. A whole screen was tried and
   * read as a lurch: the thumb fell from the far end to two-thirds of the way
   * along every time the area was topped up.
   *
   * The giving-back is the hysteresis: blank is only released once the view has
   * come two screens away from what is held, and then down to one.
   */
  const resizeVoid = useCallback(() => {
    const el = viewportRef.current;
    if (!el) return;
    const z = zoomRef.current;
    const scroll = scrollRef.current;
    const width = el.clientWidth / z;
    const height = el.clientHeight / z;
    // The view's far edges in sheet space, the gutters coming off as they do in
    // `buildView`: both scroll with the content.
    const bottom = Math.max(0, scroll.top - COL_HEADER_H) + height;
    const right = Math.max(0, scroll.left - ROW_HEADER_W) + width;

    /**
     * `keep` is the blank held in front of the view once the giving-back has
     * run; a step adds half of what a screen shows, so the blank in front stays
     * under a screen and the thumb never falls far when it is topped up.
     */
    const topUp = (
      current: number,
      edge: number,
      used: number,
      size: number,
      keep: number,
      step: number,
    ): number => {
      const held = used + current * size;
      if (edge >= held) return current + step;
      if (held - edge > 2 * keep * size) {
        return Math.max(keep, Math.ceil((edge + keep * size - used) / size));
      }
      return current;
    };

    const screenRows = Math.ceil(height / baseRowH);
    const screenCols = Math.ceil(width / baseColW);
    const keepRows = Math.max(VOID_MIN_ROWS, screenRows);
    const keepCols = Math.max(VOID_MIN_COLS, screenCols);
    const stepRows = Math.max(1, Math.ceil(screenRows / 2));
    const stepCols = Math.max(1, Math.ceil(screenCols / 2));

    setVoidRows((current) => topUp(current, bottom, usedRowsTotal, baseRowH, keepRows, stepRows));
    setVoidCols((current) => topUp(current, right, usedColsTotal, baseColW, keepCols, stepCols));
  }, [usedRowsTotal, usedColsTotal, baseRowH, baseColW]);
  const resizeVoidRef = useRef(resizeVoid);
  resizeVoidRef.current = resizeVoid;

  /**
   * A sheet arriving — and a filter, and a zoom — moves the used range the
   * blank area is measured from, and none of those come through a scroll.
   *
   * Laid out rather than deferred, so the blank area is the new sheet's before
   * the frame that arrives with it is painted. Left to a passive effect the
   * first paint of a switch measured itself against the sheet being left, which
   * is the "wrong until you scroll it" frame: the paint runs off the metrics
   * this render produced, so the metrics have to be the new sheet's by then.
   */
  useLayoutEffect(() => {
    resizeVoidRef.current();
  }, [layoutSheet, zoom, baseColW, baseRowH, rowView, cols, rows]);

  /**
   * Puts the scroller back on the position it believes it holds.
   *
   * The blank area is grown from what the view wants, so it arrives a render
   * after the sheet does — and a sheet restored to a view deep in the blank
   * area can find its position clamped before that area exists, leaving the
   * scroller at one offset and the paint at another. Growing the area fixes the
   * reach; this puts the position back on it.
   *
   * A scroll the user made sets both to the same thing, so this only ever moves
   * the clamped case.
   */
  useLayoutEffect(() => {
    const el = viewportRef.current;
    if (!el) return;
    const z = zoomRef.current;
    const { top, left } = scrollRef.current;
    if (Math.abs(el.scrollTop - top * z) > 0.5) el.scrollTop = top * z;
    if (Math.abs(el.scrollLeft - left * z) > 0.5) el.scrollLeft = left * z;
  }, [cols, rows]);

  /**
   * A column's left edge, for placing the shapes that float above the cells.
   *
   * A shape whose anchor column is past the sheet's extent — one this pass could
   * not place, since a one-cell anchor is only an offset the parser has to turn
   * back into a cell — keeps to the sheet's own right edge. Reading the
   * scroller's total put those shapes out in the blank area instead, which drew
   * the left of the drawing again to the right of the sheet.
   */
  const shapeColX = useCallback(
    (col: number) =>
      col >= 0 && col <= sheetColsGuess ? cols.offsets[col] : usedColsTotal,
    [cols, sheetColsGuess, usedColsTotal],
  );

  /**
   * The draw-space edge of the print area's own boundary.
   *
   * Unlike `shapeColX`/`shapeRowY` these do not stop at the sheet's extent: a
   * print area is allowed to reach past the last cell the sheet formats — that
   * is how a page is given room to breathe — and Excel draws its boundary out
   * there all the same. Clamping to the used range pulled the mark up by
   * however many rows the page ran past the data. It is measured out from the
   * used range rather than from `cols.total`, which now runs into the blank
   * area and would put the mark hundreds of columns too wide.
   */
  const printColX = useCallback(
    (col: number) =>
      col >= 0 && col <= sheetColsGuess
        ? cols.offsets[col]
        : usedColsTotal + (col - sheetColsGuess) * baseColW,
    [cols, sheetColsGuess, usedColsTotal, baseColW],
  );

  const printRowY = useCallback(
    (row: number) => {
      const display = displayOfRow(row);
      if (display >= 0 && display < displayRowCount) return rows.offsets[display];
      // A filtered view has no room to extrapolate in; the end of the rows on
      // screen is as far as the mark can honestly reach.
      if (rowView) return usedRowsTotal;
      return usedRowsTotal + (row - sheetRowsGuess) * baseRowH;
    },
    [displayOfRow, displayRowCount, rows, rowView, sheetRowsGuess, usedRowsTotal, baseRowH],
  );

  /**
   * A sheet row's top edge, for the same reason. A row a filter has hidden
   * takes the place of the next visible one, so a shape anchored to it shrinks
   * rather than disappearing — and a row past the extent keeps to the extent,
   * which is what `shapeColX` does for columns.
   */
  const shapeRowY = useCallback(
    (row: number) => {
      const last = Math.min(sheetRowsGuess, Math.max(0, row) + 4096);
      for (let probe = Math.max(0, row); probe <= last; probe += 1) {
        const display = displayOfRow(probe);
        if (display >= 0 && display < displayRowCount) return rows.offsets[display];
      }
      return usedRowsTotal;
    },
    [displayOfRow, displayRowCount, rows, sheetRowsGuess, usedRowsTotal],
  );

  const cellAt = useCallback(
    (row: number, col: number): Cell | null => {
      const br = Math.floor(row / BLOCK_ROWS);
      const bc = Math.floor(col / BLOCK_COLS);
      const block = cache.current.get(`${sheet.name}#${br}#${bc}`);
      if (!block) return null;
      return block[row - br * BLOCK_ROWS]?.[col - bc * BLOCK_COLS] ?? null;
    },
    [sheet.name],
  );

  const styleAt = useCallback(
    (row: number, col: number): CellStyle | null => {
      if (!layout) return null;
      // Everything past the sheet's own extent is drawn plain, whatever the
      // file says about the rows and columns reaching into it. A row's style is
      // written for all 16,384 columns of that row and a column's for all
      // 1,048,576 rows of that column, so a table whose columns are ruled and
      // shaded would carry that ladder of rules and its shaded band on past the
      // end of its data — out into what should be empty. Only the extent takes
      // the file's formatting.
      if (row >= sheetRowsGuess || col >= sheetColsGuess) return null;
      const index =
        styleOfRow(layout.cellRuns, row, col) ??
        layout.rowStyles[row] ??
        styleOfColumn(layout.colStyles, col);
      return index === undefined ? null : (layout.styles[index] ?? null);
    },
    [layout, sheetRowsGuess, sheetColsGuess],
  );

  const mergeAt = useCallback(
    (row: number, col: number): MergeRect | null =>
      merges.find(
        (merge) =>
          row >= merge.row &&
          row < merge.row + merge.rows &&
          col >= merge.col &&
          col < merge.col + merge.cols,
      ) ?? null,
    [merges],
  );

  /** Selection always lands on a merge's anchor, never inside one. */
  const select = useCallback(
    (row: number, col: number) => {
      const merge = mergeAt(row, col);
      onSelect({
        kind: "cell",
        row: merge ? merge.row : row,
        col: merge ? merge.col : col,
      });
    },
    [mergeAt, onSelect],
  );

  /**
   * Ctrl+arrow: Excel's jump to the edge of the data.
   *
   * The backend says where the run in that direction ends; nothing ahead at all
   * is answered with the sheet's own edge, which is what Excel's Ctrl+End means
   * by the last cell too. A block keeps its anchor under Shift and collapses to
   * the cell it lands on without it, and a run of whole rows or columns — which
   * has only the one axis it runs along — slides along that axis, measured from
   * the cell the run's own edge sits on.
   */
  const jumpToEdge = useCallback(
    async (dir: "up" | "down" | "left" | "right", extend: boolean) => {
      const sheetRows = Math.max(sheet.rows, layout?.rows ?? 0);
      const sheetCols = Math.max(sheet.cols, layout?.cols ?? 0);
      const from = selectionAnchor(selection);
      const vertical = dir === "up" || dir === "down";
      let to: { row: number; col: number } | null = null;
      try {
        to = await edgeCell(path, sheet.name, from.row, from.col, dir);
      } catch {
        // The backend is gone, and the sheet's own edge is still the answer
        // Excel would give from a cell with nothing after it.
      }
      const row = to?.row ?? (vertical ? (dir === "up" ? 0 : sheetRows - 1) : from.row);
      const col = to?.col ?? (vertical ? from.col : dir === "left" ? 0 : sheetCols - 1);

      // A run of whole rows or columns moves along the one axis it has; the
      // direction across it has nothing there for it to select. Shift grows the
      // run to the edge rather than moving it, as it does by one on its own.
      if (selection.kind === "row") {
        if (!vertical) return;
        if (extend) {
          onSelect(rowsBetween(selection.anchor, row));
          return;
        }
        const step = Math.max(
          -selection.row,
          Math.min(sheetRows - selection.row - selection.rows, row - selection.row),
        );
        onSelect(rowsBetween(selection.anchor + step, selection.active + step));
        return;
      }
      if (selection.kind === "column") {
        if (vertical) return;
        if (extend) {
          onSelect(colsBetween(selection.anchor, col));
          return;
        }
        const step = Math.max(
          -selection.col,
          Math.min(sheetCols - selection.col - selection.cols, col - selection.col),
        );
        onSelect(colsBetween(selection.anchor + step, selection.active + step));
        return;
      }
      if (extend) onSelect(rangeBetween(selection.kind === "range" ? selection.anchor : from, { row, col }));
      else select(row, col);
    },
    [selection, path, sheet.name, sheet.rows, sheet.cols, layout, onSelect, select],
  );

  /**
   * Files the filter away beside the sheet's scroll position. Without this the
   * sheet's layout is fetched again on the way back in and the filter is seeded
   * from the workbook's own `<autoFilter>` — which knows nothing about a
   * filter applied here, since this is a viewer and never writes one back.
   */
  const rememberFilters = useCallback(
    (next: Map<number, string[]>) => {
      const bySheet = views.get(path) ?? new Map<string, SheetView>();
      const view = bySheet.get(sheet.name) ?? { ...scrollRef.current };
      bySheet.set(sheet.name, { ...view, filters: next });
      views.set(path, bySheet);
    },
    [views, path, sheet.name],
  );

  const applyFilters = useCallback(
    async (next: Map<number, string[]>, target: AutoFilter | null) => {
      setFilters(next);
      rememberFilters(next);
      try {
        setVisibleRows(await computeVisibleRows(path, sheet.name, target, next));
      } catch {
        setVisibleRows(null);
      }
    },
    [path, sheet.name, rememberFilters],
  );

  // --- the sheet's layout, once per switch ---
  useEffect(() => {
    drag.current = null;
    // The incoming sheet's own cursor is where its view was left, so arriving
    // with it is not a cursor move for the pass below to carry.
    followedRef.current = selectionRef.current;
    let cancelled = false;
    sheetLayout(path, sheet.name)
      .then((parsed) => {
        if (cancelled) return;
        setLayout(parsed);
        setLayoutSheet(sheet.name);
        const colPx = (chars: number) =>
          charsToPx(chars, digitWidthPx(parsed.fontFamily, parsed.fontSizePx));
        setBaseColW(colPx(parsed.defaultColWidthChars));
        setBaseRowH(parsed.defaultRowHeightPx);
        setFontFamily(parsed.fontFamily);
        setFontSize(parsed.fontSizePx);
        setShowGrid(parsed.showGridLines);
        setColOverrides(toOverrides(parsed.colWidthsChars, colPx));
        setRowOverrides(toOverrides(parsed.rowHeightsPx));

        // Put back the view this sheet was left at, before anything is
        // fetched, so the first paint is the right block rather than A1's.
        // Excel keeps a window position per sheet and restores it on the way
        // in, even where that leaves the active cell off screen; seeding this
        // from the cursor instead dragged a sheet whose cursor sat far out to
        // the right until that cell was the first visible column.
        const view = views.get(path)?.get(sheet.name);
        setScrollTo(view?.top ?? 0, view?.left ?? 0);

        // What this sheet was left holding beats what the workbook stores: a
        // filter applied here is in no file, and an empty Map is a filter the
        // user cleared rather than one that was never set.
        const saved =
          views.get(path)?.get(sheet.name)?.filters ??
          toFilters(parsed.autoFilter?.applied);
        if (saved.size > 0) {
          setFilters(saved);
          computeVisibleRows(path, sheet.name, parsed.autoFilter, saved)
            .then((found) => {
              if (!cancelled) setVisibleRows(found);
            })
            .catch(() => {});
        } else {
          setFilters(new Map());
          setVisibleRows(null);
        }
      })
      .catch(() => {
        // .xls and .ods have no zip layout to read; the fallbacks stand.
      });
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- setScrollTo is stable
  }, [path, sheet.name, views]);

  /** Moves the scroller to a sheet-space offset, through the paint refs. */
  const setScrollTo = useCallback((top: number, left: number) => {
    scrollRef.current = { top, left };
    const el = viewportRef.current;
    const z = zoomRef.current;
    if (el) {
      el.scrollTop = top * z;
      el.scrollLeft = left * z;
    }
  }, []);

  // --- the client box, and the canvases scaled to it ---
  useEffect(() => {
    const el = viewportRef.current;
    if (!el) return;
    const measure = () => {
      sizeRef.current = { width: el.clientWidth, height: el.clientHeight };
      // The bar still occupies its height whatever the clipping above it does,
      // so this reads the same before and after the scroller is grown.
      setHBarHeight(hideHScroll ? el.offsetHeight - el.clientHeight : 0);
      const dpr = window.devicePixelRatio || 1;
      dprRef.current = dpr;
      for (const canvas of [baseRef.current, overlayRef.current]) {
        if (!canvas) continue;
        // The backing store is in device pixels; the CSS box stays in CSS ones,
        // or the canvas would be laid out at its device size.
        canvas.width = Math.max(1, Math.round(el.clientWidth * dpr));
        canvas.height = Math.max(1, Math.round(el.clientHeight * dpr));
        canvas.style.width = `${el.clientWidth}px`;
        canvas.style.height = `${el.clientHeight}px`;
      }
      paintRef.current("both");
      // A wider window is more blank area to want than it had before.
      resizeVoidRef.current();
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
    // Hiding the horizontal scrollbar gives the cells that much more height
    // without changing the scroller's own box, which no observer here sees.
  }, [hideHScroll]);

  useEffect(() => {
    onViewport?.(viewportRef.current);
    return () => onViewport?.(null);
  }, [onViewport]);

  // Fonts land after the first paint on a cold start, and canvas does not
  // reflow the way the DOM did — so the metrics have to be dropped and
  // everything drawn again once they are in.
  useEffect(() => {
    let cancelled = false;
    void document.fonts?.ready.then(() => {
      if (cancelled) return;
      clearTextCaches();
      paintRef.current("both");
    });
    return () => {
      cancelled = true;
    };
  }, []);

  /**
   * Ctrl+wheel zooms, as it does in Excel. The listener is ours rather than
   * React's because it has to `preventDefault` — otherwise the webview zooms
   * the whole page as well as the sheet.
   */
  useEffect(() => {
    const el = viewportRef.current;
    if (!el) return;
    const onWheel = (event: WheelEvent) => {
      if (!event.ctrlKey && !event.metaKey) return;
      event.preventDefault();
      const percent = Math.round(zoomRef.current * 100);
      onZoom(stepZoom(percent, event.deltaY < 0 ? 1 : -1));
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    return () => el.removeEventListener("wheel", onWheel);
  }, [onZoom]);

  // A zoom change resizes the spacer, so the scroller has to be put back on the
  // same sheet coordinates or the view jumps.
  useLayoutEffect(() => {
    const el = viewportRef.current;
    if (!el) return;
    el.scrollTop = scrollRef.current.top * zoomRef.current;
    el.scrollLeft = scrollRef.current.left * zoomRef.current;
    paintRef.current("both");
    resizeVoidRef.current();
  }, [zoom]);

  // --- paint scheduling ---
  const rafRef = useRef(0);
  const pendingLayers = useRef<"both" | "overlay">("overlay");
  const schedule = useCallback((layers: "both" | "overlay") => {
    if (layers === "both") pendingLayers.current = "both";
    if (rafRef.current) return;
    rafRef.current = requestAnimationFrame(() => {
      rafRef.current = 0;
      const next = pendingLayers.current;
      pendingLayers.current = "overlay";
      paintRef.current(next);
    });
  }, []);

  /**
   * The pending repaint, if one is scheduled. Cancelling without clearing the
   * id would leave `schedule` believing a frame was still queued and it would
   * return early for ever — which is exactly what a StrictMode mount/unmount/
   * mount did, killing every scroll and sheet-switch repaint.
   */
  useEffect(
    () => () => {
      cancelAnimationFrame(rafRef.current);
      rafRef.current = 0;
      window.clearInterval(antsTimer.current);
    },
    [],
  );

  /**
   * Marks what Ctrl+C just took, and keeps its marquee crawling until it is
   * taken away again.
   *
   * The crawl is what makes it read as a copy rather than as one more border,
   * so the timer starts and stops with the marquee instead of running for the
   * life of the grid.
   */
  const markCopied = useCallback(
    (next: { path: string; sheet: string; rect: MergeRect } | null) => {
      copiedRef.current = next;
      window.clearInterval(antsTimer.current);
      antsTimer.current = 0;
      if (next) {
        antsTimer.current = window.setInterval(() => {
          antsRef.current += ANTS_STEP;
          schedule("overlay");
        }, ANTS_MS);
      }
      schedule("overlay");
    },
    [schedule],
  );

  /**
   * Escape takes the marquee away, which is the other half of what a copy means:
   * one that is over looks different from one that still stands.
   *
   * On `window`, like Ctrl+A, and not in the viewport's own handler, because the
   * grid is not necessarily where the focus is by then: writing the clipboard can
   * leave it on the body, and clicking a toolbar button does the same. A field
   * keeps its own Escape — there it belongs to what the user is typing.
   */
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      const target = event.target as HTMLElement | null;
      if (target?.closest("input, textarea, [contenteditable]")) return;
      markCopied(null);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [markCopied]);

  const handleScroll = useCallback(() => {
    const node = viewportRef.current;
    if (!node) return;
    const z = zoomRef.current;
    const bySheet = views.get(path) ?? new Map<string, SheetView>();
    // Spreads whatever the sheet already holds, so scrolling does not drop the
    // filter filed beside the position.
    const view = {
      ...bySheet.get(sheet.name),
      top: node.scrollTop / z,
      left: node.scrollLeft / z,
    };
    scrollRef.current = view;
    bySheet.set(sheet.name, view);
    views.set(path, bySheet);
    // The blank area is what makes a scroll past the data possible at all, so
    // every scroll asks for as much of it as the next one will want.
    resizeVoidRef.current();
    schedule("both");
  }, [schedule, views, path, sheet.name]);

  /** Repaint the block layer whenever the data, the layout or the sheet moves. */
  useEffect(() => {
    schedule("both");
  }, [
    schedule,
    version,
    layout,
    layoutSheet,
    cols,
    rows,
    colOverrides,
    rowOverrides,
    baseColW,
    baseRowH,
    showGrid,
    fontFamily,
    fontSize,
    filters,
    visibleRows,
    path,
    sheet.name,
  ]);

  /**
   * A selection change is the cheap one: the overlay alone.
   *
   * Painted from a layout effect, and synchronously, rather than from a passive
   * effect scheduled onto the next frame. The canvas IS this component's
   * output, so waiting for a deferred callback made the ring visibly trail the
   * click while the cell layer — driven straight from the scroll handler —
   * stayed ahead of it.
   */
  useLayoutEffect(() => {
    paintRef.current("overlay");
  }, [selection]);

  /**
   * Brings the selected cell into view.
   *
   * Moving the selection with the keyboard, from a search hit or from Go To has
   * to carry the view with it, or the cursor walks off the bottom of a screen
   * that never moves — which is what a selected row or column did too.
   *
   * Nothing else moves it. A sheet that arrives has had its view put back where
   * it was left, and a jump that lands on another sheet is the one kind of
   * arrival whose cursor *was* asked for.
   */
  useEffect(() => {
    const el = viewportRef.current;
    if (!el || !layoutReady) return;
    const asked = (reveal ?? 0) !== revealRef.current;
    revealRef.current = reveal ?? 0;
    const follow = asked || selection !== followedRef.current;
    followedRef.current = selection;
    if (!follow) return;
    const factor = zoom / 100;

    // A block follows its active cell rather than its rectangle, which for one
    // growing leftward is the wrong corner; the active cell is the end the
    // keyboard walks, so an extension carries it into view as it goes.
    const atRow =
      selection.kind === "column"
        ? -1
        : selection.kind === "range"
          ? selection.active.row
          : selection.kind === "row"
            ? selection.active
            : selection.row;
    const atCol =
      selection.kind === "row"
        ? -1
        : selection.kind === "range"
          ? selection.active.col
          : selection.kind === "column"
            ? selection.active
            : selection.col;

    // Scrolling *up* to a cell means matching its offset, not the offset plus
    // the gutter: the gutter scrolls with the content and would leave the cell
    // tucked behind the header.
    if (atRow >= 0) {
      const display = displayOfRow(atRow);
      if (display >= 0) {
        const top = (rows.offsets[display] ?? 0) * factor;
        const bottom =
          (COL_HEADER_H + (rows.offsets[Math.min(display + 1, displayRowCount)] ?? 0)) * factor;
        if (top < el.scrollTop) el.scrollTop = top;
        else if (bottom > el.scrollTop + el.clientHeight) el.scrollTop = bottom - el.clientHeight;
      }
    }
    if (atCol >= 0) {
      const left = (cols.offsets[atCol] ?? 0) * factor;
      const right =
        (ROW_HEADER_W + (cols.offsets[Math.min(atCol + 1, sheetColsGuess)] ?? 0)) * factor;
      if (left < el.scrollLeft) el.scrollLeft = left;
      else if (right > el.scrollLeft + el.clientWidth) el.scrollLeft = right - el.clientWidth;
    }
  }, [
    selection,
    rows,
    cols,
    displayOfRow,
    displayRowCount,
    sheetColsGuess,
    zoom,
    layoutReady,
    reveal,
  ]);

  // --- fetching the cells the viewport needs ---
  /**
   * Fetches whatever blocks the window needs and does not already have.
   *
   * Called from the paint pass rather than from an effect, and that is the
   * point: the scroll offset lives in a ref, so nothing about this component
   * changes identity when it moves — an effect keyed on the shape of the view
   * never re-ran on a scroll, so every row past the first screenful stayed
   * blank no matter how far down you went.
   */
  const ensureBlocks = useCallback(
    (view: GridView) => {
      const firstCol = Math.floor(view.colStart / BLOCK_COLS);
      const lastCol = Math.floor(Math.max(view.colStart, view.colEnd - 1) / BLOCK_COLS);
      /**
       * Past the used range there is nothing to fetch.
       *
       * Every cell out there is blank and the formatting stops there too, so a
       * block beyond it is a round trip carrying a block's worth of cells to
       * learn nothing. That trip is not free even so: its reply renders the
       * grid again, and it is asked for on exactly the frame the view runs off
       * the end of the data into the blank area — which is where the scroll
       * visibly slowed down. The blank area needs no data.
       */
      const lastUsedCol = Math.ceil(sheetColsGuess / BLOCK_COLS) - 1;
      const lastUsedRow = Math.ceil(sheetRowsGuess / BLOCK_ROWS) - 1;
      const needed = new Set<string>();
      for (let display = view.rowStart; display < view.rowEnd; display += 1) {
        const sheetRow = view.sheetRowOf(display);
        if (sheetRow < 0) continue;
        const block = Math.floor(sheetRow / BLOCK_ROWS);
        if (block > lastUsedRow) continue;
        for (let bc = firstCol; bc <= lastCol && bc <= lastUsedCol; bc += 1) {
          needed.add(`${sheet.name}#${block}#${bc}`);
        }
      }

      for (const key of needed) {
        if (cache.current.has(key) || pending.current.has(key)) continue;
        pending.current.add(key);
        const [, rawRow, rawCol] = key.split("#");
        getCells(
          path,
          sheet.name,
          Number(rawRow) * BLOCK_ROWS,
          Number(rawCol) * BLOCK_COLS,
          BLOCK_ROWS,
          BLOCK_COLS,
        )
          .then((data) => {
            cache.current.set(key, data);
            setVersion((n) => n + 1);
          })
          .catch(() => {
            cache.current.set(key, []);
          })
          .finally(() => pending.current.delete(key));
      }
    },
    [path, sheet.name, sheetColsGuess, sheetRowsGuess],
  );

  const buildView = useCallback((): GridView | null => {
    if (!layout || layoutSheet !== sheet.name) return null;
    const scroll = scrollRef.current;
    const size = sizeRef.current;
    const z = zoomRef.current;
    if (size.width <= 0 || size.height <= 0) return null;

    const sheetRows = Math.max(sheet.rows, layout.rows);
    const sheetCols = Math.max(sheet.cols, layout.cols);

    const top = Math.max(0, scroll.top - COL_HEADER_H);
    const left = Math.max(0, scroll.left - ROW_HEADER_W);
    const width = size.width / z;
    const height = size.height / z;

    const range = {
      rowStart: Math.max(0, indexAt(rows.offsets, top) - OVERSCAN_ROWS),
      // The window runs over everything the scroller holds, blank area
      // included: its cells are empty and carry no formatting, but the grid
      // goes on being drawn there.
      rowEnd: Math.min(rowCount, indexAt(rows.offsets, top + height) + 1 + OVERSCAN_ROWS),
      colStart: Math.max(0, indexAt(cols.offsets, left) - OVERSCAN_COLS),
      colEnd: Math.min(colCount, indexAt(cols.offsets, left + width) + 1 + OVERSCAN_COLS),
    };

    // Which merged regions and Center-Across runs are on screen — both depend
    // on the range, so they are worked out here rather than in a memo.
    const first = sheetRowOf(range.rowStart);
    const last = sheetRowOf(Math.max(range.rowStart, range.rowEnd - 1));
    const lo = Math.min(first, last);
    const hi = Math.max(first, last);

    /**
     * Style resolution walks the cell's row runs and then its column runs, so
     * asking per cell per frame is linear in the number of runs. This pass asks
     * about each visible cell more than once (once for the runs, once for the
     * paint, once per neighbour its borders resolve against), so the answers are
     * held for the length of the pass.
     *
     * A window array indexed by display position was tried here, on the theory
     * that the hashing was the cost. It measured no faster — the pass is not
     * bound by any one lookup — and the index arithmetic was the only part of
     * this file that ever produced a silent wrong-border bug. The map stays.
     */
    const styleMemo = new Map<number, CellStyle | null>();
    const styleCalls = { resolved: 0, memo: 0 };
    const styleAtFast = (row: number, col: number): CellStyle | null => {
      // Nothing out in the blank area has a style to resolve, and saying so
      // here rather than inside `styleAt` keeps those cells out of the readout's
      // count as well — the number is meant to say what the pass had to look
      // up, and a screenful of empty offsets is not work the file asked for.
      if (row >= sheetRows || col >= sheetCols) return null;
      const key = cellKey(sheetCols, row, col);
      const hit = styleMemo.get(key);
      if (hit !== undefined) {
        styleCalls.memo += 1;
        return hit;
      }
      styleCalls.resolved += 1;
      const value = styleAt(row, col);
      styleMemo.set(key, value);
      return value;
    };

    const shownMerges: MergeRect[] = [];
    const coveredMerges = new Set<number>();
    /** The same regions as `coveredMerges`, so a lookup is not a scan. */
    const mergeIndex = new Map<number, MergeRect>();
    for (const merge of merges) {
      const end = merge.row + merge.rows - 1;
      if (end < lo || merge.row > hi) continue;
      if (displayOfRow(merge.row) < 0) continue;
      shownMerges.push(merge);
      for (let r = Math.max(merge.row, lo); r <= Math.min(end, hi); r += 1) {
        for (let c = merge.col; c < merge.col + merge.cols; c += 1) {
          const key = cellKey(sheetCols, r, c);
          coveredMerges.add(key);
          mergeIndex.set(key, merge);
        }
      }
    }
    const mergeAtFast = (row: number, col: number): MergeRect | null =>
      mergeIndex.get(cellKey(sheetCols, row, col)) ?? null;

    const shownRuns: (MergeRect & { text: string | null })[] = [];
    const coveredRuns = new Set<number>();
    const joinable = (row: number, col: number) =>
      col >= 0 &&
      col < sheetCols &&
      styleAtFast(row, col)?.horizontal === "centerContinuous" &&
      mergeAtFast(row, col) === null;
    for (let row = lo; row <= hi; row += 1) {
      if (row < 0 || row >= sheetRows || displayOfRow(row) < 0) continue;
      for (let col = range.colStart; col < range.colEnd; col += 1) {
        if (!joinable(row, col)) continue;
        if (joinable(row, col - 1)) continue;
        let end = col;
        while (joinable(row, end + 1)) end += 1;
        if (end === col) continue;
        let text: string | null = null;
        for (let c = col; c <= end && text === null; c += 1) {
          text = cellAt(row, c)?.text ?? null;
        }
        shownRuns.push({ row, col, rows: 1, cols: end - col + 1, text });
        for (let c = col; c <= end; c += 1) coveredRuns.add(cellKey(sheetCols, row, c));
      }
    }

    return {
      cols,
      rows,
      sheetRows,
      sheetCols,
      ...range,
      width: size.width,
      height: size.height,
      zoom: z,
      scrollTop: scroll.top,
      scrollLeft: scroll.left,
      sheetRowOf,
      displayOfRow,
      cellAt,
      styleAt: styleAtFast,
      styleCalls,
      mergeAt: mergeAtFast,
      workbookFontFamily: fontFamily,
      workbookFontSizePx: fontSize,
      showGrid,
      selection: selectionRef.current,
      copied: copiedRef.current?.path === path && copiedRef.current?.sheet === sheet.name
        ? copiedRef.current.rect
        : null,
      ants: antsRef.current,
      merges,
      mergeView: { shown: shownMerges, covered: coveredMerges },
      centerRuns: { shown: shownRuns, covered: coveredRuns },
      autoFilter,
      printArea,
      outsidePrint,
      richText: layout.richText ?? {},
      theme: canvasTheme(),
    };
  }, [
    layout,
    layoutSheet,
    sheet.name,
    sheet.rows,
    sheet.cols,
    rowView,
    cols,
    rows,
    rowCount,
    colCount,
    merges,
    sheetRowOf,
    displayOfRow,
    cellAt,
    styleAt,
    mergeAt,
    fontFamily,
    fontSize,
    showGrid,
    autoFilter,
    printArea,
    outsidePrint,
  ]);

  const paint = useCallback(
    (layers: "both" | "overlay") => {
      const view = buildView();
      // Nothing drawn until the incoming sheet's layout is in: leaving the last
      // sheet's pixels untouched is what keeps a switch from flashing blank.
      if (!view) return;
      // Whatever this frame is about to show has to have its data asked for.
      ensureBlocks(view);
      const startedAt = performance.now();
      const scale = view.zoom * dprRef.current;
      if (layers === "both") {
        const ctx = baseRef.current?.getContext("2d");
        if (ctx) {
          ctx.setTransform(scale, 0, 0, scale, 0, 0);
          const stats = paintBase(ctx, view);
          perf.current.stats = stats;
          // Only real base passes enter the window: an overlay-only frame runs
          // no base pass, and folding a zero in would drag the median down to a
          // lie about how fast this actually is.
          record(perf.current.base, performance.now() - startedAt);
          record(perf.current.surface, stats.surfaceMs);
          record(perf.current.text, stats.textMs);
          record(perf.current.cells, stats.rows * stats.cols);

          // The readout is rebuilt from these refs, so a render is what makes it
          // show them. Throttled: a render per frame during a scroll would be a
          // performance problem of its own, and this is a dev-only panel.
          if (readoutOnRef.current && performance.now() - lastReadoutTick.current > 250) {
            lastReadoutTick.current = performance.now();
            setReadoutTick((n) => n + 1);
          }
        }
      }
      const overlayStarted = performance.now();
      const ctx = overlayRef.current?.getContext("2d");
      if (ctx) {
        ctx.setTransform(scale, 0, 0, scale, 0, 0);
        paintOverlay(ctx, view);
      }
      perf.current.overlayMs = performance.now() - overlayStarted;
      perf.current.layers = layers === "both" ? 2 : 1;

      // The filter buttons are DOM, so they have to be moved by hand: the
      // canvas scrolls under them and a scroll does not re-render React.
      const area = overlayAreaRef.current;
      if (area) {
        const gutterX = ROW_HEADER_W * view.zoom;
        const gutterY = COL_HEADER_H * view.zoom;
        area.style.left = `${gutterX}px`;
        area.style.top = `${gutterY}px`;
        area.style.width = `${Math.max(0, view.width - gutterX)}px`;
        area.style.height = `${Math.max(0, view.height - gutterY)}px`;

        const headerDisplay = view.autoFilter ? view.displayOfRow(view.autoFilter.row) : -1;
        const headerH = headerDisplay >= 0 ? spanAt(view.rows.offsets, headerDisplay) : 0;
        for (const [col, slot] of filterSlots.current) {
          if (headerDisplay < 0) {
            slot.style.display = "none";
            continue;
          }
          slot.style.display = "";
          slot.style.transform =
            `translate(${(view.cols.offsets[col] - view.scrollLeft) * view.zoom}px, ` +
            `${(view.rows.offsets[headerDisplay] - view.scrollTop) * view.zoom}px)`;
          slot.style.width = `${spanAt(view.cols.offsets, col) * view.zoom}px`;
          slot.style.height = `${headerH * view.zoom}px`;
        }
      }

      // The shapes and the print mark are DOM too, and the canvas underneath
      // them does not scroll — so this layer is translated and scaled with the
      // scroll rather than being scrolled, and clipped to the cell area so a
      // shape near row 0 cannot paint over the gutters.
      const page = pageLayerRef.current;
      if (page) {
        const gutterX = ROW_HEADER_W * view.zoom;
        const gutterY = COL_HEADER_H * view.zoom;
        page.style.left = `${gutterX}px`;
        page.style.top = `${gutterY}px`;
        page.style.width = `${Math.max(0, view.width - gutterX)}px`;
        page.style.height = `${Math.max(0, view.height - gutterY)}px`;

        const inner = shapesInnerRef.current;
        if (inner) {
          inner.style.transform =
            `translate(${-view.scrollLeft * view.zoom}px, ${-view.scrollTop * view.zoom}px) ` +
            `scale(${view.zoom})`;
        }

        const mark = markRef.current;
        if (mark) {
          const area = view.printArea;
          if (area) {
            const x0 = printColX(area.col);
            const y0 = printRowY(area.row);
            mark.style.display = "";
            // The line straddles the boundary: shifted out by its inner share
            // and widened by the whole line, so it lands as Excel measured it —
            // 1 screen pixel inside the print area and 2 outside.
            mark.style.left = `${(x0 - view.scrollLeft) * view.zoom - PRINT_AREA_INSIDE}px`;
            mark.style.top = `${(y0 - view.scrollTop) * view.zoom - PRINT_AREA_INSIDE}px`;
            mark.style.width = `${
              (printColX(area.col + area.cols) - x0) * view.zoom + PRINT_AREA_LINE
            }px`;
            mark.style.height = `${
              (printRowY(area.row + area.rows) - y0) * view.zoom + PRINT_AREA_LINE
            }px`;
          } else {
            mark.style.display = "none";
          }
        }
      }
    },
    [buildView, ensureBlocks, printColX, printRowY],
  );
  const paintRef = useRef(paint);
  paintRef.current = paint;

  // --- publishers ---
  useEffect(() => {
    onDetail?.(selection.kind === "cell" ? cellAt(selection.row, selection.col) : null);
  }, [onDetail, cellAt, selection, version]);

  useEffect(() => {
    const sheetRows = Math.max(sheet.rows, layout?.rows ?? 0);
    const sheetCols = Math.max(sheet.cols, layout?.cols ?? 0);
    onExtent?.(sheetRows, sheetCols);
  }, [onExtent, sheet.rows, sheet.cols, layout]);

  // --- hit testing ---
  /**
   * What is under a viewport point: a cell, or one of the gutter strips. The
   * DOM renderer read the `data-gutter` attribute off the element under the
   * cursor; with the gutters painted there is no element, so the classification
   * is by coordinate.
   */
  const hitAt = useCallback(
    (clientX: number, clientY: number): Hit => {
      const el = viewportRef.current;
      if (!el) return null;
      const rect = el.getBoundingClientRect();
      const z = zoomRef.current;
      const px = clientX - rect.left;
      const py = clientY - rect.top;
      const gutterW = ROW_HEADER_W * z;
      const gutterH = COL_HEADER_H * z;

      // The scroller's scrollbars are inside its border box but are not the
      // sheet. Hit-testing one as a cell started a block selection on the
      // rightmost visible column and then dragged it along as the thumb moved —
      // and because the content scrolls under a stationary pointer, the cell
      // climbed a row at a time even though the pointer never left the bar.
      if (px >= el.clientWidth || py >= el.clientHeight) return null;

      if (px < gutterW && py < gutterH) return null; // the corner
      // The gutters are pinned, so their coordinate is the viewport's own.
      if (py < gutterH) {
        const x = (clientX - rect.left + el.scrollLeft) / z - ROW_HEADER_W;
        if (x < 0 || x >= cols.total) return null;
        return { kind: "colHeader", col: Math.min(colCount - 1, indexAt(cols.offsets, x)) };
      }
      if (px < gutterW) {
        const y = (clientY - rect.top + el.scrollTop) / z - COL_HEADER_H;
        if (y < 0 || y >= rows.total) return null;
        const display = Math.min(rowCount - 1, indexAt(rows.offsets, y));
        const row = sheetRowOf(display);
        return row < 0 ? null : { kind: "rowHeader", row };
      }

      const x = (clientX - rect.left + el.scrollLeft) / z - ROW_HEADER_W;
      const y = (clientY - rect.top + el.scrollTop) / z - COL_HEADER_H;
      if (x < 0 || y < 0 || x >= cols.total || y >= rows.total) return null;
      const row = sheetRowOf(Math.min(rowCount - 1, indexAt(rows.offsets, y)));
      if (row < 0) return null;
      return { kind: "cell", row, col: Math.min(colCount - 1, indexAt(cols.offsets, x)) };
    },
    [cols, rows, colCount, rowCount, sheetRowOf],
  );

  /**
   * The selection a drag from `anchor` has reached, or null when it has not
   * reached anything new — which is what keeps a click that never moved a click
   * rather than a one-cell block.
   *
   * A gutter drag follows the pointer's row or column wherever it goes, the
   * sheet included: Excel keeps selecting rows after the pointer leaves the
   * gutter, and only the end it started from stays put.
   */
  const dragTo = useCallback((anchor: DragAnchor, hit: Exclude<Hit, null>): Selection | null => {
    if (anchor.kind === "cell") {
      if (hit.kind !== "cell") return null;
      if (hit.row === anchor.row && hit.col === anchor.col) return null;
      return rangeBetween(anchor, hit);
    }
    if (anchor.kind === "row") {
      const row = hit.kind === "colHeader" ? -1 : hit.row;
      return row < 0 || row === anchor.row ? null : rowsBetween(anchor.row, row);
    }
    const col = hit.kind === "rowHeader" ? -1 : hit.col;
    return col < 0 || col === anchor.col ? null : colsBetween(anchor.col, col);
  }, []);

  /**
   * Carries a drag on while the pointer is held past an edge of the viewport.
   *
   * The browser does not scroll a div for you when a drag leaves it, and the
   * pointer soon stops being over a cell at all, so without this a block can
   * never be dragged past what is on screen — Excel grows it a row at a time
   * instead. The selection is re-derived from the pointer pinned to the edge,
   * which is the row or column the block has reached.
   */
  const runAutoScroll = useCallback(() => {
    autoScroll.current = 0;
    const el = viewportRef.current;
    const anchor = rangeDrag.current;
    const point = dragPointer.current;
    if (!el || !anchor || !point) return;

    const rect = el.getBoundingClientRect();
    const z = zoomRef.current;
    const left = rect.left + ROW_HEADER_W * z;
    const top = rect.top + COL_HEADER_H * z;
    const right = rect.left + el.clientWidth;
    const bottom = rect.top + el.clientHeight;

    // How far past an edge the pointer is; zero while it is inside.
    const over = (value: number, low: number, high: number) =>
      value < low ? value - low : value > high ? value - high : 0;
    const dx = over(point.x, left, right);
    const dy = over(point.y, top, bottom);

    if (dx !== 0 || dy !== 0) {
      // The further out it is held, the faster it goes, up to about half a row
      // or column a frame.
      const speed = (delta: number) => Math.max(-40, Math.min(40, delta * 0.6));
      const beforeTop = el.scrollTop;
      const beforeLeft = el.scrollLeft;
      el.scrollLeft += speed(dx);
      el.scrollTop += speed(dy);

      if (el.scrollTop !== beforeTop || el.scrollLeft !== beforeLeft) {
        scrollRef.current = { top: el.scrollTop / z, left: el.scrollLeft / z };
        schedule("both");

        const hit = hitAt(
          Math.max(left + 1, Math.min(right - 1, point.x)),
          Math.max(top + 1, Math.min(bottom - 1, point.y)),
        );
        const next = hit ? dragTo(anchor, hit) : null;
        if (next) {
          anchor.dragged = true;
          onSelect(next);
        }
      }
    }

    autoScroll.current = requestAnimationFrame(runAutoScroll);
  }, [dragTo, hitAt, onSelect, schedule]);

  const handleClick = (event: React.MouseEvent<HTMLDivElement>) => {
    if (suppressClick.current) {
      suppressClick.current = false;
      return;
    }
    const hit = hitAt(event.clientX, event.clientY);
    if (!hit) return;
    if (hit.kind === "cell") select(hit.row, hit.col);
    else if (hit.kind === "rowHeader") onSelect(rowsBetween(hit.row, hit.row));
    else onSelect(colsBetween(hit.col, hit.col));
  };

  const handleDoubleClick = (event: React.MouseEvent<HTMLDivElement>) => {
    const hit = hitAt(event.clientX, event.clientY);
    if (hit?.kind === "colHeader") autofitColumn(hit.col);
  };

  const handlePointerDown = (event: React.PointerEvent<HTMLDivElement>) => {
    suppressClick.current = false;
    // Every shortcut in the key handler needs the viewport focused, and a press
    // on the sheet is the moment to take it back. Sheet switches give it focus
    // already, but a press that landed elsewhere — the search box, a toolbar
    // button — left Ctrl+A selecting the page's DOM instead of the sheet.
    viewportRef.current?.focus({ preventScroll: true });
    // And the page's own select-all leaves its highlight behind; nothing on the
    // sheet is a text selection, so a press clears it.
    window.getSelection()?.removeAllRanges();
    if (event.button !== 0) return;
    const hit = hitAt(event.clientX, event.clientY);
    if (!hit) return;
    // An anchor on a gutter picks a run of rows or columns; one on a cell runs a
    // block selection. Both are armed here and resolved by the closing click if
    // the pointer never moves off the anchor.
    if (hit.kind === "rowHeader") {
      rangeDrag.current = { kind: "row", row: hit.row, dragged: false };
    } else if (hit.kind === "colHeader") {
      rangeDrag.current = { kind: "column", col: hit.col, dragged: false };
    } else {
      rangeDrag.current = { kind: "cell", row: hit.row, col: hit.col, dragged: false };
    }
    dragPointer.current = { x: event.clientX, y: event.clientY };
    if (!autoScroll.current) autoScroll.current = requestAnimationFrame(runAutoScroll);
    document.body.style.userSelect = "none";
  };

  /** Dragging across cells selects the block it spans, as Excel's does. */
  useEffect(() => {
    const onMove = (event: PointerEvent) => {
      const anchor = rangeDrag.current;
      if (!anchor) return;
      // Kept whatever the pointer is over: once it leaves the viewport the
      // auto-scroll pass is the only thing that can still act on it.
      dragPointer.current = { x: event.clientX, y: event.clientY };
      const hit = hitAt(event.clientX, event.clientY);
      if (!hit) return;
      const next = dragTo(anchor, hit);
      if (!next) return;
      anchor.dragged = true;
      onSelect(next);
    };
    const onUp = () => {
      suppressClick.current = rangeDrag.current?.dragged ?? false;
      rangeDrag.current = null;
      dragPointer.current = null;
      cancelAnimationFrame(autoScroll.current);
      autoScroll.current = 0;
      document.body.style.userSelect = "";
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
    };
  }, [dragTo, hitAt, onSelect]);

  /** Sizes a column to the widest text currently loaded for it. */
  const autofitColumn = (index: number) => {
    let widest = MIN_COL_W;
    for (const [key, block] of cache.current) {
      const [name, , rawCol] = key.split("#");
      if (name !== sheet.name) continue;
      const blockColStart = Number(rawCol) * BLOCK_COLS;
      if (index < blockColStart || index >= blockColStart + BLOCK_COLS) continue;
      const local = index - blockColStart;
      for (const row of block) {
        const text = row[local]?.text;
        if (text) widest = Math.max(widest, estimateWidth(text));
      }
    }
    setColOverrides((prev) => new Map(prev).set(index, Math.min(MAX_COL_W, Math.ceil(widest))));
  };

  /** Ctrl+C: the cell's text, or the whole row/column, tab or newline joined. */
  const copySelection = useCallback(async () => {
    const sheetRows = Math.max(sheet.rows, layout?.rows ?? 0);
    const sheetCols = Math.max(sheet.cols, layout?.cols ?? 0);
    let text: string;
    let rect: MergeRect;
    try {
      if (selection.kind === "cell") {
        text = cellAt(selection.row, selection.col)?.text ?? "";
        rect = { row: selection.row, col: selection.col, rows: 1, cols: 1 };
      } else if (selection.kind === "row") {
        text = (await axisValues(path, sheet.name, { row: selection.row }, 0, sheetCols - 1)).join(
          "\t",
        );
        rect = { row: selection.row, col: 0, rows: selection.rows, cols: sheetCols };
      } else if (selection.kind === "range") {
        const block = await getCells(
          path,
          sheet.name,
          selection.row,
          selection.col,
          selection.rows,
          selection.cols,
        );
        text = block
          .map((line) => line.map((cell) => cell?.text ?? "").join("\t"))
          .join("\n");
        rect = {
          row: selection.row,
          col: selection.col,
          rows: selection.rows,
          cols: selection.cols,
        };
      } else {
        text = (await axisValues(path, sheet.name, { col: selection.col }, 0, sheetRows - 1)).join(
          "\n",
        );
        rect = { row: 0, col: selection.col, rows: sheetRows, cols: selection.cols };
      }
    } catch {
      // Nothing sensible to do if the backend is gone; leave the clipboard be.
      return;
    }
    // The marquee only follows a copy that took, so it never claims a write
    // that the clipboard refused.
    if (await copyText(text)) markCopied({ path, sheet: sheet.name, rect });
  }, [selection, cellAt, path, sheet.name, sheet.rows, sheet.cols, layout, markCopied]);

  const handleKeyDown = (event: React.KeyboardEvent) => {
    const sheetRows = Math.max(sheet.rows, layout?.rows ?? 0);
    const sheetCols = Math.max(sheet.cols, layout?.cols ?? 0);
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
      event.preventDefault();
      void copySelection();
      return;
    }

    // The cursor is free to walk out into the blank area past the data, as it
    // is in Excel, so it is clamped to what the scroller holds. Ctrl+End and
    // End still mean the sheet's own last cell, which is what `lastRow` and
    // `lastCol` are for.
    const clampRow = (r: number) => Math.min(rowCount - 1, Math.max(0, r));
    const clampCol = (c: number) => Math.min(colCount - 1, Math.max(0, c));
    const lastRow = Math.max(0, sheetRows - 1);
    const lastCol = Math.max(0, sheetCols - 1);
    const el = viewportRef.current;
    const z = zoomRef.current;
    const visible = el ? Math.floor((el.clientHeight / z - COL_HEADER_H) / baseRowH) : 20;
    const page = Math.max(1, visible - 2 * OVERSCAN_ROWS);
    const cursor = selectionAnchor(selection);

    // Ctrl belongs to the sheet rather than the cursor: Excel jumps to the two
    // corners with Ctrl+Home/End and steps between sheets with Ctrl+PageUp/Down.
    // Ahead of the row and column branches, so a whole-row selection answers to
    // them too — those branches return early for anything but an arrow key.
    if (event.ctrlKey || event.metaKey) {
      // Ctrl+A is not handled here: it belongs to the sheet wherever the focus
      // happens to be, so it lives on `window` below.
      switch (event.key) {
        case "Home":
          select(0, 0);
          break;
        case "End":
          select(lastRow, lastCol);
          break;
        case "PageUp":
          onSheetStep?.(-1);
          break;
        case "PageDown":
          onSheetStep?.(1);
          break;
        // Excel's jump to the edge of the data, which the whole-row and
        // whole-column branches below never see — they answer this branch's
        // arrows as an ordinary one-cell move.
        case "ArrowUp":
        case "ArrowDown":
        case "ArrowLeft":
        case "ArrowRight": {
          const dir =
            event.key === "ArrowUp"
              ? "up"
              : event.key === "ArrowDown"
                ? "down"
                : event.key === "ArrowLeft"
                  ? "left"
                  : "right";
          void jumpToEdge(dir, event.shiftKey);
          break;
        }
        default:
          return;
      }
      event.preventDefault();
      return;
    }

    // A run of whole rows keeps its own axis: the arrows slide it, Shift grows
    // it from the end that was walking, and Home/End collapse it to a cell on
    // the active row.
    //
    // The slide's step has to keep the *whole* run on the sheet. Clamping the
    // top alone let 27:29 slide onto row 31 of a 30-row sheet.
    if (selection.kind === "row") {
      const slide = (next: number) => {
        const step = Math.max(
          -selection.row,
          Math.min(rowCount - selection.row - selection.rows, next - selection.row),
        );
        onSelect(rowsBetween(selection.anchor + step, selection.active + step));
      };
      const grow = (next: number) =>
        onSelect(rowsBetween(selection.anchor, clampRow(next)));
      switch (event.key) {
        case "ArrowUp":
          if (event.shiftKey) grow(selection.active - 1);
          else slide(selection.row - 1);
          break;
        case "ArrowDown":
          if (event.shiftKey) grow(selection.active + 1);
          else slide(selection.row + 1);
          break;
        case "PageUp":
          if (event.shiftKey) grow(selection.active - page);
          else slide(selection.row - page);
          break;
        case "PageDown":
          if (event.shiftKey) grow(selection.active + page);
          else slide(selection.row + page);
          break;
        case "Home":
          select(selection.active, 0);
          break;
        case "End":
          select(selection.active, lastCol);
          break;
        default:
          return;
      }
      event.preventDefault();
      return;
    }
    if (selection.kind === "column") {
      const slide = (next: number) => {
        const step = Math.max(
          -selection.col,
          Math.min(colCount - selection.col - selection.cols, next - selection.col),
        );
        onSelect(colsBetween(selection.anchor + step, selection.active + step));
      };
      const grow = (next: number) =>
        onSelect(colsBetween(selection.anchor, clampCol(next)));
      switch (event.key) {
        case "ArrowLeft":
          if (event.shiftKey) grow(selection.active - 1);
          else slide(selection.col - 1);
          break;
        case "ArrowRight":
          if (event.shiftKey) grow(selection.active + 1);
          else slide(selection.col + 1);
          break;
        // A column has no vertical page of its own to move by; PageUp and
        // PageDown scroll, which the wheel already does.
        case "Home":
          select(0, selection.active);
          break;
        case "End":
          select(lastRow, selection.active);
          break;
        default:
          return;
      }
      event.preventDefault();
      return;
    }

    // Shift extends a block rather than collapsing to a cell. The far end walks
    // from where it already is — `cursor` is a block's active cell, which is
    // that end — so repeated presses accumulate and a press back retraces.
    const moveTo = (row: number, col: number) => {
      const target = { row: clampRow(row), col: clampCol(col) };
      if (!event.shiftKey) {
        select(target.row, target.col);
        return;
      }
      const anchor = selection.kind === "range" ? selection.anchor : cursor;
      onSelect(rangeBetween(anchor, target));
    };

    switch (event.key) {
      case "ArrowUp":
        moveTo(cursor.row - 1, cursor.col);
        break;
      case "ArrowDown":
        moveTo(cursor.row + 1, cursor.col);
        break;
      case "ArrowLeft":
        moveTo(cursor.row, cursor.col - 1);
        break;
      case "ArrowRight":
        moveTo(cursor.row, cursor.col + 1);
        break;
      case "PageUp":
        moveTo(cursor.row - page, cursor.col);
        break;
      case "PageDown":
        moveTo(cursor.row + page, cursor.col);
        break;
      case "Home":
        moveTo(cursor.row, 0);
        break;
      case "End":
        moveTo(cursor.row, lastCol);
        break;
      default:
        return;
    }
    event.preventDefault();
  };

  // --- resize by dragging a gutter boundary ---
  useEffect(() => {
    const onMove = (event: PointerEvent) => {
      const active = drag.current;
      if (!active) return;
      const moved =
        (active.kind === "col" ? event.clientX - active.start : event.clientY - active.start) /
        zoomRef.current;
      if (active.kind === "col") {
        const next = Math.max(MIN_COL_W, Math.min(MAX_COL_W, Math.round(active.size + moved)));
        setColOverrides((prev) => new Map(prev).set(active.index, next));
      } else {
        const next = Math.max(minRowH, Math.round(active.size + moved));
        setRowOverrides((prev) => new Map(prev).set(active.index, next));
      }
    };
    const onUp = () => {
      drag.current = null;
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
    };
  }, [minRowH]);

  /** The boundary a pointer is within grabbing distance of, or null. */
  const resizeTargetAt = useCallback(
    (clientX: number, clientY: number): { kind: "col" | "row"; index: number } | null => {
      const el = viewportRef.current;
      if (!el) return null;
      const rect = el.getBoundingClientRect();
      const z = zoomRef.current;
      const px = clientX - rect.left;
      const py = clientY - rect.top;
      // The grab distance is in screen pixels, the offsets are in the sheet's.
      const grab = RESIZE_HANDLE / 2 / z;

      if (py < COL_HEADER_H * z && px >= ROW_HEADER_W * z) {
        const x = (px + el.scrollLeft) / z - ROW_HEADER_W;
        // The boundary under the pointer is the trailing edge of the column it
        // is in, or the leading edge of that column — which is the previous
        // column's trailing one. Two comparisons, not a scan of every column.
        const col = Math.min(sheetColsGuess - 1, indexAt(cols.offsets, x));
        if (Math.abs(cols.offsets[col + 1] - x) <= grab) return { kind: "col", index: col };
        if (col > 0 && Math.abs(cols.offsets[col] - x) <= grab)
          return { kind: "col", index: col - 1 };
        return null;
      }
      if (px < ROW_HEADER_W * z && py >= COL_HEADER_H * z) {
        const y = (py + el.scrollTop) / z - COL_HEADER_H;
        const display = Math.min(displayRowCount - 1, indexAt(rows.offsets, y));
        const rowAt = (index: number) => (index < 0 ? -1 : sheetRowOf(index));
        if (Math.abs(rows.offsets[display + 1] - y) <= grab) {
          const row = rowAt(display);
          if (row >= 0) return { kind: "row", index: row };
        }
        if (display > 0 && Math.abs(rows.offsets[display] - y) <= grab) {
          const row = rowAt(display - 1);
          if (row >= 0) return { kind: "row", index: row };
        }
        return null;
      }
      return null;
    },
    [cols, rows, sheetColsGuess, sheetRowOf, displayRowCount],
  );

  const handlePointerDownCapture = (event: React.PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) return;
    const target = resizeTargetAt(event.clientX, event.clientY);
    if (!target) return;
    event.preventDefault();
    event.stopPropagation();
    drag.current = {
      kind: target.kind,
      index: target.index,
      start: target.kind === "col" ? event.clientX : event.clientY,
      size:
        target.kind === "col"
          ? spanAt(cols.offsets, target.index)
          : spanAt(rows.offsets, displayOfRow(target.index)),
    };
    event.currentTarget.setPointerCapture(event.pointerId);
  };

  /**
   * The cell's text, but only if it does not fit — canvas has no native
   * `title`, so the DOM renderer's hover readout has to be built here. Only a
   * truncated cell gets one: "the value you cannot see" is the question a hover
   * answers, and reproducing a title on every cell would mean hit-testing and
   * a tooltip for a string that is already on screen.
   */
  const truncatedTextAt = (row: number, col: number): string | null => {
    const text = cellAt(row, col)?.text ?? "";
    if (text === "") return null;
    const style = styleAt(row, col);
    const sizePx = style?.sizePx ?? fontSize;
    const font = fontString(
      cellFontStack(style?.family ?? null, fontFamily),
      sizePx,
      style?.bold ?? false,
    );
    const available = spanAt(cols.offsets, col) - CELL_PAD_LEFT - CELL_PAD_RIGHT;
    if (available <= 0) return text;
    if (style?.wrap) {
      const display = displayOfRow(row);
      if (display < 0) return null;
      const needed = layoutLines(font, text, available, true).length * sizePx * LINE_HEIGHT;
      return needed > spanAt(rows.offsets, display) ? text : null;
    }
    return textWidth(font, text) > available ? text : null;
  };

  const updateTooltip = (hit: ReturnType<typeof hitAt>) => {
    const key = hit?.kind === "cell" ? `${hit.row}:${hit.col}` : null;
    // Moving within one cell must not touch React at all.
    if (key === hoverKey.current) return;
    hoverKey.current = key;
    const el = viewportRef.current;
    if (!el || !hit || hit.kind !== "cell") {
      setTip(null);
      return;
    }
    const text = truncatedTextAt(hit.row, hit.col);
    const display = displayOfRow(hit.row);
    if (!text || display < 0) {
      setTip(null);
      return;
    }
    const rect = el.getBoundingClientRect();
    const z = zoomRef.current;
    setTip({
      text,
      x: rect.left + (ROW_HEADER_W + cols.offsets[hit.col] - scrollRef.current.left) * z,
      y:
        rect.top +
        (COL_HEADER_H + rows.offsets[display] - scrollRef.current.top) * z +
        spanAt(rows.offsets, display) * z,
    });
  };

  const handlePointerMove = (event: React.PointerEvent<HTMLDivElement>) => {
    if (drag.current) return;
    const el = viewportRef.current;
    if (!el) return;
    const target = resizeTargetAt(event.clientX, event.clientY);
    el.style.cursor = target ? (target.kind === "col" ? "col-resize" : "row-resize") : "";
    updateTooltip(target ? null : hitAt(event.clientX, event.clientY));
  };

  const handlePointerLeave = () => {
    hoverKey.current = null;
    setTip(null);
  };

  /** Whether the hidden readout is on; Ctrl+Shift+D flips it. */
  const readoutOn = useSyncExternalStore(subscribeShapeDebug, shapeDebugEnabled);
  readoutOnRef.current = readoutOn;

  /**
   * What the hidden readout prints, for the panel `debugPanels.ts` toggles.
   *
   * Gathered here because this is where the layout is known: `styleAt` applies
   * the same cell → row → column rule the drawing does, so what is printed is
   * what is on screen.
   *
   * `import.meta.env.DEV` first, so a packaged build drops the whole block, and
   * the toggle second, so it is not built on every selection change while
   * nobody is looking at it.
   */
  const debugFacts: ShapeDebugFacts | null = useMemo(() => {
    if (!import.meta.env.DEV || !readoutOn) return null;

    const { row, col } = selectionAnchor(selection);
    const display = displayOfRow(row);
    const style = styleAt(row, col);
    const value = cellAt(row, col);
    const edge = (found: BorderEdge | null | undefined) =>
      found ? `${found.line} ${found.width}px ${found.color ?? "auto"} r${found.rank}` : "none";
    const merge = merges.find(
      (m) => row >= m.row && row < m.row + m.rows && col >= m.col && col < m.col + m.cols,
    );
    const cellReadout = [
      `${cellRef(row, col)}   ${spanAt(cols.offsets, col)}×${spanAt(rows.offsets, display)} @ (${cols.offsets[col] ?? "?"}, ${rows.offsets[display] ?? "?"})`,
      `  value: ${value ? `${value.kind} ${JSON.stringify(value.text)}` : "not loaded"}`,
      style
        ? `  style: font ${style.family ?? "(workbook)"} ${style.sizePx ?? fontSize}px${style.bold ? " bold" : ""}  colour ${style.color ?? "auto"}  fill ${style.fill ?? "none"}  align ${style.horizontal ?? "auto"}/${style.vertical ?? "bottom"}  indent ${style.indent}  wrap ${style.wrap ? "yes" : "no"}`
        : "  style: none",
      `  borders: top ${edge(style?.border?.top)}  right ${edge(style?.border?.right)}  bottom ${edge(style?.border?.bottom)}  left ${edge(style?.border?.left)}`,
      `  merge: ${merge ? `r${merge.row} c${merge.col} ${merge.rows}×${merge.cols}` : "no"}   outside print area: ${outsidePrint(row, col) ? "yes" : "no"}`,
    ].join("\n");

    return {
      sheet: sheet.name,
      rows: sheetRowsGuess,
      cols: sheetColsGuess,
      fontFamily,
      fontSizePx: fontSize,
      digitWidth: digitWidthPx(fontFamily, fontSize),
      defaultColChars: layout?.defaultColWidthChars ?? null,
      defaultColPx: baseColW,
      defaultRowPx: baseRowH,
      showGrid,
      merges: merges.length,
      colWidths: colOverrides,
      printArea,
      cellReadout,
      timings: describePaint(perf.current),
    };
  }, [
    readoutOn,
    selection,
    displayOfRow,
    styleAt,
    cellAt,
    merges,
    cols,
    rows,
    outsidePrint,
    fontSize,
    fontFamily,
    sheet.name,
    sheetRowsGuess,
    sheetColsGuess,
    layout,
    baseColW,
    baseRowH,
    showGrid,
    colOverrides,
    printArea,
    // The paint counters live in a ref, so the memo needs this to notice them.
    readoutTick,
  ]);

  const contentW = (ROW_HEADER_W + cols.total) * zoomRef.current;
  const contentH = (COL_HEADER_H + rows.total) * zoomRef.current;

  return (
    <div className="flex h-full min-h-0 flex-col">
      {autoFilter && visibleRows && (
        <div className="flex shrink-0 items-center gap-2 border-b border-border bg-accent/40 px-2 py-1 text-[11px]">
          <ListFilter className="size-3 shrink-0" />
          <span className="tabular-nums">
            Showing {visibleRows.length} of {autoFilter.rows - 1} rows
          </span>
          <Button
            size="sm"
            variant="ghost"
            className="h-5 px-1.5 text-[11px]"
            onClick={() => {
              setFilters(new Map());
              rememberFilters(new Map());
              setVisibleRows(null);
            }}
          >
            Clear filters
          </Button>
        </div>
      )}

      {bar}

      <div className="relative min-h-0 flex-1 overflow-hidden">
        <div
          ref={viewportRef}
          tabIndex={0}
          onScroll={handleScroll}
          onClick={handleClick}
          onDoubleClick={handleDoubleClick}
          onPointerDownCapture={handlePointerDownCapture}
          onPointerDown={handlePointerDown}
          onPointerMove={handlePointerMove}
          onPointerLeave={handlePointerLeave}
          onKeyDown={handleKeyDown}
          style={hBarHeight ? { marginBottom: -hBarHeight } : undefined}
          className="absolute inset-0 overflow-auto outline-none [scrollbar-gutter:stable]"
        >
          <div style={{ width: contentW, height: contentH }} />
        </div>

        {/* The cell layer: painted, and never in the way of a pointer. */}
        <canvas ref={baseRef} className="pointer-events-none absolute top-0 left-0" />
        {/* The selection layer, stacked over it. */}
        <canvas ref={overlayRef} className="pointer-events-none absolute top-0 left-0" />

        {/*
          The filter buttons. They are React, so they cannot be painted; they
          are moved to their cells by hand in the paint pass and clipped to the
          cell area, so one never floats over a gutter. Only the button itself
          takes a pointer, so the cell underneath still selects.
        */}
        <div
          ref={overlayAreaRef}
          className="pointer-events-none absolute top-0 left-0 overflow-hidden"
        >
          {autoFilter &&
            filterCols.map((col) => (
              <div
                key={col}
                ref={(element) => {
                  if (element) filterSlots.current.set(col, element);
                  else filterSlots.current.delete(col);
                }}
                className="pointer-events-none absolute top-0 left-0"
              >
                <FilterButton
                  path={path}
                  sheet={sheet.name}
                  col={col}
                  firstRow={autoFilter.row + 1}
                  lastRow={autoFilter.row + autoFilter.rows - 1}
                  selected={filters.get(col) ?? []}
                  onApply={(values) => {
                    const next = new Map(filters);
                    if (values && values.length > 0) next.set(col, values);
                    else next.delete(col);
                    void applyFilters(next, autoFilter);
                  }}
                />
              </div>
            ))}
        </div>

        {/*
          The sheet's shapes and the print-area mark. They sit above the cells
          and below the gutters, in the same coordinate space the cells were
          painted in — which the canvas does not scroll, so this layer is
          translated and scaled with the scroll instead of scrolling. Clipping
          it to the cell area is what keeps a shape anchored at row 0 from
          painting over the column headers.
        */}
        {layout && layoutReady && (
          <div
            ref={pageLayerRef}
            className="pointer-events-none absolute top-0 left-0 overflow-hidden"
            /*
             * A shape's text names no font of its own in the file, so it renders
             * in whatever it inherits — and this layer is a sibling of the
             * scroller rather than a child of it, so without this declaration it
             * inherits the app's UI font instead of the workbook's. That changes
             * every glyph and, on a line that sits near the box's width, where
             * the line breaks. The scroller that used to hold this layer carried
             * the same three declarations.
             */
            style={{
              fontFamily: fontFamily
                ? `"${fontFamily}", ${CJK_FALLBACK}, system-ui, sans-serif`
                : undefined,
              fontSize,
              lineHeight: LINE_HEIGHT,
            }}
          >
            <div
              ref={shapesInnerRef}
              className="absolute top-0 left-0"
              style={{ transformOrigin: "0 0" }}
            >
              <ShapeLayer
                shapes={layout.drawings}
                colX={shapeColX}
                rowY={shapeRowY}
                // The layer is clipped to the sheet's own extent, not to the
                // scroller. An `<svg>` hides what falls outside it, which is
                // what keeps a picture wider than the sheet's last column — or
                // a shape whose cell the parser could only approximate — from
                // spilling out into the blank area the scroller runs on into;
                // sized from the scroller, every one of them showed again.
                width={usedColsTotal}
                height={usedRowsTotal}
                path={path}
                // Null unless the readout is on, so the layer's `memo` still
                // holds in normal use — the facts change with the selection.
                debug={debugFacts}
              />
            </div>
            <div
              ref={markRef}
              aria-hidden
              className="pointer-events-none absolute"
              style={{
                display: "none",
                border: `${PRINT_AREA_LINE}px solid ${PRINT_AREA_INK}`,
              }}
            />
          </div>
        )}

        {tip && (
          <div
            className="pointer-events-none fixed z-50 max-w-80 border border-border bg-background px-1.5 py-1 text-[11px] break-words whitespace-pre-wrap text-foreground shadow-md"
            style={{ left: tip.x, top: tip.y }}
          >
            {tip.text}
          </div>
        )}
      </div>
    </div>
  );
}
