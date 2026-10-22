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
import { Eye, EyeOff } from "lucide-react";

import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuSeparator,
  ContextMenuSub,
  ContextMenuSubContent,
  ContextMenuSubTrigger,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";
import { FilterButton } from "@/components/FilterButton";
import { ColorPalette } from "@/components/ColorPalette";
import { Accelerated, acceleratorOf, pickByAccelerator } from "@/components/Accelerated";
import { ShapeLayer, type ShapeDebugFacts } from "@/components/ShapeLayer";
import { canvasTheme } from "@/lib/canvasTheme";
import { useTranslation } from "react-i18next";
import {
  paintBase,
  paintOverlay,
  valueWidth,
  type GridView,
  type PaintStats,
} from "@/lib/canvasPaint";
import { validationListItems } from "@/lib/validationList";
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
  type DataBarCell,
  type FilterCriterion,
  type IconCell,
  type MergeRect,
  type NoteCell,
  type RichRun,
  type SheetInfo,
  type SheetLayout,
} from "@/lib/excel";
import {
  ACCENT_INK,
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
  OUTSIDE_PRINT,
  PRINT_AREA_INSIDE,
  PRINT_AREA_INK,
  PRINT_AREA_LINE,
  RESIZE_HANDLE,
  RING_W,
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
import { MAX_ZOOM, MIN_ZOOM, stepZoom } from "@/lib/zoom";
import { paginate } from "@/lib/pagination";
import {
  activeArea,
  areasOf,
  colsBetween,
  rangeBetween,
  rangeOver,
  rowsBetween,
  selectionAnchor,
  selectionCovers,
  toggleArea,
  withArea,
  type Area,
  type Selection,
} from "@/lib/selection";
import { markKey, type Marks } from "@/lib/marks";
import {
  colLevelButtonRect,
  groupButtonBox,
  groupOpen,
  groupToggleLines,
  outlineGutter,
  rowLevelButtonRect,
  colOutline as buildColOutline,
  rowOutline as buildRowOutline,
  type OutlineGroup,
} from "@/lib/outline";

/** A separator no path and no sheet name can hold, so a key can be split on it. */
const SEP = "\u0000";

/**
 * Names one sheet of one workbook.
 *
 * A sheet name alone does not say which workbook it belongs to, so anything
 * that remembers something about a sheet has to carry the workbook too. A name
 * alone is what let a second file read the first one's cells.
 */
const sheetIdentity = (path: string, sheet: string) => `${path}${SEP}${sheet}`;

/** The key one block of cells is held under, and the prefix that claims a sheet's. */
const blockPrefix = (path: string, sheet: string) => `${sheetIdentity(path, sheet)}${SEP}`;

/** Where a pointer landed: a cell, one of the two gutter strips, or the
 *  select-all square where they meet. */
type Hit =
  | { kind: "cell"; row: number; col: number }
  | { kind: "rowHeader"; row: number }
  | { kind: "colHeader"; col: number }
  | { kind: "corner" }
  | null;

/** What a drag-select is anchored on, and whether it has moved off it yet. */
type DragAnchor =
  | { kind: "cell"; row: number; col: number; dragged: boolean }
  | { kind: "row"; row: number; dragged: boolean }
  | { kind: "column"; col: number; dragged: boolean };

/** A press on Excel's outline chrome: a group's button or a corner level. */
type OutlineAction =
  | { kind: "rowGroup"; group: OutlineGroup }
  | { kind: "colGroup"; group: OutlineGroup }
  | { kind: "rowLevel"; level: number }
  | { kind: "colLevel"; level: number };

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
  /**
   * The column widths and row heights the reader has dragged or fitted here.
   *
   * The file's own sizes are only where a sheet starts; what has since been
   * done to it belongs to the sheet for as long as the workbook is open, so
   * leaving one — for another sheet, for another workbook and back — returns it
   * as it was left. Closing the workbook is what puts the file's own back.
   */
  sizes?: { col: Map<number, number>; row: Map<number, number> };
  /**
   * The rows and columns the reader has hidden or shown.
   *
   * The same kind of state as the sizes above, and for the same reason: nothing
   * is written to the workbook, so leaving a sheet — for another one, or for
   * another workbook and back — has to return it as it was left, and closing
   * the workbook is what puts the file's own answer back.
   */
  hidden?: { rows: Map<number, boolean>; cols: Map<number, boolean> };
}

/** Those positions, by workbook path and then by sheet name. */
export type SheetViews = Map<string, Map<string, SheetView>>;

/** The heading a context menu was opened on, or none when it was a cell. */
type MenuTarget = { kind: "row"; index: number } | { kind: "col"; index: number } | null;

/** Excel's input prompt for the active cell, from the file's validation rule. */
interface PromptInfo {
  row: number;
  col: number;
  title: string;
  text: string;
}

/**
 * The rows — or the columns — an area covers, or null when it runs the other
 * way. A column reaches every row there is and a row every column, so neither
 * can say which heading the other axis should act on.
 */
function spanOfArea(area: Area, rows: boolean): [number, number] | null {
  switch (area.kind) {
    case "cell":
      return rows ? [area.row, area.row] : [area.col, area.col];
    case "row":
      return rows ? [area.row, area.row + area.rows - 1] : null;
    case "column":
      return rows ? null : [area.col, area.col + area.cols - 1];
    case "range":
      return rows
        ? [area.row, area.row + area.rows - 1]
        : [area.col, area.col + area.cols - 1];
  }
}

interface GridProps {
  path: string;
  sheet: SheetInfo;
  /**
   * Bumped when the workbook has been read again from disk. The path is the
   * same file either way, so this is what tells the grid that what it holds —
   * the layout, and the blocks of cells fetched under this path — is from the
   * workbook as it used to be.
   */
  revision: number;
  selection: Selection;
  /**
   * Where each sheet of each open workbook was left scrolled, keyed by
   * workbook and then by sheet. The app owns it rather than the grid, so that
   * closing a tab forgets its sheets' positions — the same way it forgets
   * their cursors.
   */
  views: SheetViews;
  onSelect: (selection: Selection) => void;
  /**
   * The review marks this workbook carries, keyed by sheet and cell. They are
   * the app's own, so closing the tab is what forgets them.
   */
  marks: Marks;
  /** Shades the selection, or lifts the shading off it when `fill` is null. */
  onMark: (fill: string | null) => void;
  /** The colour the toolbar's bucket is holding, so the palette can ring it. */
  markColor: string;
  onDetail?: (cell: Cell | null) => void;
  /** The extent actually on screen, which can exceed calamine's range. */
  onExtent?: (rows: number, cols: number) => void;
  /**
   * How many rows the sheet's filter is hiding, or null when it is hiding none.
   * The status bar says it, which is where Excel says it too.
   */
  onFiltered?: (filtered: { shown: number; total: number } | null) => void;
  /** Steps to the sheet `delta` along in the workbook, for Ctrl+PageUp/Down. */
  onSheetStep?: (delta: number) => void;
  /**
   * Bumped when the caller places the cursor somewhere it has to be *seen* — Go
   * To, or a search hit. Arriving on a sheet is not one of those: Excel puts
   * back the view that sheet was left at, off-screen cursor and all, so a jump
   * is the only thing allowed to drag it.
   */
  reveal?: number;
  /**
   * A note the reader is pointing at somewhere else — the notes sidebar — to be
   * shown in the sheet as if the cell under it were hovered. Null hides it.
   */
  previewNote?: { row: number; col: number } | null;
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
   * Excel's Page Break Preview rather than its ordinary view. True shades
   * around the print area and boxes it; false leaves the sheet drawn with no
   * page at all, which is where a workbook that never set a print area belongs.
   */
  pageBreak: boolean;
  /**
   * The element the sheet scrolls within, handed out so a scrollbar drawn
   * elsewhere is a view of it rather than a second scroller.
   */
  onViewport?: (element: HTMLDivElement | null) => void;
}

interface ResizeDrag {
  kind: "col" | "row";
  index: number;
  /** The rows the grip acts on: a run when the selection covers it, else one. */
  rows: number[];
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
 * Where Excel shows a note's box on hover, in screen pixels at 100%: just
 * outside the cell's top-right corner, 11.25pt across and 9pt up (read off
 * COM while the note was displayed). The VML shape supplies the size and, for
 * other viewers, the stored position; this is where Excel puts it.
 */
const NOTE_BOX_GAP_PX = 15;
const NOTE_BOX_TOP_GAP_PX = 12;

/**
 * How long the pointer rests on a note's cell before the box opens, and how
 * long it may leave before the box closes. The same beat either way: short
 * enough to feel immediate, long enough to move between a cell and its box
 * without the box blinking shut.
 */
const NOTE_HOVER_MS = 300;

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
  revision,
  selection,
  views,
  onSelect,
  marks,
  onMark,
  markColor,
  onDetail,
  onExtent,
  onFiltered,
  onSheetStep,
  reveal,
  previewNote,
  zoom,
  onZoom,
  bar,
  hideHScroll,
  pageBreak,
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
  /** The data-validation input prompt, when the active cell carries one. */
  const promptRef = useRef<HTMLDivElement>(null);
  const promptInfoRef = useRef<PromptInfo | null>(null);
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
   *
   * A field being typed in keeps it, though, and so does anything inside the
   * search panel: a jump to a hit on another sheet is a sheet change like any
   * other, and taking the keyboard off the results would end the walk on the
   * first hit that happened to be on another sheet — the next press of the key
   * would go to the sheet instead, and move the cell rather than the walk. The
   * whole panel is meant, not just its box: the list is walked from wherever in
   * it the focus happens to be, which after a click on a hit is that hit.
   */
  useEffect(() => {
    const active = document.activeElement;
    if (
      active instanceof HTMLElement &&
      active.closest("input, textarea, [contenteditable], [data-search-panel]")
    ) {
      return;
    }
    viewportRef.current?.focus({ preventScroll: true });
    window.getSelection()?.removeAllRanges();
  }, [path, sheet.name]);

  const cache = useRef(new Map<string, Cell[][]>());
  /** The revision the blocks in `cache` were read at, for a workbook re-read. */
  const cachedRevision = useRef(revision);
  const pending = useRef(new Set<string>());
  const drag = useRef<ResizeDrag | null>(null);
  /** Which workbook and sheet the sizes in state belong to; see the effect on them. */
  const seeded = useRef<string | null>(null);
  /**
   * The cell the file's own view starts at, waiting for the metrics that turn
   * it into a scroll offset. A sheet with no view of its own in this session
   * opens where Excel left it, not at A1.
   */
  const pendingView = useRef<[number, number] | null>(null);
  /**
   * A drag-select in progress, and whether it has moved.
   *
   * The anchor is a cell, a whole row or a whole column, because a drag that
   * starts on a gutter picks rows or columns — dragging one row header down is
   * how a run like `3:7` is made by mouse, and only that end moves.
   */
  const rangeDrag = useRef<DragAnchor | null>(null);
  /**
   * The selection a Ctrl+drag is adding to, if that is what this drag is.
   *
   * Held here rather than read as it goes, because each move redraws the block
   * being pulled out from scratch: adding to the selection as it stands would
   * pile up one area per pointer event.
   */
  const dragBase = useRef<Selection | null>(null);
  /** Where the pointer last was, which a drag keeps reading after it leaves. */
  const dragPointer = useRef<{ x: number; y: number } | null>(null);
  /** The frame that carries a drag on past the edge of the viewport. */
  const autoScroll = useRef(0);
  /** Set when such a drag ends, so its closing click cannot collapse the block. */
  const suppressClick = useRef(false);
  /** Set by a press on an outline button, whose closing click is not a select. */
  const outlinePress = useRef(false);

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
  /** Which workbook's sheet the layout in state belongs to — see `sheetIdentity`. */
  const [layoutSheet, setLayoutSheet] = useState<string | null>(null);
  const [fontFamily, setFontFamily] = useState<string | null>(null);
  const [fontSize, setFontSize] = useState(12);
  const [showGrid, setShowGrid] = useState(true);
  const [filters, setFilters] = useState<Map<number, string[]>>(() => new Map());
  /**
   * The cell whose text is open, if any. Double-clicking one lifts its text into
   * a field over it, because painted text cannot be selected and copied with the
   * mouse and this can. The field is read-only: there is nowhere for an edit to
   * go, and offering one that could not be kept would be a lie.
   */
  const [editing, setEditing] = useState<{ row: number; col: number } | null>(null);
  /**
   * The list a list-validated cell offers, opened from the arrow Excel draws
   * on it. Fixed positioning, because the sheet it belongs to scrolls under
   * it; it closes rather than following a scroll.
   */
  const [valueList, setValueList] = useState<{
    x: number;
    y: number;
    items: string[];
  } | null>(null);
  const editorRef = useRef<HTMLDivElement>(null);
  /** What the paint pass reads, since a scroll does not re-render React. */
  const editingRef = useRef(editing);
  editingRef.current = editing;
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
   * The note under the pointer, in screen coordinates, placed beside its cell
   * as Excel shows it. `noteTipRef` mirrors it so the pointer can be tested
   * against the open box without a re-render.
   */
  const [noteTip, setNoteTip] = useState<{
    note: NoteCell;
    x: number;
    y: number;
    w: number;
    h: number;
    /** The leader line: the box's corner it starts at, and the cell's it points to. */
    from: { x: number; y: number };
    to: { x: number; y: number };
    /**
     * Where the pointer keeps this note open: the box and the strip between it
     * and the cell. Without the strip, moving from the cell onto the box
     * crosses the neighbour's cells — which, if one of them has a note of its
     * own, would open that one instead.
     */
    guard: { x0: number; y0: number; x1: number; y1: number };
  } | null>(null);
  const noteTipRef = useRef<typeof noteTip>(null);
  const showNote = (value: typeof noteTip) => {
    noteTipRef.current = value;
    setNoteTip(value);
  };
  /**
   * A note opens after a short beat rather than the instant the pointer
   * crosses a cell, and closes on the same beat after the pointer leaves —
   * long enough to move from the cell onto its box, which floats just outside
   * the cell. Being over the box cancels the close; leaving it starts the
   * grace again.
   */
  const noteShowTimer = useRef(0);
  const noteHideTimer = useRef(0);
  const cancelNoteHide = () => {
    window.clearTimeout(noteHideTimer.current);
    noteHideTimer.current = 0;
  };
  const scheduleNoteShow = (note: NoteCell) => {
    cancelNoteHide();
    if (noteTipRef.current?.note === note) return;
    window.clearTimeout(noteShowTimer.current);
    noteShowTimer.current = window.setTimeout(() => {
      noteShowTimer.current = 0;
      showNoteFor(note);
    }, NOTE_HOVER_MS);
  };
  const scheduleNoteHide = () => {
    window.clearTimeout(noteShowTimer.current);
    noteShowTimer.current = 0;
    if (!noteTipRef.current) return;
    cancelNoteHide();
    noteHideTimer.current = window.setTimeout(() => {
      noteHideTimer.current = 0;
      showNote(null);
    }, NOTE_HOVER_MS);
  };
  useEffect(
    () => () => {
      window.clearTimeout(noteShowTimer.current);
      window.clearTimeout(noteHideTimer.current);
    },
    [],
  );
  /**
   * What Ctrl+C took, and which sheet it came from.
   *
   * Excel's copy is otherwise silent — no dialog, no flash — so this marquee is
   * the only sign it worked, and it stays until Escape or the next copy, on the
   * sheet it came from rather than on whichever sheet is showing.
   */
  const copiedRef = useRef<{ path: string; sheet: string; rects: MergeRect[] } | null>(null);
  /** How far the marquee's dashes have crawled, walked by the timer below. */
  const antsRef = useRef(0);
  const antsTimer = useRef(0);

  const layoutReady = layout !== null && layoutSheet === sheetIdentity(path, sheet.name);

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
   * the sheet unselected, and leaves that highlight behind afterwards.
   *
   * While the field over a cell is open it is the thing being worked on, and
   * Ctrl+A belongs to it — every time it is pressed, as it is in Excel, rather
   * than walking out to the region and then to the whole sheet. Selecting its
   * text here rather than leaving it to the browser is what makes that hold
   * whether or not the field happens to be the thing holding the focus: a field
   * that lost it would otherwise be answering this as the sheet.
   */
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== "a") return;
      const field = editorRef.current;
      if (editingRef.current && field) {
        event.preventDefault();
        field.focus({ preventScroll: true });
        const range = document.createRange();
        range.selectNodeContents(field);
        const selection = window.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(range);
        return;
      }
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
  /**
   * Columns the file filters with a condition rather than a value list. Their
   * hidden rows are already the file's own, so only the button's funnel and the
   * blue row numbers depend on this.
   */
  const customFilterCols = useMemo(
    () => new Set(autoFilter?.custom ?? []),
    [autoFilter],
  );
  /** The notes by cell, for the mark's paint and the hover box. */
  const notesByCell = useMemo(() => {
    const map = new Map<number, NoteCell>();
    for (const note of layout?.notes ?? []) {
      map.set(cellKey(sheetColsGuess, note.row, note.col), note);
    }
    return map;
  }, [layout, sheetColsGuess]);

  /**
   * Where shapes have been dragged to, by sheet — session state only. A shape
   * that covers the cells under it can be moved aside, and nothing is written
   * back, so reopening the file puts everything back where it was.
   */
  const [shapeMoves, setShapeMoves] = useState<
    Map<string, Map<number, { x: number; y: number }>>
  >(() => new Map());
  const shapeMovesKey = `${path}|${sheet.name}`;
  const shapeOffsets = shapeMoves.get(shapeMovesKey);
  const moveShape = useCallback(
    (index: number, x: number, y: number) => {
      setShapeMoves((current) => {
        const next = new Map(current);
        const sheetOffsets = new Map(next.get(shapeMovesKey) ?? []);
        sheetOffsets.set(index, { x, y });
        next.set(shapeMovesKey, sheetOffsets);
        return next;
      });
    },
    [shapeMovesKey],
  );

  /**
   * A note's box in sheet pixels. The VML anchors it in cells, so it follows
   * the grid this viewer measured rather than Excel's own column widths.
   */
  const noteBox = (note: NoteCell) => {
    const [c1, dx1, r1, dy1, c2, dx2, r2, dy2] = note.anchor;
    const top = displayOfRow(r1);
    const bottom = displayOfRow(r2);
    if (top < 0 || bottom < 0) return null;
    const x = (cols.offsets[c1] ?? 0) + dx1;
    const y = rows.offsets[top] + dy1;
    const right = (cols.offsets[c2] ?? x) + dx2;
    const bottomY = rows.offsets[bottom] + dy2;
    if (!Number.isFinite(right) || !Number.isFinite(bottomY)) return null;
    if (right <= x || bottomY <= y) return null;
    return { x, y, w: right - x, h: bottomY - y };
  };
  /**
   * The page the shading and the box are drawn around — and nothing at all in
   * the ordinary view, which is the whole of the difference between Excel's two
   * views here. `outsidePrint` below already answers "no" whenever this is null,
   * and the box over the page is the same question asked of the same value.
   */
  const printArea = pageBreak ? (layout?.printArea ?? null) : null;

  /**
   * The columns carrying a filter button. All of them are mounted and moved
   * imperatively rather than being mounted per visible range: a table is a few
   * dozen columns wide at most, and the overlay clips whatever has scrolled
   * out, so tracking the range in React would buy nothing.
   */
  const filterCols = useMemo(() => {
    if (!autoFilter) return [] as number[];
    // A column the file marks `showButton="0"` gets no drop-down: Excel writes
    // that for every column a merged header covers, so a merged title shows a
    // single button at its right edge instead of one on each column.
    const hidden = new Set(autoFilter.hiddenButtons ?? []);
    const out: number[] = [];
    for (let c = autoFilter.col; c < autoFilter.col + autoFilter.cols; c += 1) {
      if (c < sheetColsGuess && !hidden.has(c)) out.push(c);
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
    if (!autoFilter) return visibleRows;
    // A filter hides rows inside its range and nothing else: the sheet's own
    // rows above the header stay, and so do the ones below the range, which
    // are not the table's to hide — a document with a note under its table
    // keeps the note.
    const headerRow = autoFilter.row;
    const above = Array.from({ length: headerRow + 1 }, (_, index) => index);
    const firstBelow = headerRow + autoFilter.rows;
    const below = Array.from(
      { length: Math.max(0, sheetRowsGuess - firstBelow) },
      (_, index) => firstBelow + index,
    );
    return [...above, ...visibleRows.filter((row) => row > headerRow), ...below];
  }, [visibleRows, autoFilter, sheetRowsGuess]);

  const rowView = useMemo(() => {
    if (!displayRows) return null;
    const index = new Map<number, number>();
    displayRows.forEach((sheetRow, display) => index.set(sheetRow, display));
    return { rows: displayRows, index };
  }, [displayRows]);

  const { t } = useTranslation();

  /**
   * What the reader has hidden or shown this session, laid over the file's own
   * answer. Nothing is written back — reopening the workbook shows exactly what
   * Excel shows — so these live and die with the sheet on screen.
   */
  const [hiddenOverride, setHiddenOverride] = useState<{
    rows: Map<number, boolean>;
    cols: Map<number, boolean>;
  }>({ rows: new Map(), cols: new Map() });

  /**
   * The header the context menu was opened on, when it was one.
   *
   * Kept in a ref beside the state: the state is what the menu draws from, but
   * the action must read the target the pointer chose, not whichever render the
   * item's handler closed over — and a target that outlived its menu sent the
   * next command wherever the last right-click had been.
   */
  const [menuTarget, setMenuTarget] = useState<MenuTarget>(null);
  const menuTargetRef = useRef<MenuTarget>(null);

  const aimMenuAt = (target: MenuTarget) => {
    menuTargetRef.current = target;
    setMenuTarget(target);
  };

  /**
   * The rows and columns the sheet hides. Excel draws none of them and skips the
   * number, so the gap is the only sign it gives; both take no room at all,
   * which is what puts everything after them where Excel has it.
   *
   * Rows inside an autofilter's range are the exception when the file carries
   * filter criteria: Excel stores the rows its filter hides as `hidden="1"`, but
   * that is a cached answer rather than an independent choice — Show All Data
   * unhides every one of them and recomputes from the criteria. Reading them as
   * the sheet's own hidden rows kept rows hidden that a new filter wanted to
   * show, and left them hidden after the filter was cleared.
   */
  const hiddenRows = useMemo(() => {
    const set = new Set<number>();
    const filter = layout?.autoFilter ?? null;
    const filtered =
      filter !== null && Object.keys(filter.applied ?? {}).length > 0;
    const inside = (row: number) =>
      filter !== null && row >= filter.row && row < filter.row + filter.rows;
    for (const row of layout?.hiddenRows ?? []) {
      if (filtered && inside(row)) continue;
      set.add(row);
    }
    for (const [row, hidden] of hiddenOverride.rows) {
      if (hidden) set.add(row);
      else set.delete(row);
    }
    return set;
  }, [layout, hiddenOverride.rows]);

  const hiddenCols = useMemo(() => {
    const set = new Set(layout?.hiddenCols ?? []);
    for (const [col, hidden] of hiddenOverride.cols) {
      if (hidden) set.add(col);
      else set.delete(col);
    }
    return set;
  }, [layout, hiddenOverride.cols]);

  /**
   * The sheet's outline groups, rebuilt from the levels the file records, and
   * the header size they ask for: Excel widens the row header by one 24px
   * strip per level (plus one for "outside any group"), and the column header
   * the same way, so the grouping chrome has its own room beside the numbers.
   */
  const outlineRows = useMemo(
    () => (layout?.outline ? buildRowOutline(layout.outline) : null),
    [layout],
  );
  const outlineCols = useMemo(
    () => (layout?.outline ? buildColOutline(layout.outline) : null),
    [layout],
  );
  const rowGutter = outlineGutter(outlineRows?.strips ?? 0);
  const colGutter = outlineGutter(outlineCols?.strips ?? 0);
  const headerW = ROW_HEADER_W + rowGutter;
  const headerH = COL_HEADER_H + colGutter;

  /**
   * Columns the file hid that the reader has since asked to see. The file keeps
   * no width for them — Excel writes `width="0"` beside the flag — so these
   * come back at the sheet's own default, which is what Excel puts back.
   */
  const revealedCols = useMemo(() => {
    const set = new Set<number>();
    for (const [col, hidden] of hiddenOverride.cols) if (!hidden) set.add(col);
    return set;
  }, [hiddenOverride.cols]);

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
      buildMetrics(colCount, (index) => {
        // `<col>` describes columns a sheet may hold nothing in, and those runs
        // reach past the last cell; outside the extent the sheet's own default
        // width is the honest answer, not the file's.
        if (index >= sheetColsGuess) return baseColW;
        if (hiddenCols.has(index)) return 0;
        const own = colOverrides.get(index);
        if (own !== undefined && own > 0) return own;
        // The file's width for a column it hid is the zero beside the flag, so
        // one the reader has brought back takes the sheet's own default.
        return revealedCols.has(index) ? baseColW : (own ?? baseColW);
      }),
    [colCount, sheetColsGuess, colOverrides, baseColW, hiddenCols, revealedCols],
  );

  const rows = useMemo(
    () =>
      buildMetrics(rowCount, (display) => {
        const row = sheetRowOf(display);
        if (row < 0 || row >= sheetRowsGuess) return baseRowH;
        if (hiddenRows.has(row)) return 0;
        return rowOverrides.get(row) ?? baseRowH;
      }),
    [rowView, rowCount, sheetRowsGuess, rowOverrides, baseRowH, sheetRowOf, hiddenRows],
  );

  const minRowH = Math.max(MIN_ROW_H, Math.ceil(fontSize * LINE_HEIGHT));

  /**
   * Where the pages fall, for Page Break Preview.
   *
   * The file stores only the breaks placed by hand; the rest follow from the
   * paper, its margins and the print scale, and are worked out here from the
   * same column widths and row heights the grid paints with — anything else
   * would put a break somewhere the cells are not.
   */
  const pagination = useMemo(() => {
    if (!printArea || !layout?.pageSetup) return null;
    return paginate(layout.pageSetup, printArea, cols, rows, displayOfRow);
  }, [printArea, layout, cols, rows, displayOfRow]);

  /** What Page Break Preview writes in the middle of each page. */
  const pageLabel = useCallback(
    (number: number) => t("print.page", { n: number }),
    [t],
  );

  /**
   * The app's own font, for the page number. Excel draws it in the UI font
   * rather than the workbook's — the watermark is part of the view, not of the
   * document, and a sheet set in a Mincho would otherwise print its page
   * number in one too.
   */
  const uiFont = useMemo(
    () =>
      (typeof window === "undefined"
        ? ""
        : window.getComputedStyle(document.body).fontFamily) || "sans-serif",
    [],
  );

  /**
   * The frozen bands, in display rows and sheet columns.
   *
   * A frozen row that a filter has taken out of the view holds no height, so
   * the boundary is the first *shown* row at or past the file's count rather
   * than the count itself. The heights and widths come from the same metrics
   * the paint pass lays cells out with, which is what keeps a pinned band and
   * the rows under it from drifting apart.
   */
  const freeze = useMemo(() => {
    const frozenSheetRows = Math.max(0, Math.min(layout?.frozenRows ?? 0, rowCount));
    const frozenSheetCols = Math.max(0, Math.min(layout?.frozenCols ?? 0, colCount));
    let frozenDisplayRows = 0;
    if (frozenSheetRows > 0) {
      if (!rowView) {
        frozenDisplayRows = Math.min(frozenSheetRows, displayRowCount);
      } else {
        let lo = 0;
        let hi = displayRowCount;
        while (lo < hi) {
          const mid = (lo + hi) >> 1;
          if (sheetRowOf(mid) < frozenSheetRows) lo = mid + 1;
          else hi = mid;
        }
        frozenDisplayRows = lo;
      }
    }
    return {
      rows: frozenDisplayRows,
      cols: frozenSheetCols,
      height: frozenDisplayRows > 0 ? (rows.offsets[frozenDisplayRows] ?? 0) : 0,
      width: frozenSheetCols > 0 ? (cols.offsets[frozenSheetCols] ?? 0) : 0,
    };
  }, [layout, rowView, rowCount, colCount, displayRowCount, sheetRowOf, rows, cols]);

  /**
   * Excel's input prompt: a small yellow note beside the cell the cursor is on,
   * when the file's validation rule asks for one. A range selection shows none
   * — only a single active cell has somewhere to put it.
   */
  const promptInfo = useMemo<PromptInfo | null>(() => {
    if (!layout) return null;
    const at = activeArea(selection);
    if (at.kind !== "cell") return null;
    const rule = (layout.validations ?? []).find(
      (rule) =>
        rule.showInput &&
        (rule.prompt || rule.promptTitle) &&
        rule.ranges.some(
          (range) =>
            at.row >= range.row &&
            at.row < range.row + range.rows &&
            at.col >= range.col &&
            at.col < range.col + range.cols,
        ),
    );
    if (!rule) return null;
    return {
      row: at.row,
      col: at.col,
      title: rule.promptTitle ?? "",
      text: rule.prompt ?? "",
    };
  }, [layout, selection]);
  promptInfoRef.current = promptInfo;


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
    const bottom = Math.max(0, scroll.top - headerH) + height;
    const right = Math.max(0, scroll.left - headerW) + width;

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
  }, [usedRowsTotal, usedColsTotal, baseRowH, baseColW, headerW, headerH]);
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
      const block = cache.current.get(`${blockPrefix(path, sheet.name)}${br}${SEP}${bc}`);
      if (!block) return null;
      return block[row - br * BLOCK_ROWS]?.[col - bc * BLOCK_COLS] ?? null;
    },
    [path, sheet.name],
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
   * Opens the text of a cell over it, which is what a double-click and F2 both
   * come to. A merged cell opens as its top-left corner, since that is the only
   * part of the merge the text is in — the rest of it reads as empty — and an
   * empty cell has nothing to offer, so it stays a plain press.
   */
  const openEditor = useCallback(
    (row: number, col: number) => {
      const merge = mergeAt(row, col);
      const at = { row: merge?.row ?? row, col: merge?.col ?? col };
      if ((cellAt(at.row, at.col)?.text ?? "") === "") return;
      setEditing(at);
    },
    [mergeAt, cellAt],
  );

  /**
   * F2 opens the active cell's text, which is the keyboard's way to what a
   * double-click on it does — and the one that works without the pointer.
   *
   * On `window` for the reason Ctrl+A is: a press that landed on a toolbar
   * button leaves the focus there, and the key still has to reach the sheet.
   * What is focused elsewhere keeps it: a field is typing into its own text, and
   * the sheet strip is not the sheet — Excel's F2 means nothing over a tab, so
   * neither does this one.
   */
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "F2" || event.ctrlKey || event.metaKey || event.altKey) return;
      const target = event.target as HTMLElement | null;
      if (target?.closest("input, textarea, [contenteditable], [data-sheet-strip]")) return;
      const at = selectionAnchor(selection);
      openEditor(at.row, at.col);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [selection, openEditor]);

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
        onSelect(rowsBetween(selection.anchor + step, selection.edge + step));
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
        onSelect(colsBetween(selection.anchor + step, selection.edge + step));
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

  /**
   * The status bar is told how many rows the filter is hiding, which is where
   * Excel says it. It used to be a band wedged above the grid, which cost a row
   * of the sheet to say something the sheet is already showing.
   */
  useEffect(() => {
    onFiltered?.(
      visibleRows && autoFilter
        ? { shown: visibleRows.length, total: autoFilter.rows - 1 }
        : null,
    );
  }, [visibleRows, autoFilter, onFiltered]);

  // --- the sheet's layout, once per switch ---
  useEffect(() => {
    drag.current = null;
    /**
     * A workbook that has been read again is another workbook wearing the same
     * path, so the blocks fetched from the old one are dropped before the sheet
     * is asked for. The next paint finds them missing and fetches them — the
     * fetcher runs off the view, and the view is what it has either way.
     */
    if (cachedRevision.current !== revision) {
      cachedRevision.current = revision;
      for (const key of [...cache.current.keys()]) {
        if (key.startsWith(`${path}${SEP}`)) cache.current.delete(key);
      }
    }
    // Which workbook and sheet the sizes now in state belong to, so nothing
    // writes them under another sheet's name before this one's layout lands.
    seeded.current = null;
    // The incoming sheet's own cursor is where its view was left, so arriving
    // with it is not a cursor move for the pass below to carry.
    followedRef.current = selectionRef.current;
    let cancelled = false;
    sheetLayout(path, sheet.name)
      .then((parsed) => {
        if (cancelled) return;
        setLayout(parsed);
        setLayoutSheet(sheetIdentity(path, sheet.name));
        const colPx = (chars: number) =>
          charsToPx(chars, digitWidthPx(parsed.fontFamily, parsed.fontSizePx));
        setBaseColW(colPx(parsed.defaultColWidthChars));
        setBaseRowH(parsed.defaultRowHeightPx);
        setFontFamily(parsed.fontFamily);
        setFontSize(parsed.fontSizePx);
        setShowGrid(parsed.showGridLines);
        seeded.current = sheetIdentity(path, sheet.name);
        // What the reader has dragged or fitted since wins over the file's own
        // sizes, which are only where this sheet started. Same for what they
        // hid: the sheet comes back as it was left, not as the file has it.
        const saved = views.get(path)?.get(sheet.name);
        setHiddenOverride(saved?.hidden ?? { rows: new Map(), cols: new Map() });
        setColOverrides(saved?.sizes?.col ?? toOverrides(parsed.colWidthsChars, colPx));
        setRowOverrides(saved?.sizes?.row ?? toOverrides(parsed.rowHeightsPx));

        // Put back the view this sheet was left at, before anything is
        // fetched, so the first paint is the right block rather than A1's.
        // Excel keeps a window position per sheet and restores it on the way
        // in, even where that leaves the active cell off screen; seeding this
        // from the cursor instead dragged a sheet whose cursor sat far out to
        // the right until that cell was the first visible column.
        setScrollTo(saved?.top ?? 0, saved?.left ?? 0);
        // The file's own view of this sheet: Excel opens it at the zoom and
        // scroll position it was left at, so a reader with no session view of
        // the sheet lands where Excel would.
        if (parsed.zoomScale) {
          onZoom(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, parsed.zoomScale)));
        }
        pendingView.current = saved ? null : (parsed.viewTopLeft ?? null);

        // What this sheet was left holding beats what the workbook stores: a
        // filter applied here is in no file, and an empty Map is a filter the
        // user cleared rather than one that was never set.
        const filters =
          saved?.filters ?? toFilters(parsed.autoFilter?.applied);
        if (filters.size > 0) {
          setFilters(filters);
          computeVisibleRows(path, sheet.name, parsed.autoFilter, filters)
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
  }, [path, sheet.name, revision, views]);

  /**
   * Keeps what the reader has dragged or fitted on the sheet it was done to.
   *
   * Held back until the layout for *this* workbook and sheet has arrived: a
   * switch — another sheet, or another workbook whose sheet happens to share
   * this one's name — leaves the previous sizes in state for a render, and
   * writing those under the new sheet's name would hand it the wrong ones.
   */
  useEffect(() => {
    if (seeded.current !== sheetIdentity(path, sheet.name)) return;
    const bySheet = views.get(path) ?? new Map<string, SheetView>();
    const view = bySheet.get(sheet.name) ?? { ...scrollRef.current };
    bySheet.set(sheet.name, {
      ...view,
      sizes: { col: colOverrides, row: rowOverrides },
      hidden: hiddenOverride,
    });
    views.set(path, bySheet);
  }, [views, path, sheet.name, colOverrides, rowOverrides, hiddenOverride]);

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

  /**
   * A shape takes the pointer so it can be dragged, which would otherwise
   * swallow the wheel over it — the page layer sits beside the scroller, not
   * inside it. The listener is native and non-passive for the same reason the
   * zoom one is: a ctrl+wheel has to be cancelled here, or the webview zooms
   * the whole page on top of the sheet, and a plain wheel must not scroll
   * anything but the grid.
   */
  useEffect(() => {
    const layer = pageLayerRef.current;
    if (!layer) return;
    const onWheel = (event: WheelEvent) => {
      const el = viewportRef.current;
      if (!el) return;
      event.preventDefault();
      event.stopPropagation();
      if (event.ctrlKey || event.metaKey) {
        el.dispatchEvent(
          new WheelEvent("wheel", {
            deltaX: event.deltaX,
            deltaY: event.deltaY,
            deltaMode: event.deltaMode,
            ctrlKey: true,
            shiftKey: event.shiftKey,
            altKey: event.altKey,
            metaKey: event.metaKey,
            clientX: event.clientX,
            clientY: event.clientY,
            bubbles: true,
            cancelable: true,
          }),
        );
        return;
      }
      el.scrollBy({ left: event.deltaX, top: event.deltaY });
    };
    layer.addEventListener("wheel", onWheel, { passive: false });
    return () => layer.removeEventListener("wheel", onWheel);
  }, [layoutReady]);

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

  /**
   * The file's own view, applied once the row and column sizes it is measured
   * against exist. It runs after the zoom effect above, which would otherwise
   * put the scroll back where the old zoom left it.
   */
  useEffect(() => {
    const target = pendingView.current;
    if (!target) return;
    pendingView.current = null;
    const display = displayOfRow(target[0]);
    // A frozen band is drawn on top of the scrolling area, so the row the file
    // names is already visible at scroll zero once it sits just below the band.
    // Excel's `ScrollRow` is that cell, not an absolute offset — subtracting the
    // band is what puts the pane's first cell where the file left it.
    const top = Math.max(0, (display >= 0 ? rows.offsets[display] : 0) - freeze.height);
    const left = Math.max(0, (cols.offsets[target[1]] ?? 0) - freeze.width);
    setScrollTo(top, left);
  }, [cols, rows, displayOfRow, freeze, setScrollTo]);

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
    (next: { path: string; sheet: string; rects: MergeRect[] } | null) => {
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

  /**
   * The field over a cell is moved by the paint pass rather than by React — a
   * scroll re-renders nothing — so opening or closing one has to ask for a frame
   * the way a scroll does.
   */
  useEffect(() => {
    schedule("both");
  }, [editing, schedule]);

  /** It belongs to a cell on a sheet, and goes when that sheet does. */
  useEffect(() => {
    setEditing(null);
  }, [path, sheet.name]);

  /** Focused as it opens, with the caret at the end of the text. */
  useEffect(() => {
    const el = editorRef.current;
    if (!editing || !el) return;
    el.focus({ preventScroll: true });
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }, [editing]);

  /**
   * Nothing may change the text. A native listener, because this is the one
   * event that can still be cancelled while the change is only proposed —
   * every edit the browser knows how to make arrives as one of these, whatever
   * it was typed, pasted, dropped or dragged in.
   */
  useEffect(() => {
    const el = editorRef.current;
    if (!editing || !el) return;
    const refuse = (event: Event) => event.preventDefault();
    el.addEventListener("beforeinput", refuse);
    return () => el.removeEventListener("beforeinput", refuse);
  }, [editing]);

  /**
   * And a second lock behind that one, because the first is not the only way in:
   * `execCommand` — which is what a context menu's Paste uses — makes its edit
   * without proposing one first, so there is nothing there to refuse. Whatever
   * arrives by any route is put straight back.
   */
  useEffect(() => {
    const el = editorRef.current;
    if (!editing || !el) return;
    const asRead = el.innerHTML;
    const observer = new MutationObserver(() => {
      if (el.innerHTML !== asRead) el.innerHTML = asRead;
    });
    observer.observe(el, { childList: true, characterData: true, subtree: true });
    return () => observer.disconnect();
  }, [editing]);

  /** A press anywhere else puts it away, the way leaving a cell does in Excel. */
  useEffect(() => {
    if (!editing) return;
    const onDown = (event: PointerEvent) => {
      if (!editorRef.current?.contains(event.target as Node)) setEditing(null);
    };
    document.addEventListener("pointerdown", onDown, true);
    return () => document.removeEventListener("pointerdown", onDown, true);
  }, [editing]);

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
    pageBreak,
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
    marks,
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

    // A block follows its edge rather than its rectangle, which for one growing
    // leftward is the wrong corner: the edge is the end the pointer and the
    // keyboard walk, so an extension carries it into view as it goes. A
    // Ctrl+clicked selection follows the area just added, that being the one
    // being looked at.
    const at = activeArea(selection);
    const moving =
      at.kind === "range"
        ? at.edge
        : at.kind === "cell"
          ? { row: at.row, col: at.col }
          : null;
    const atRow =
      at.kind === "column" ? -1 : moving ? moving.row : at.kind === "row" ? at.edge : -1;
    const atCol =
      at.kind === "row" ? -1 : moving ? moving.col : at.kind === "column" ? at.edge : -1;

    // Scrolling *up* to a cell means matching its offset, not the offset plus
    // the gutter: the gutter scrolls with the content and would leave the cell
    // tucked behind the header.
    if (atRow >= 0) {
      const display = displayOfRow(atRow);
      if (display >= 0) {
        // A pinned row is always in view; one in the scrolling area has to
        // clear the frozen band above it, not just the scroll offset.
        if (display >= freeze.rows) {
          const frozen = freeze.height * factor;
          const top = (rows.offsets[display] ?? 0) * factor;
          const bottom =
            (headerH + (rows.offsets[Math.min(display + 1, displayRowCount)] ?? 0)) * factor;
          if (top < el.scrollTop + frozen) el.scrollTop = top - frozen;
          else if (bottom > el.scrollTop + el.clientHeight) el.scrollTop = bottom - el.clientHeight;
        }
      }
    }
    if (atCol >= 0 && atCol >= freeze.cols) {
      const frozen = freeze.width * factor;
      const left = (cols.offsets[atCol] ?? 0) * factor;
      const right =
        (headerW + (cols.offsets[Math.min(atCol + 1, sheetColsGuess)] ?? 0)) * factor;
      if (left < el.scrollLeft + frozen) el.scrollLeft = left - frozen;
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
    freeze,
    headerW,
    headerH,
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
      const prefix = blockPrefix(path, sheet.name);
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
          needed.add(`${prefix}${block}${SEP}${bc}`);
        }
      }

      for (const key of needed) {
        if (cache.current.has(key) || pending.current.has(key)) continue;
        pending.current.add(key);
        const [rawRow, rawCol] = key.slice(prefix.length).split(SEP);
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
    if (!layout || layoutSheet !== sheetIdentity(path, sheet.name)) return null;
    const scroll = scrollRef.current;
    const size = sizeRef.current;
    const z = zoomRef.current;
    if (size.width <= 0 || size.height <= 0) return null;

    const sheetRows = Math.max(sheet.rows, layout.rows);
    const sheetCols = Math.max(sheet.cols, layout.cols);

    const top = Math.max(0, scroll.top - headerH);
    const left = Math.max(0, scroll.left - headerW);
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
    const seenMerges = new Set<number>();
    /**
     * Adds the merges a rectangular window touches. A frozen band is a window
     * of its own — the pinned rows and the pinned columns sit outside the
     * scrolling range, and a region that straddles the freeze line has to be
     * drawable in both panes.
     */
    const addMerges = (rowLo: number, rowHi: number, colLo: number, colHi: number) => {
      for (const merge of merges) {
        const end = merge.row + merge.rows - 1;
        if (end < rowLo || merge.row > rowHi) continue;
        const endCol = merge.col + merge.cols - 1;
        if (endCol < colLo || merge.col > colHi) continue;
        if (displayOfRow(merge.row) < 0) continue;
        const anchor = cellKey(sheetCols, merge.row, merge.col);
        if (!seenMerges.has(anchor)) {
          seenMerges.add(anchor);
          shownMerges.push(merge);
        }
        for (let r = Math.max(merge.row, rowLo); r <= Math.min(end, rowHi); r += 1) {
          for (let c = Math.max(merge.col, colLo); c <= Math.min(endCol, colHi); c += 1) {
            const key = cellKey(sheetCols, r, c);
            coveredMerges.add(key);
            mergeIndex.set(key, merge);
          }
        }
      }
    };
    const frozenSheetRows = Math.min(layout.frozenRows ?? 0, sheetRows);
    const frozenSheetCols = Math.min(layout.frozenCols ?? 0, sheetCols);
    addMerges(lo, hi, range.colStart, range.colEnd);
    if (frozenSheetRows > 0) addMerges(0, frozenSheetRows - 1, range.colStart, range.colEnd);
    if (frozenSheetCols > 0) addMerges(lo, hi, 0, frozenSheetCols - 1);
    const mergeAtFast = (row: number, col: number): MergeRect | null =>
      mergeIndex.get(cellKey(sheetCols, row, col)) ?? null;

    /**
     * Review marks, keyed off the pass the same way the styles are — a cell's
     * mark is asked about once for its own fill. Most sheets carry none, so the
     * empty map short-circuits before any key is built.
     */
    const markMemo = new Map<number, string | null>();
    const markAtFast = (row: number, col: number): string | null => {
      if (marks.size === 0) return null;
      // The key only identifies a cell *inside* the sheet: out in the blank area
      // a column past the last one shares its number with the next row's first
      // columns, and this pass does paint out there. On a sheet narrower than
      // the window those blank cells are asked about first, so they answered for
      // the row below and the mark went missing on everything but a merge, whose
      // anchor is read before any row is. Nothing out there carries a mark.
      if (row >= sheetRows || col >= sheetCols) return null;
      const key = cellKey(sheetCols, row, col);
      const hit = markMemo.get(key);
      if (hit !== undefined) return hit;
      const value = marks.get(markKey(sheet.name, row, col)) ?? null;
      markMemo.set(key, value);
      return value;
    };

    /**
     * The validation rule over a cell. Rules are few and each names its blocks,
     * so a walk is cheaper than the per-cell index a mark needs — and this is
     * only asked about the active cell.
     */
    const validations = layout.validations ?? [];
    const validationAtFast = (row: number, col: number) => {
      for (const rule of validations) {
        for (const range of rule.ranges) {
          if (
            row >= range.row &&
            row < range.row + range.rows &&
            col >= range.col &&
            col < range.col + range.cols
          ) {
            return rule;
          }
        }
      }
      return null;
    };

    const shownRuns: (MergeRect & { text: string | null })[] = [];
    const coveredRuns = new Set<number>();
    const joinable = (row: number, col: number) =>
      col >= 0 &&
      col < sheetCols &&
      styleAtFast(row, col)?.horizontal === "centerContinuous" &&
      mergeAtFast(row, col) === null;
    const textAt = (row: number, col: number) => cellAt(row, col)?.text ?? "";
    const seenRuns = new Set<number>();
    /** The same windows the merges get: the scrolling one and the pinned bands. */
    const addRuns = (rowFrom: number, rowTo: number, colFrom: number, colTo: number) => {
      for (let row = rowFrom; row <= rowTo; row += 1) {
        if (row < 0 || row >= sheetRows || displayOfRow(row) < 0) continue;
        for (let col = colFrom; col < colTo; col += 1) {
          if (!joinable(row, col)) continue;
          const text = textAt(row, col) || null;
          /**
           * A run is the cell holding the text and the empty cells after it, and
           * it stops at the next cell with a value of its own — which is why two
           * labels side by side stay two runs. Reading it as one unbroken stretch
           * of the same alignment ran `変更項番` across J..L and then on across
           * the `変更者` columns beside it, centring it over all six.
           */
          if (text === null && joinable(row, col - 1)) continue;
          let end = col;
          while (joinable(row, end + 1) && textAt(row, end + 1) === "") end += 1;
          if (end === col) continue;
          const anchor = cellKey(sheetCols, row, col);
          if (!seenRuns.has(anchor)) {
            seenRuns.add(anchor);
            shownRuns.push({ row, col, rows: 1, cols: end - col + 1, text });
          }
          for (let c = col; c <= end; c += 1) coveredRuns.add(cellKey(sheetCols, row, c));
        }
      }
    };
    addRuns(lo, hi, range.colStart, range.colEnd);
    if (frozenSheetRows > 0) addRuns(0, frozenSheetRows - 1, range.colStart, range.colEnd);
    if (frozenSheetCols > 0) addRuns(lo, hi, 0, frozenSheetCols - 1);

    /**
     * Each run's cells back to the run, so the painter can drop the borders
     * between them — Excel draws a Center-Across run as one box.
     */
    const runIndex = new Map<number, MergeRect>();
    for (const run of shownRuns) {
      for (let c = run.col; c < run.col + run.cols; c += 1) {
        runIndex.set(cellKey(sheetCols, run.row, c), run);
      }
    }

    /**
     * The runs a cell draws part of its text in, keyed by the cell. Sparse: only
     * cells that format part of their text are in it, and two cells spelling
     * the same text may carry different runs.
     */
    const richIndex = new Map<number, RichRun[]>();
    for (const cell of layout.richCells ?? []) {
      richIndex.set(cellKey(sheetCols, cell.row, cell.col), cell.runs);
    }

    /** The data bars by cell, for the bars drawn over the cell's fill. */
    const barIndex = new Map<number, DataBarCell>();
    for (const bar of layout.dataBars ?? []) {
      barIndex.set(cellKey(sheetCols, bar.row, bar.col), bar);
    }

    /** The conditional icons by cell. */
    const iconIndex = new Map<number, IconCell>();
    for (const icon of layout.icons ?? []) {
      iconIndex.set(cellKey(sheetCols, icon.row, icon.col), icon);
    }

    /** Cells carrying a note, which Excel marks with a red corner triangle. */
    const noteIndex = new Set<number>();
    for (const note of layout.notes ?? []) {
      noteIndex.add(cellKey(sheetCols, note.row, note.col));
    }

    return {
      cols,
      rows,
      sheetRows,
      sheetCols,
      hiddenRows,
      hiddenCols,
      headerW,
      headerH,
      outlineRows,
      outlineCols,
      ...range,
      width,
      height,
      zoom: z,
      scrollTop: scroll.top,
      scrollLeft: scroll.left,
      sheetRowOf,
      displayOfRow,
      cellAt,
      styleAt: styleAtFast,
      styleCalls,
      mergeAt: mergeAtFast,
      markAt: markAtFast,
      validationAt: validationAtFast,
      workbookFontFamily: fontFamily,
      workbookFontSizePx: fontSize,
      showGrid,
      selection: selectionRef.current,
      copied:
        copiedRef.current?.path === path && copiedRef.current?.sheet === sheet.name
          ? copiedRef.current.rects
          : [],
      ants: antsRef.current,
      merges,
      mergeView: { shown: shownMerges, covered: coveredMerges, at: mergeIndex },
      centerRuns: { shown: shownRuns, covered: coveredRuns, at: runIndex },
      autoFilter,
      filtering:
        (visibleRows !== null && filters.size > 0) || customFilterCols.size > 0,
      printArea,
      pageBreaks: pagination,
      pageLabel,
      uiFont,
      outsidePrint,
      cellRich: richIndex,
      dataBars: barIndex,
      icons: iconIndex,
      notes: noteIndex,
      theme: canvasTheme(),
    };
  }, [
    layout,
    layoutSheet,
    path,
    sheet.name,
    sheet.rows,
    sheet.cols,
    rowView,
    cols,
    rows,
    hiddenRows,
    hiddenCols,
    headerW,
    headerH,
    outlineRows,
    outlineCols,
    rowCount,
    colCount,
    merges,
    sheetRowOf,
    displayOfRow,
    cellAt,
    styleAt,
    mergeAt,
    marks,
    fontFamily,
    fontSize,
    showGrid,
    autoFilter,
    visibleRows,
    filters,
    customFilterCols,
    printArea,
    pagination,
    pageLabel,
    uiFont,
    outsidePrint,
  ]);

  const paint = useCallback(
    (layers: "both" | "overlay") => {
      const view = buildView();
      // Nothing drawn until the incoming sheet's layout is in: leaving the last
      // sheet's pixels untouched is what keeps a switch from flashing blank.
      if (!view) return;
      /**
       * The panes a frozen view is divided into, each with the region of the
       * canvas it owns. A pane is the same sheet drawn at its own scroll
       * offset: the pinned rows and columns use no offset at all, so they stay
       * put while the rest scrolls, and the clip is what stops the scrolling
       * content from painting over them.
       */
      const gx = headerW;
      const gy = headerH;
      const fr = freeze.rows;
      const fc = freeze.cols;
      const fw = freeze.width;
      const fh = freeze.height;
      const bodyW = Math.max(0, view.width - gx - fw);
      const bodyH = Math.max(0, view.height - gy - fh);
      const frozenRowEnd = Math.min(rowCount, fr + 1 + OVERSCAN_ROWS);
      const frozenColEnd = Math.min(colCount, fc + 1 + OVERSCAN_COLS);
      const panes: { view: GridView; clips: { x: number; y: number; w: number; h: number }[] }[] = [];
      if (fr === 0 && fc === 0) {
        panes.push({ view, clips: [] });
      } else {
        if (bodyW > 0 || bodyH > 0) {
          panes.push({
            view,
            clips: [
              { x: gx + fw, y: gy + fh, w: bodyW, h: bodyH },
              { x: 0, y: gy + fh, w: gx, h: bodyH },
              { x: gx + fw, y: 0, w: bodyW, h: gy },
              { x: 0, y: 0, w: gx, h: gy },
            ],
          });
        }
        if (fr > 0 && fh > 0) {
          panes.push({
            view: { ...view, scrollTop: 0, rowStart: 0, rowEnd: frozenRowEnd },
            clips: [
              { x: gx + fw, y: gy, w: bodyW, h: fh },
              { x: gx + fw, y: 0, w: bodyW, h: gy },
              ...(fc === 0 ? [{ x: 0, y: gy, w: gx, h: fh }] : []),
            ],
          });
        }
        if (fc > 0 && fw > 0) {
          panes.push({
            view: { ...view, scrollLeft: 0, colStart: 0, colEnd: frozenColEnd },
            clips: [
              { x: gx, y: gy + fh, w: fw, h: bodyH },
              { x: 0, y: gy + fh, w: gx, h: bodyH },
              ...(fr === 0 ? [{ x: gx, y: 0, w: fw, h: gy }] : []),
            ],
          });
        }
        if (fr > 0 && fc > 0 && fh > 0 && fw > 0) {
          panes.push({
            view: {
              ...view,
              scrollTop: 0,
              scrollLeft: 0,
              rowStart: 0,
              rowEnd: frozenRowEnd,
              colStart: 0,
              colEnd: frozenColEnd,
            },
            clips: [
              { x: gx, y: gy, w: fw, h: fh },
              { x: 0, y: gy, w: gx, h: fh },
              { x: gx, y: 0, w: fw, h: gy },
            ],
          });
        }
      }
      const clipTo = (ctx: CanvasRenderingContext2D, clips: { x: number; y: number; w: number; h: number }[]) => {
        if (clips.length === 0) return;
        ctx.beginPath();
        for (const clip of clips) {
          if (clip.w > 0 && clip.h > 0) ctx.rect(clip.x, clip.y, clip.w, clip.h);
        }
        ctx.clip();
      };

      // Whatever this frame is about to show has to have its data asked for.
      // A pane can reach rows the scrolling window does not, which is the whole
      // point of a frozen one, so each asks for its own.
      for (const pane of panes) ensureBlocks(pane.view);
      const startedAt = performance.now();
      const scale = view.zoom * dprRef.current;
      /**
       * Clear the whole backing store first, in device pixels. The panes clear
       * their own rectangles under the zoom transform, which at anything below
       * 100% leaves the right and bottom of the canvas outside that transform —
       * stale pixels from the previous zoom, reading as "the corner zoomed, the
       * rest didn't".
       */
      const clearAll = (ctx: CanvasRenderingContext2D) => {
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
        ctx.setTransform(scale, 0, 0, scale, 0, 0);
      };
      if (layers === "both") {
        const ctx = baseRef.current?.getContext("2d");
        if (ctx) {
          clearAll(ctx);
          const stats: PaintStats = {
            rows: 0,
            cols: 0,
            spans: 0,
            texts: 0,
            surfaceMs: 0,
            surfaceFills: 0,
            textMs: 0,
            styleResolved: 0,
            styleMemoHits: 0,
          };
          for (const pane of panes) {
            ctx.save();
            clipTo(ctx, pane.clips);
            const part = paintBase(ctx, pane.view);
            ctx.restore();
            stats.rows += part.rows;
            stats.cols += part.cols;
            stats.spans += part.spans;
            stats.texts += part.texts;
            stats.surfaceMs += part.surfaceMs;
            stats.surfaceFills += part.surfaceFills;
            stats.textMs += part.textMs;
            stats.styleResolved += part.styleResolved;
            stats.styleMemoHits += part.styleMemoHits;
          }
          // Excel marks a frozen split with a rule darker than a gridline, drawn
          // across the gutters as well and kept when gridlines are turned off.
          // Every pane is clipped to its own region, so this goes in after them.
          ctx.fillStyle = view.theme.freezeLine;
          if (fr > 0 && fh > 0) ctx.fillRect(0, gy + fh, view.width, 1);
          if (fc > 0 && fw > 0) ctx.fillRect(gx + fw, 0, 1, view.height);
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
        clearAll(ctx);
        for (const pane of panes) {
          ctx.save();
          clipTo(ctx, pane.clips);
          paintOverlay(ctx, pane.view);
          ctx.restore();
        }
      }
      perf.current.overlayMs = performance.now() - overlayStarted;
      perf.current.layers = layers === "both" ? 2 : 1;

      // The filter buttons are DOM, so they have to be moved by hand: the
      // canvas scrolls under them and a scroll does not re-render React.
      const area = overlayAreaRef.current;
      if (area) {
        const gutterX = headerW * view.zoom;
        const gutterY = headerH * view.zoom;
        area.style.left = `${gutterX}px`;
        area.style.top = `${gutterY}px`;
        // The view is measured in sheet pixels; the layer that carries the
        // buttons is laid out in screen ones.
        area.style.width = `${Math.max(0, view.width * view.zoom - gutterX)}px`;
        area.style.height = `${Math.max(0, view.height * view.zoom - gutterY)}px`;

        const headerDisplay = view.autoFilter ? view.displayOfRow(view.autoFilter.row) : -1;
        const filterHeaderH = headerDisplay >= 0 ? spanAt(view.rows.offsets, headerDisplay) : 0;
        for (const [col, slot] of filterSlots.current) {
          if (headerDisplay < 0) {
            slot.style.display = "none";
            continue;
          }
          const x = (view.cols.offsets[col] - (col < fc ? 0 : view.scrollLeft)) * view.zoom;
          const y =
            (view.rows.offsets[headerDisplay] - (headerDisplay < fr ? 0 : view.scrollTop)) *
            view.zoom;
          // A column scrolled under the frozen band belongs to the scrolling
          // pane, which is clipped there; this layer is above the canvas, so
          // its button has to be hidden by hand or it shows over the fixed
          // columns. Same for a header row scrolled under frozen rows.
          const underFreeze =
            (col >= fc && x < freeze.width * view.zoom) ||
            (headerDisplay >= fr && y < freeze.height * view.zoom);
          slot.style.display = underFreeze ? "none" : "";
          if (underFreeze) continue;
          slot.style.transform = `translate(${x}px, ${y}px)`;
          slot.style.width = `${spanAt(view.cols.offsets, col) * view.zoom}px`;
          slot.style.height = `${filterHeaderH * view.zoom}px`;
        }

        // The file's input prompt sits beside the active cell — to its right
        // when that fits, to its left otherwise, and never past the view's own
        // edges, which is where Excel keeps it too.
        const prompt = promptRef.current;
        if (prompt) {
          const info = promptInfoRef.current;
          const display = info ? view.displayOfRow(info.row) : -1;
          if (!info || display < 0) {
            prompt.style.display = "none";
          } else {
            const x = (view.cols.offsets[info.col] - (info.col < fc ? 0 : view.scrollLeft)) * view.zoom;
            const y = (view.rows.offsets[display] - (display < fr ? 0 : view.scrollTop)) * view.zoom;
            const w = spanAt(view.cols.offsets, info.col) * view.zoom;
            prompt.style.display = "";
            prompt.style.transform = "translate(0px, 0px)";
            const pw = prompt.offsetWidth;
            const ph = prompt.offsetHeight;
            const left =
              x + w + 2 + pw <= area.clientWidth ? x + w + 2 : Math.max(0, x - pw - 2);
            const top = Math.min(Math.max(0, y), Math.max(0, area.clientHeight - ph - 1));
            prompt.style.transform = `translate(${left}px, ${top}px)`;
          }
        }

        // The field over an open cell is DOM as well, and goes where its cell
        // is — over the whole merge when the cell is part of one, in that cell's
        // own font and fill, so what is read there reads as what was on the
        // sheet rather than as a box that appeared over it.
        const field = editorRef.current;
        if (field) {
          const at = editingRef.current;
          const merge = at ? view.mergeAt(at.row, at.col) : null;
          const row = at ? (merge?.row ?? at.row) : -1;
          // Where that row is on screen. The offsets are indexed by the rows
          // this view shows, which a filter has thinned out, so the sheet's own
          // row number would put the field over whatever cell now stands there
          // — a row above the edit having been hidden moves it up.
          const display = row < 0 ? -1 : view.displayOfRow(row);
          if (!at || display < 0) {
            // No cell open, or a row the filter has taken out of the view:
            // either way there is nothing for the field to sit on.
            field.style.display = "none";
          } else {
            const col = merge?.col ?? at.col;
            const style = view.styleAt(at.row, at.col);
            let w = 0;
            for (let c = col; c < col + (merge?.cols ?? 1); c += 1) {
              w += spanAt(view.cols.offsets, c);
            }
            // A value is read in the room beside a narrow cell when the cells
            // after it are empty, and the field has to be that same box: the
            // width of the column alone is nothing at all for a column two
            // pixels wide. A merge is painted whole and never spills.
            if (!merge) w = valueWidth(view, row, col, w, CELL_PAD_LEFT, CELL_PAD_RIGHT);
            // The merge's own rows, as far down the view as the last of them
            // still shows.
            const lastRow = row + (merge?.rows ?? 1) - 1;
            let lastDisplay = display;
            for (let r = row; r <= lastRow; r += 1) {
              const shown = view.displayOfRow(r);
              if (shown >= 0) lastDisplay = shown;
            }
            let h = 0;
            for (let r = display; r <= lastDisplay; r += 1) {
              h += spanAt(view.rows.offsets, r);
            }
            const size = (style?.sizePx ?? view.workbookFontSizePx) * view.zoom;
            field.style.display = "";
            field.style.transform =
              `translate(${(view.cols.offsets[col] - (col < fc ? 0 : view.scrollLeft)) * view.zoom}px, ` +
              `${(view.rows.offsets[display] - (display < fr ? 0 : view.scrollTop)) * view.zoom}px)`;
            // The cell's boundary is one pixel past its last one on the right
            // and bottom — that is where the ring and the cell's own border are
            // drawn. The field has to reach that far, and it wears the
            // selection's own rule rather than a border of its own: an inset
            // shadow takes no room in the box, so the rule lands on the pixels
            // the selection covers while the cell's text stays exactly where the
            // canvas drew it. A border cannot — its width comes out of the
            // content box, which pushed the text along and left a top-aligned
            // cell's line nowhere to go.
            field.style.width = `${(w + 1) * view.zoom}px`;
            field.style.height = `${(h + 1) * view.zoom}px`;
            field.style.boxShadow = `inset 0 0 0 ${RING_W * view.zoom}px ${ACCENT_INK}`;
            // The cell's own insets, scaled with the view: 3 left and 2 right,
            // which is Excel's and is not symmetric. The extra pixel the field
            // is wider by belongs to the right one, so the text ends where the
            // canvas ends it rather than one pixel short.
            field.style.paddingLeft = `${CELL_PAD_LEFT * view.zoom}px`;
            field.style.paddingRight = `${(CELL_PAD_RIGHT + 1) * view.zoom}px`;
            field.style.paddingBottom = "0px";
            field.style.font =
              `${style?.italic ? "italic " : ""}${style?.bold ? "700 " : ""}${size}px ` +
              cellFontStack(style?.family ?? null, view.workbookFontFamily);
            field.style.color = style?.color ?? "";
            // What the canvas painted under the text, in the canvas's own
            // order: the reader's mark wins over everything, the print shade
            // sits over the file's fill, and white is only what is left when
            // the file says nothing. The field's own white used to cover
            // whatever was there, so a marked cell lost its colour the moment
            // it was opened.
            field.style.background =
              view.markAt(row, col) ??
              (view.outsidePrint(at.row, at.col) ? OUTSIDE_PRINT : (style?.fill ?? "#fff"));
            field.style.textAlign =
              style?.horizontal === "center" ||
              style?.horizontal === "centerContinuous"
                ? "center"
                : style?.horizontal === "right"
                  ? "right"
                  : "left";
            // Where the cell's own text sat: one line box of it, placed by the
            // cell's vertical alignment. Without this the text of a bottom- or
            // middle-aligned cell — which is most of them — would jump to the
            // top of the field and the cell would look as though it had moved.
            const lineHeight = size * LINE_HEIGHT;
            const room = h * view.zoom - lineHeight;
            const vertical = style?.vertical ?? "bottom";
            field.style.lineHeight = `${lineHeight}px`;
            field.style.paddingTop = `${
              Math.max(0, vertical === "top" ? 0 : vertical === "middle" ? room / 2 : room)
            }px`;
          }
        }
      }

      // The shapes and the print mark are DOM too, and the canvas underneath
      // them does not scroll — so this layer is translated and scaled with the
      // scroll rather than being scrolled, and clipped to the cell area so a
      // shape near row 0 cannot paint over the gutters.
      const page = pageLayerRef.current;
      if (page) {
        const gutterX = headerW * view.zoom;
        const gutterY = headerH * view.zoom;
        page.style.left = `${gutterX}px`;
        page.style.top = `${gutterY}px`;
        page.style.width = `${Math.max(0, view.width * view.zoom - gutterX)}px`;
        page.style.height = `${Math.max(0, view.height * view.zoom - gutterY)}px`;

        const inner = shapesInnerRef.current;
        if (inner) {
          inner.style.transform =
            `translate(${-view.scrollLeft * view.zoom}px, ${-view.scrollTop * view.zoom}px) ` +
            `scale(${view.zoom})`;
          // A shape anchored in a frozen band is pinned to it: the layer was
          // just moved by the scroll, so those shapes add it back. The layer
          // does not re-render on a scroll, which is why this is done here
          // rather than with a prop.
          if (freeze.rows > 0 || freeze.cols > 0) {
            for (const node of inner.querySelectorAll<SVGGElement>("[data-shape-x]")) {
              const x = Number(node.dataset.shapeX ?? 0);
              const y = Number(node.dataset.shapeY ?? 0);
              const dx = node.hasAttribute("data-pin-x") ? view.scrollLeft : 0;
              const dy = node.hasAttribute("data-pin-y") ? view.scrollTop : 0;
              node.setAttribute("transform", `translate(${x + dx} ${y + dy})`);
            }
          }
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
    [buildView, ensureBlocks, printColX, printRowY, freeze, rowCount, colCount],
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
      const gutterW = headerW * z;
      const gutterH = headerH * z;
      const frozenW = freeze.width * z;
      const frozenH = freeze.height * z;
      /** A point inside a frozen band is the viewport's own; the rest scrolls. */
      const sheetX = (viewportX: number) =>
        (viewportX + (viewportX < gutterW + frozenW ? 0 : el.scrollLeft)) / z - headerW;
      const sheetY = (viewportY: number) =>
        (viewportY + (viewportY < gutterH + frozenH ? 0 : el.scrollTop)) / z - headerH;

      // The scroller's scrollbars are inside its border box but are not the
      // sheet. Hit-testing one as a cell started a block selection on the
      // rightmost visible column and then dragged it along as the thumb moved —
      // and because the content scrolls under a stationary pointer, the cell
      // climbed a row at a time even though the pointer never left the bar.
      if (px >= el.clientWidth || py >= el.clientHeight) return null;

      if (px < gutterW && py < gutterH) return { kind: "corner" };
      // The gutters are pinned, so their coordinate is the viewport's own.
      if (py < gutterH) {
        const x = sheetX(clientX - rect.left);
        if (x < 0 || x >= cols.total) return null;
        return { kind: "colHeader", col: Math.min(colCount - 1, indexAt(cols.offsets, x)) };
      }
      if (px < gutterW) {
        const y = sheetY(clientY - rect.top);
        if (y < 0 || y >= rows.total) return null;
        const display = Math.min(rowCount - 1, indexAt(rows.offsets, y));
        const row = sheetRowOf(display);
        return row < 0 ? null : { kind: "rowHeader", row };
      }

      const x = sheetX(clientX - rect.left);
      const y = sheetY(clientY - rect.top);
      if (x < 0 || y < 0 || x >= cols.total || y >= rows.total) return null;
      const row = sheetRowOf(Math.min(rowCount - 1, indexAt(rows.offsets, y)));
      if (row < 0) return null;
      return { kind: "cell", row, col: Math.min(colCount - 1, indexAt(cols.offsets, x)) };
    },
    [cols, rows, colCount, rowCount, sheetRowOf, freeze, headerW, headerH],
  );

  /**
   * What the outline chrome under a point does when pressed: a group's +/-,
   * or one of the corner buttons that opens every group to a level. Null for
   * everything else, so the press falls through to the usual selection.
   */
  const outlineActionAt = useCallback(
    (clientX: number, clientY: number): OutlineAction | null => {
      const el = viewportRef.current;
      if (!el) return null;
      const z = zoomRef.current;
      const rect = el.getBoundingClientRect();
      const px = clientX - rect.left;
      const py = clientY - rect.top;
      if (px < 0 || py < 0 || px >= el.clientWidth || py >= el.clientHeight) return null;
      const top = headerH * z;
      const left = headerW * z;

      // The level buttons share the corner: rows along its bottom, columns up
      // its right side, exactly as they were measured off Excel's window.
      if (px < left && py < top) {
        for (let index = 0; outlineRows && index < outlineRows.strips; index += 1) {
          const box = rowLevelButtonRect(index, colGutter);
          if (
            px >= box.x * z &&
            px < (box.x + box.w) * z &&
            py >= box.y * z &&
            py < (box.y + box.h) * z
          ) {
            return { kind: "rowLevel", level: index };
          }
        }
        for (let index = 0; outlineCols && index < outlineCols.strips; index += 1) {
          const box = colLevelButtonRect(index, rowGutter);
          if (
            px >= box.x * z &&
            px < (box.x + box.w) * z &&
            py >= box.y * z &&
            py < (box.y + box.h) * z
          ) {
            return { kind: "colLevel", level: index };
          }
        }
        return null;
      }

      if (outlineRows && px < (headerW - ROW_HEADER_W) * z && py >= top) {
        for (const group of outlineRows.groups) {
          if (hiddenRows.has(group.summary)) continue;
          const display = displayOfRow(group.summary);
          if (display < 0) continue;
          const box = groupButtonBox(
            group.level - 1,
            rows.offsets[display],
            spanAt(rows.offsets, display),
            false,
          );
          const x = box.x * z;
          const y =
            (headerH + box.y) * z - (display < freeze.rows ? 0 : el.scrollTop);
          if (px >= x && px < x + box.w * z && py >= y && py < y + box.h * z) {
            return { kind: "rowGroup", group };
          }
        }
      }

      if (outlineCols && py < (headerH - COL_HEADER_H) * z && px >= left) {
        for (const group of outlineCols.groups) {
          if (hiddenCols.has(group.summary)) continue;
          const box = groupButtonBox(
            group.level - 1,
            cols.offsets[group.summary],
            spanAt(cols.offsets, group.summary),
            true,
          );
          const x =
            (headerW + box.x) * z - (group.summary < freeze.cols ? 0 : el.scrollLeft);
          const y = box.y * z;
          if (px >= x && px < x + box.w * z && py >= y && py < y + box.h * z) {
            return { kind: "colGroup", group };
          }
        }
      }
      return null;
    },
    [
      outlineRows,
      outlineCols,
      hiddenRows,
      hiddenCols,
      displayOfRow,
      rows,
      cols,
      freeze,
      headerW,
      headerH,
    ],
  );

  /** Applies an outline press, as a session-only change like hiding a row. */
  const applyOutlineAction = (action: OutlineAction) => {
    const isRows = action.kind === "rowGroup" || action.kind === "rowLevel";
    const lineOutline = isRows ? outlineRows : outlineCols;
    if (!lineOutline) return;
    if (action.kind === "rowLevel" || action.kind === "colLevel") {
      // Excel's corner buttons show up to their own level and close what is
      // deeper, whatever each group had been left at.
      setHiddenOverride((prev) => {
        const map = new Map(isRows ? prev.rows : prev.cols);
        for (const [line, level] of lineOutline.levels) map.set(line, level > action.level);
        return isRows ? { ...prev, rows: map } : { ...prev, cols: map };
      });
      return;
    }
    const group = action.group;
    const hidden = isRows ? hiddenRows : hiddenCols;
    const open = isRows
      ? groupOpen(group, hidden, displayOfRow)
      : (() => {
          for (let col = group.from; col <= group.to; col += 1) {
            if (!hiddenCols.has(col)) return true;
          }
          return false;
        })();
    const lines = groupToggleLines(
      group,
      (line) => lineOutline.levels.get(line) ?? 0,
      open,
    );
    setHiddenOverride((prev) => {
      const map = new Map(isRows ? prev.rows : prev.cols);
      // Open now means the click closes it, so the line's new hidden state is
      // exactly whether the group was open.
      for (const line of lines) map.set(line, open);
      return isRows ? { ...prev, rows: map } : { ...prev, cols: map };
    });
  };

  /**
   * The selection a drag from `anchor` has reached, or null when it has not
   * reached anything new — which is what keeps a click that never moved a click
   * rather than a one-cell block.
   *
   * A gutter drag follows the pointer's row or column wherever it goes, the
   * sheet included: Excel keeps selecting rows after the pointer leaves the
   * gutter, and only the end it started from stays put.
   */
  const dragTo = useCallback((anchor: DragAnchor, hit: Exclude<Hit, null>): Area | null => {
    if (hit.kind === "corner") return null;
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
    const left = rect.left + headerW * z;
    const top = rect.top + headerH * z;
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
          const base = dragBase.current;
          onSelect(base ? withArea(base, next) : next);
        }
      }
    }

    autoScroll.current = requestAnimationFrame(runAutoScroll);
  }, [dragTo, hitAt, onSelect, schedule, headerW, headerH]);

  const handleClick = (event: React.MouseEvent<HTMLDivElement>) => {
    if (outlinePress.current) {
      outlinePress.current = false;
      return;
    }
    if (suppressClick.current) {
      suppressClick.current = false;
      return;
    }
    const hit = hitAt(event.clientX, event.clientY);
    if (!hit) return;
    const add = event.ctrlKey || event.metaKey;
    // Excel's select-all square: the whole sheet, as Ctrl+A ends up doing.
    if (hit.kind === "corner") {
      const sheetRows = Math.max(sheet.rows, layout?.rows ?? 0);
      onSelect(rowsBetween(0, sheetRows - 1));
      return;
    }
    if (hit.kind === "cell") {
      // A Ctrl+click adds what it landed on to the selection rather than moving
      // there — and takes off the area it landed inside of, the whole of that
      // area rather than the part of it pointed at, which is how Excel lets a
      // block be picked up and then put back.
      if (add) {
        const merge = mergeAt(hit.row, hit.col);
        const row = merge ? merge.row : hit.row;
        const col = merge ? merge.col : hit.col;
        onSelect(toggleArea(selection, { kind: "cell", row, col }, { kind: "cell", row, col }));
        return;
      }
      // Shift extends the block from the cell it is pinned at: the anchor is
      // the pivot, so the same corner keeps growing and a click back inside the
      // block redraws it smaller rather than taking a corner off.
      if (event.shiftKey) {
        // A run of whole rows or columns stays a run when a cell is clicked: the
        // whole line that cell is on joins it, the way the next header would.
        if (selection.kind === "row") {
          onSelect(rowsBetween(selection.anchor, hit.row));
        } else if (selection.kind === "column") {
          onSelect(colsBetween(selection.anchor, hit.col));
        } else {
          onSelect(rangeBetween(selectionAnchor(selection), { row: hit.row, col: hit.col }));
        }
        return;
      }
      select(hit.row, hit.col);
    } else if (add) {
      // A gutter takes a whole line into the selection, on its own or off it.
      if (hit.kind === "rowHeader") {
        onSelect(
          toggleArea(selection, rowsBetween(hit.row, hit.row), {
            kind: "row",
            row: hit.row,
            col: 0,
          }),
        );
      } else {
        onSelect(
          toggleArea(selection, colsBetween(hit.col, hit.col), {
            kind: "column",
            row: 0,
            col: hit.col,
          }),
        );
      }
    } else {
      // Shift grows the run from the row or column the selection is anchored
      // on, the way it grows a block from its anchor cell — so a run is made by
      // picking one end and shift-clicking the other.
      const from = selectionAnchor(selection);
      if (hit.kind === "rowHeader") {
        const anchor = selection.kind === "row" ? selection.anchor : from.row;
        onSelect(event.shiftKey ? rowsBetween(anchor, hit.row) : rowsBetween(hit.row, hit.row));
      } else {
        const anchor = selection.kind === "column" ? selection.anchor : from.col;
        onSelect(event.shiftKey ? colsBetween(anchor, hit.col) : colsBetween(hit.col, hit.col));
      }
    }
  };

  const handleDoubleClick = (event: React.MouseEvent<HTMLDivElement>) => {
    // A double-click on a boundary fits what it bounds, which is Excel's way
    // and why the row is tested for first: a row answers only on its edge —
    // where the pointer was already being read as a resize — while a column
    // answers anywhere on its header.
    const target = resizeTargetAt(event.clientX, event.clientY);
    if (target?.kind === "row") {
      autofitRows(rowsForResize(target.index));
      return;
    }
    const hit = hitAt(event.clientX, event.clientY);
    if (hit?.kind === "colHeader") {
      autofitColumn(hit.col);
      return;
    }
    // A cell with text in it opens that text for reading. An empty one has
    // nothing to offer, so it stays an ordinary double-click.
    if (hit?.kind === "cell") openEditor(hit.row, hit.col);
  };

  /**
   * A right-click moves the cursor to what is under the pointer before the menu
   * opens, the way Excel does — so the colour it asks for lands on the cell that
   * was pointed at rather than wherever the cursor happened to be. A press
   * inside the block already selected leaves it alone, which is how a whole
   * block is shaded by right-clicking any one of its cells.
   */
  const handleContextMenu = (event: React.MouseEvent<HTMLDivElement>) => {
    const hit = hitAt(event.clientX, event.clientY);
    if (!hit || hit.kind === "corner") {
      aimMenuAt(null);
      return;
    }
    if (hit.kind === "cell") {
      aimMenuAt(null);
      if (!selectionCovers(selection, hit.row, hit.col)) select(hit.row, hit.col);
      return;
    }
    // A press inside what is selected leaves it alone — a run of whole rows, or
    // a block of cells whose rows the heading then acts on — so a whole stretch
    // can be coloured at once by right-clicking any part of it. A press
    // anywhere else takes the whole heading, which is what makes the colour
    // land on the row rather than on one cell of it.
    if (hit.kind === "rowHeader") {
      aimMenuAt({ kind: "row", index: hit.row });
      const span = headerSpan("row");
      if (!span || hit.row < span[0] || hit.row > span[1]) {
        onSelect(rowsBetween(hit.row, hit.row));
      }
      return;
    }
    aimMenuAt({ kind: "col", index: hit.col });
    const span = headerSpan("col");
    if (!span || hit.col < span[0] || hit.col > span[1]) {
      onSelect(colsBetween(hit.col, hit.col));
    }
  };

  /**
   * The rows or columns a heading acts on: everything the selection reaches
   * along that axis, or null when it reaches none — which is what makes a
   * right-click take the single heading it landed on instead.
   *
   * Where inside that stretch the pointer landed does not matter, and requiring
   * it to be inside was the bug: the heading under the pointer is not always the
   * one the reader meant — the sheet may have scrolled, or the pointer may have
   * drifted a column — and the command then acted on that one column, over the
   * three that were plainly selected on screen.
   *
   * The selection need not be a run of whole rows either. A block of cells
   * hides every row it spans, the way Excel does it. The other way round says
   * nothing: a column selection reaches every row there is.
   */
  const headerSpan = (kind: "row" | "col"): [number, number] | null => {
    const rows = kind === "row";
    let from = Number.POSITIVE_INFINITY;
    let to = Number.NEGATIVE_INFINITY;
    for (const area of areasOf(selection)) {
      const span = spanOfArea(area, rows);
      if (!span) continue;
      from = Math.min(from, span[0]);
      to = Math.max(to, span[1]);
    }
    return from <= to ? [from, to] : null;
  };

  /**
   * Hides or shows a run, for this session only. The selection is left where it
   * is, so Hide followed by Unhide on the same run puts it straight back the way
   * Excel does.
   */
  const setHeaderHidden = (hidden: boolean) => {
    const target = menuTargetRef.current;
    if (!target) return;
    const rows = target.kind === "row";
    let [from, to] = headerSpan(target.kind) ?? [target.index, target.index];
    if (!hidden) {
      /**
       * Showing takes in the hidden run that touches the selection. A hidden
       * row has no height, so there is no way to point at one: the only heading
       * the pointer can reach is the row beside it, and acting on that alone
       * left the run impossible to bring back.
       */
      const set = rows ? hiddenRows : hiddenCols;
      while (from > 0 && set.has(from - 1)) from -= 1;
      while (set.has(to + 1)) to += 1;
    }
    const key = rows ? "rows" : "cols";
    setHiddenOverride((current) => {
      const next = new Map(current[key]);
      for (let index = from; index <= to; index += 1) next.set(index, hidden);
      return { ...current, [key]: next };
    });
  };

  /** The validation rule over a sheet cell, independent of the paint view. */
  const validationAtCell = useCallback(
    (row: number, col: number) => {
      for (const rule of layout?.validations ?? []) {
        for (const range of rule.ranges) {
          if (
            row >= range.row &&
            row < range.row + range.rows &&
            col >= range.col &&
            col < range.col + range.cols
          ) {
            return rule;
          }
        }
      }
      return null;
    },
    [layout],
  );

  /**
   * Opens a list validation's choices when the press was on its arrow — the
   * small button the overlay draws at the active cell's right edge. The list is
   * a viewer's list: this app does not write a value back, so choosing one
   * closes it rather than inventing an edit the file does not have.
   */
  const handleDropdownPress = useCallback(
    (clientX: number, clientY: number): boolean => {
      const at = activeArea(selection);
      if (at.kind !== "cell") return false;
      const rule = validationAtCell(at.row, at.col);
      if (!rule || rule.kind !== "list" || rule.hideDropdown) return false;
      const el = viewportRef.current;
      if (!el) return false;
      const display = displayOfRow(at.row);
      if (display < 0) return false;

      const z = zoomRef.current;
      const rect = el.getBoundingClientRect();
      const sheetX = cols.offsets[at.col] ?? 0;
      const sheetY = rows.offsets[display] ?? 0;
      const w = spanAt(cols.offsets, at.col);
      const h = spanAt(rows.offsets, display);
      // The same geometry the overlay painted the button with.
      const screenX = (headerW + sheetX) * z - (at.col < freeze.cols ? 0 : el.scrollLeft);
      const screenY = (headerH + sheetY) * z - (display < freeze.rows ? 0 : el.scrollTop);
      const size = Math.min(16, Math.max(9, h - 2)) * z;
      const left = rect.left + screenX + (w - size / z - 1) * z;
      const top = rect.top + screenY + (h * z - size) / 2;
      if (clientX < left || clientX > left + size || clientY < top || clientY > top + size) {
        return false;
      }
      void validationListItems(path, sheet.name, rule.formula1).then((items) => {
        if (items.length === 0) return;
        const boxHeight = Math.min(items.length * 24 + 8, 272);
        const y = Math.min(rect.top + screenY + h * z + 2, window.innerHeight - boxHeight - 8);
        const x = Math.min(rect.left + screenX, window.innerWidth - 208);
        setValueList({ x, y, items });
      });
      return true;
    },
    [selection, validationAtCell, displayOfRow, cols, rows, freeze, path, sheet.name, headerW, headerH],
  );

  // A list belongs to a cell on screen; a scroll or a press anywhere else is
  // the end of it. Escape closes it, as it closes every menu.
  useEffect(() => {
    if (!valueList) return;
    const close = () => setValueList(null);
    const key = (event: KeyboardEvent) => {
      if (event.key === "Escape") close();
    };
    const view = viewportRef.current;
    window.addEventListener("pointerdown", close, true);
    window.addEventListener("keydown", key);
    view?.addEventListener("scroll", close);
    return () => {
      window.removeEventListener("pointerdown", close, true);
      window.removeEventListener("keydown", key);
      view?.removeEventListener("scroll", close);
    };
  }, [valueList]);

  const handlePointerDown = (event: React.PointerEvent<HTMLDivElement>) => {
    suppressClick.current = false;
    outlinePress.current = false;
    // Every shortcut in the key handler needs the viewport focused, and a press
    // on the sheet is the moment to take it back. Sheet switches give it focus
    // already, but a press that landed elsewhere — the search box, a toolbar
    // button — left Ctrl+A selecting the page's DOM instead of the sheet.
    viewportRef.current?.focus({ preventScroll: true });
    // And the page's own select-all leaves its highlight behind; nothing on the
    // sheet is a text selection, so a press clears it.
    window.getSelection()?.removeAllRanges();
    if (event.button !== 0) return;
    // Excel's in-cell dropdown: a press on the arrow opens the choices and does
    // not move the selection.
    if (handleDropdownPress(event.clientX, event.clientY)) {
      event.preventDefault();
      return;
    }
    // A press on a grouping button toggles that group and moves nothing else.
    const outline = outlineActionAt(event.clientX, event.clientY);
    if (outline) {
      event.preventDefault();
      // The click that follows the release would otherwise select the row or
      // column the button sits on; a button is chrome, not a header.
      outlinePress.current = true;
      suppressClick.current = true;
      applyOutlineAction(outline);
      return;
    }
    const hit = hitAt(event.clientX, event.clientY);
    if (!hit || hit.kind === "corner") return;
    // An anchor on a gutter picks a run of rows or columns; one on a cell runs a
    // block selection. Both are armed here and resolved by the closing click if
    // the pointer never moves off the anchor.
    if (hit.kind === "rowHeader") {
      rangeDrag.current = { kind: "row", row: hit.row, dragged: false };
    } else if (hit.kind === "colHeader") {
      rangeDrag.current = { kind: "column", col: hit.col, dragged: false };
    } else {
      // Armed at the region's own top-left, as a click is: a block dragged out
      // of a merged cell has to start where Excel says it started, or the
      // region the pointer went down inside is left out of it.
      const merge = mergeAt(hit.row, hit.col);
      rangeDrag.current = {
        kind: "cell",
        row: merge ? merge.row : hit.row,
        col: merge ? merge.col : hit.col,
        dragged: false,
      };
    }
    dragPointer.current = { x: event.clientX, y: event.clientY };
    // Ctrl pulls a block out *beside* what is selected rather than over it, so
    // what is there now is what the drag adds to.
    dragBase.current = event.ctrlKey || event.metaKey ? selection : null;
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
      const base = dragBase.current;
      onSelect(base ? withArea(base, next) : next);
    };
    const onUp = () => {
      suppressClick.current = rangeDrag.current?.dragged ?? false;
      rangeDrag.current = null;
      dragBase.current = null;
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
    const mine = blockPrefix(path, sheet.name);
    for (const [key, block] of cache.current) {
      if (!key.startsWith(mine)) continue;
      const [, rawCol] = key.slice(mine.length).split(SEP);
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

  /**
   * Sizes each of `indices` to the tallest its cells come to once wrapped — the
   * same bargain `autofitColumn` strikes sideways, and drawn from the same
   * loaded blocks, so a row is fitted to what has been read rather than to a
   * sheet nobody has paged through. A cell that does not wrap is one line
   * however long it is, and no row goes below the workbook's own default.
   */
  const autofitRows = (indices: number[]) => {
    const wanted = new Set(indices);
    const tallest = new Map<number, number>(
      indices.map((row) => [row, baseRowH]),
    );
    const mine = blockPrefix(path, sheet.name);
    for (const [key, block] of cache.current) {
      if (!key.startsWith(mine)) continue;
      const [rawRow] = key.slice(mine.length).split(SEP);
      const blockRowStart = Number(rawRow) * BLOCK_ROWS;
      for (let local = 0; local < block.length; local += 1) {
        const row = blockRowStart + local;
        if (!wanted.has(row)) continue;
        const cells = block[local];
        if (!cells) continue;
        for (let col = 0; col < cells.length; col += 1) {
          const text = cells[col]?.text;
          if (!text) continue;
          const style = styleAt(row, col);
          const merge = mergeAt(row, col);
          // A merge is one cell whose text sits in its top-left corner: only
          // that cell has anything to fit, and the rest of what the merge
          // covers is read for its own cells rather than for the text above it.
          if (merge && (merge.row !== row || merge.col !== col)) continue;
          const sizePx = style?.sizePx ?? fontSize;
          const font = fontString(
            cellFontStack(style?.family ?? null, fontFamily),
            sizePx,
            style?.bold ?? false,
            style?.italic ?? false,
          );
          // And it wraps in the whole of itself, not in its first column, or
          // its text would be counted as far more lines than it comes to.
          const first = merge?.col ?? col;
          const last = merge ? merge.col + merge.cols - 1 : col;
          let width = 0;
          for (let c = first; c <= last; c += 1) width += spanAt(cols.offsets, c);
          const lines = layoutLines(
            font,
            text,
            width - CELL_PAD_LEFT - CELL_PAD_RIGHT,
            style?.wrap ?? false,
          );
          tallest.set(
            row,
            Math.max(tallest.get(row) ?? baseRowH, lines.length * sizePx * LINE_HEIGHT),
          );
        }
      }
    }
    setRowOverrides((prev) => {
      const map = new Map(prev);
      for (const [row, height] of tallest) map.set(row, Math.ceil(height));
      return map;
    });
  };

  /**
   * Ctrl+C: the cell's text, or the whole row or column, tab or newline joined.
   *
   * A Ctrl+clicked selection is copied whole, its blocks in reading order. Excel
   * puts a blank line between one block and the next, which is what keeps the
   * seam between them visible in a text paste instead of running the last row of
   * one into the first row of the next.
   */
  const copySelection = useCallback(async () => {
    const sheetRows = Math.max(sheet.rows, layout?.rows ?? 0);
    const sheetCols = Math.max(sheet.cols, layout?.cols ?? 0);

    /** One block's text, and the box the marquee is drawn around. */
    const take = async (area: Area): Promise<{ text: string; rect: MergeRect }> => {
      if (area.kind === "cell") {
        return {
          text: cellAt(area.row, area.col)?.text ?? "",
          rect: { row: area.row, col: area.col, rows: 1, cols: 1 },
        };
      }
      if (area.kind === "row") {
        const text = (
          await axisValues(path, sheet.name, { row: area.row }, 0, sheetCols - 1)
        ).join("\t");
        return { text, rect: { row: area.row, col: 0, rows: area.rows, cols: sheetCols } };
      }
      if (area.kind === "column") {
        const text = (
          await axisValues(path, sheet.name, { col: area.col }, 0, sheetRows - 1)
        ).join("\n");
        return { text, rect: { row: 0, col: area.col, rows: sheetRows, cols: area.cols } };
      }
      const block = await getCells(path, sheet.name, area.row, area.col, area.rows, area.cols);
      return {
        text: block
          .map((line) => line.map((cell) => cell?.text ?? "").join("\t"))
          .join("\n"),
        rect: { row: area.row, col: area.col, rows: area.rows, cols: area.cols },
      };
    };

    let parts: { text: string; rect: MergeRect }[];
    try {
      parts = await Promise.all(areasOf(selection).map(take));
    } catch {
      // Nothing sensible to do if the backend is gone; leave the clipboard be.
      return;
    }
    // The marquee only follows a copy that took, so it never claims a write
    // that the clipboard refused.
    if (await copyText(parts.map((part) => part.text).join("\n\n"))) {
      markCopied({ path, sheet: sheet.name, rects: parts.map((part) => part.rect) });
    }
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
    const visible = el ? Math.floor((el.clientHeight / z - headerH) / baseRowH) : 20;
    const page = Math.max(1, visible - 2 * OVERSCAN_ROWS);
    // A Ctrl+clicked selection answers the arrows as its active cell alone: the
    // other areas come off with the first press, as they do in Excel. The two
    // cells a move is measured from are not the same one: the anchor is Excel's
    // active cell — where a plain move leaves from, and where a block stays
    // pinned — and the edge is the end Shift keeps walking, which is what lets
    // repeated presses accumulate and a press back retrace.
    const anchor = selectionAnchor(selection);
    const shown: Selection =
      selection.kind === "areas"
        ? { kind: "cell", row: anchor.row, col: anchor.col }
        : selection;
    const edge = shown.kind === "range" ? shown.edge : anchor;

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
    if (shown.kind === "row") {
      const slide = (next: number) => {
        const step = Math.max(
          -shown.row,
          Math.min(rowCount - shown.row - shown.rows, next - shown.row),
        );
        onSelect(rowsBetween(shown.anchor + step, shown.edge + step));
      };
      const grow = (next: number) =>
        onSelect(rowsBetween(shown.anchor, clampRow(next)));
      switch (event.key) {
        case "ArrowUp":
          if (event.shiftKey) grow(shown.edge - 1);
          else slide(shown.row - 1);
          break;
        case "ArrowDown":
          if (event.shiftKey) grow(shown.edge + 1);
          else slide(shown.row + 1);
          break;
        case "PageUp":
          if (event.shiftKey) grow(shown.edge - page);
          else slide(shown.row - page);
          break;
        case "PageDown":
          if (event.shiftKey) grow(shown.edge + page);
          else slide(shown.row + page);
          break;
        case "Home":
          select(shown.anchor, 0);
          break;
        case "End":
          select(shown.anchor, lastCol);
          break;
        default:
          return;
      }
      event.preventDefault();
      return;
    }
    if (shown.kind === "column") {
      const slide = (next: number) => {
        const step = Math.max(
          -shown.col,
          Math.min(colCount - shown.col - shown.cols, next - shown.col),
        );
        onSelect(colsBetween(shown.anchor + step, shown.edge + step));
      };
      const grow = (next: number) =>
        onSelect(colsBetween(shown.anchor, clampCol(next)));
      switch (event.key) {
        case "ArrowLeft":
          if (event.shiftKey) grow(shown.edge - 1);
          else slide(shown.col - 1);
          break;
        case "ArrowRight":
          if (event.shiftKey) grow(shown.edge + 1);
          else slide(shown.col + 1);
          break;
        // A column has no vertical page of its own to move by; PageUp and
        // PageDown scroll, which the wheel already does.
        case "Home":
          select(0, shown.anchor);
          break;
        case "End":
          select(lastRow, shown.anchor);
          break;
        default:
          return;
      }
      event.preventDefault();
      return;
    }

    // Shift extends a block rather than collapsing to a cell, walking its edge
    // from where it already is — so repeated presses accumulate, a press back
    // retraces, and the block stays pinned at the anchor throughout. A plain
    // press leaves from the anchor instead, which is Excel's active cell, and
    // collapses the block to whatever cell it lands on.
    const step = event.shiftKey ? edge : anchor;
    const moveTo = (row: number, col: number) => {
      const target = { row: clampRow(row), col: clampCol(col) };
      if (!event.shiftKey) {
        select(target.row, target.col);
        return;
      }
      onSelect(rangeBetween(anchor, target));
    };

    switch (event.key) {
      case "ArrowUp":
        moveTo(step.row - 1, step.col);
        break;
      case "ArrowDown":
        moveTo(step.row + 1, step.col);
        break;
      case "ArrowLeft":
        moveTo(step.row, step.col - 1);
        break;
      case "ArrowRight":
        moveTo(step.row, step.col + 1);
        break;
      case "PageUp":
        moveTo(step.row - page, step.col);
        break;
      case "PageDown":
        moveTo(step.row + page, step.col);
        break;
      case "Home":
        moveTo(step.row, 0);
        break;
      case "End":
        moveTo(step.row, lastCol);
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
        setRowOverrides((prev) => {
          const map = new Map(prev);
          for (const row of active.rows) map.set(row, next);
          return map;
        });
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

      if (py < headerH * z && px >= headerW * z) {
        const x = (px + el.scrollLeft) / z - headerW;
        // The boundary under the pointer is the trailing edge of the column it
        // is in, or the leading edge of that column — which is the previous
        // column's trailing one. Two comparisons, not a scan of every column.
        const col = Math.min(sheetColsGuess - 1, indexAt(cols.offsets, x));
        if (Math.abs(cols.offsets[col + 1] - x) <= grab) return { kind: "col", index: col };
        if (col > 0 && Math.abs(cols.offsets[col] - x) <= grab)
          return { kind: "col", index: col - 1 };
        return null;
      }
      if (px < headerW * z && py >= headerH * z) {
        const y = (py + el.scrollTop) / z - headerH;
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
    [cols, rows, sheetColsGuess, sheetRowOf, displayRowCount, headerW, headerH],
  );

  /**
   * Which rows a grip on one of them acts on. A run the selection already
   * covers is resized whole — dragging one of its row heights is how Excel
   * gives several rows the same height — and a row outside it stands alone.
   */
  const rowsForResize = (row: number): number[] => {
    const covers =
      selection.kind === "row" &&
      row >= selection.row &&
      row < selection.row + selection.rows;
    if (!covers) return [row];
    return Array.from({ length: selection.rows }, (_, offset) => selection.row + offset);
  };

  const handlePointerDownCapture = (event: React.PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) return;
    const target = resizeTargetAt(event.clientX, event.clientY);
    if (!target) return;
    event.preventDefault();
    event.stopPropagation();
    drag.current = {
      kind: target.kind,
      index: target.index,
      rows: target.kind === "row" ? rowsForResize(target.index) : [target.index],
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
      style?.italic ?? false,
    );
    // A cell that shrinks to fit shows all of itself, so there is nothing the
    // hover could be answering.
    if (style?.shrink) return null;
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

  /**
   * Shows a note where its author left it, with Excel's leader to the cell.
   * Shared by the pointer moving over a cell and the notes sidebar pointing at
   * one, so both look the same. False when the note's box or cell cannot be
   * placed right now (a hidden row, say).
   */
  const showNoteFor = (note: NoteCell): boolean => {
    const el = viewportRef.current;
    if (!el) return false;
    const box = noteBox(note);
    const display = displayOfRow(note.row);
    if (!box || display < 0) return false;
    const rect = el.getBoundingClientRect();
    const z = zoomRef.current;
    const w = box.w * z;
    const h = box.h * z;
    // Excel shows a note beside the cell, not at the shape's stored position:
    // the box sits just outside the merged region's top-right corner — 11.25pt
    // across and 9pt up, read off COM while the note was displayed. The shape
    // supplies the size; this supplies the place. With no room on the right it
    // mirrors to the left.
    const merge = mergeAt(note.row, note.col);
    const cellLeft =
      rect.left + (headerW + cols.offsets[note.col] - scrollRef.current.left) * z;
    const mergeEnd = merge === null ? undefined : cols.offsets[merge.col + merge.cols];
    const cellWidth =
      merge === null || mergeEnd === undefined
        ? spanAt(cols.offsets, note.col)
        : mergeEnd - cols.offsets[merge.col];
    const cellRight = cellLeft + cellWidth * z;
    const cellTop =
      rect.top + (headerH + rows.offsets[display] - scrollRef.current.top) * z;
    const gapX = NOTE_BOX_GAP_PX * z;
    const right = cellRight + gapX;
    const flip = right + w > rect.left + el.clientWidth;
    const x = flip ? cellLeft - gapX - w : right;
    const y = cellTop - NOTE_BOX_TOP_GAP_PX * z;
    // The leader always points at the cell's own mark — its top-right corner —
    // whichever side the box had to take.
    const from = { x: flip ? x + w : x, y };
    const to = { x: cellRight, y: cellTop };
    const guard = flip
      ? { x0: x, y0: y, x1: cellLeft, y1: y + h }
      : { x0: cellRight, y0: y, x1: x + w, y1: y + h };
    showNote({ note, x, y, w, h, from, to, guard });
    return true;
  };

  /**
   * A note the reader is pointing at in the notes sidebar, shown in the sheet
   * as hovering its cell would. Leaving the entry takes it away again.
   */
  useEffect(() => {
    if (!previewNote) {
      showNote(null);
      return;
    }
    const note = notesByCell.get(
      cellKey(sheetColsGuess, previewNote.row, previewNote.col),
    );
    if (note) showNoteFor(note);
  }, [previewNote, notesByCell, sheetColsGuess]);

  const updateTooltip = (
    hit: ReturnType<typeof hitAt>,
    clientX: number,
    clientY: number,
  ) => {
    // A note's box is click-through, so a pointer over it is still "over" the
    // cell beneath — which would close the box the moment it was reached. Keep
    // it while the pointer is inside it, or in the strip between it and its
    // cell: crossing that strip is how the pointer gets from one to the other,
    // and the cells it crosses may carry notes of their own.
    const open = noteTipRef.current;
    if (
      open &&
      clientX >= open.guard.x0 &&
      clientX <= open.guard.x1 &&
      clientY >= open.guard.y0 &&
      clientY <= open.guard.y1
    ) {
      cancelNoteHide();
      return;
    }
    const key = hit?.kind === "cell" ? `${hit.row}:${hit.col}` : null;
    // Moving within one cell must not touch React at all.
    if (key === hoverKey.current) return;
    hoverKey.current = key;
    const el = viewportRef.current;
    if (!el || !hit || hit.kind !== "cell") {
      setTip(null);
      scheduleNoteHide();
      return;
    }
    // A note comes first: Excel opens it on hover, and the truncated-text
    // readout is this reader's own aid, so it waits.
    // A merged region is one cell to Excel, so hovering any part of it is
    // hovering the anchor that carries the note.
    const merge = mergeAt(hit.row, hit.col);
    const note = notesByCell.get(
      cellKey(
        sheetColsGuess,
        merge ? merge.row : hit.row,
        merge ? merge.col : hit.col,
      ),
    );
    if (note) {
      scheduleNoteShow(note);
      setTip(null);
      return;
    }
    scheduleNoteHide();
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
      x: rect.left + (headerW + cols.offsets[hit.col] - scrollRef.current.left) * z,
      y:
        rect.top +
        (headerH + rows.offsets[display] - scrollRef.current.top) * z +
        spanAt(rows.offsets, display) * z,
    });
  };

  const handlePointerMove = (event: React.PointerEvent<HTMLDivElement>) => {
    if (drag.current) return;
    const el = viewportRef.current;
    if (!el) return;
    const target = resizeTargetAt(event.clientX, event.clientY);
    const outline = target ? null : outlineActionAt(event.clientX, event.clientY);
    el.style.cursor = target
      ? target.kind === "col"
        ? "col-resize"
        : "row-resize"
      : outline
        ? "pointer"
        : "";
    updateTooltip(
      target || outline ? null : hitAt(event.clientX, event.clientY),
      event.clientX,
      event.clientY,
    );
  };

  const handlePointerLeave = () => {
    hoverKey.current = null;
    setTip(null);
    scheduleNoteHide();
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

  const contentW = (headerW + cols.total) * zoomRef.current;
  const contentH = (headerH + rows.total) * zoomRef.current;

  /**
   * The open cell's text, and the workbook's runs for it when it has any — a
   * rich-text cell draws some of its words in a colour of their own, and a
   * cell's own colour says nothing about them.
   */
  const editingText = editing ? (cellAt(editing.row, editing.col)?.text ?? "") : "";
  const runs = editing
    ? (layout?.richCells?.find(
        (cell) => cell.row === editing.row && cell.col === editing.col,
      )?.runs ?? null)
    : null;

  return (
    <div className="flex h-full min-h-0 flex-col">
      {bar}

      <div className="relative min-h-0 flex-1 overflow-hidden">
        {/*
          The whole sheet answers a right-click: the trigger is the scroller
          itself, so the menu opens wherever on the grid it was pressed, and the
          colour it carries lands on what was under the pointer.
        */}
        <ContextMenu
          onOpenChange={(open) => {
            // The target belongs to the menu that opened on it, and to nothing
            // after that.
            if (!open) aimMenuAt(null);
          }}
        >
          <ContextMenuTrigger asChild>
            <div
              ref={viewportRef}
              tabIndex={0}
              onScroll={handleScroll}
              onClick={handleClick}
              onContextMenu={handleContextMenu}
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
          </ContextMenuTrigger>
          <ContextMenuContent className="min-w-[13.5rem]" onKeyDown={pickByAccelerator}>
            {menuTarget && (
              <>
                <ContextMenuItem
                  data-accelerator={acceleratorOf(t("grid.hide"))}
                  icon={<EyeOff />}
                  onSelect={() => setHeaderHidden(true)}
                >
                  <Accelerated label={t("grid.hide")} />
                </ContextMenuItem>
                <ContextMenuItem
                  data-accelerator={acceleratorOf(t("grid.unhide"))}
                  icon={<Eye />}
                  onSelect={() => setHeaderHidden(false)}
                >
                  <Accelerated label={t("grid.unhide")} />
                </ContextMenuItem>
                <ContextMenuSeparator />
              </>
            )}
            <ContextMenuSub>
              <ContextMenuSubTrigger>{t("palette.fill")}</ContextMenuSubTrigger>
              <ContextMenuSubContent>
                <ColorPalette variant="context" current={markColor} onPick={onMark} />
              </ContextMenuSubContent>
            </ContextMenuSub>
          </ContextMenuContent>
        </ContextMenu>

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
                  filtered={customFilterCols.has(col)}
                  onApply={(values) => {
                    const next = new Map(filters);
                    if (values && values.length > 0) next.set(col, values);
                    else next.delete(col);
                    void applyFilters(next, autoFilter);
                  }}
                />
              </div>
            ))}

          {/* The file's input prompt, beside the active cell. */}
          {promptInfo && !valueList && (
            <div
              ref={promptRef}
              className="pointer-events-none absolute top-0 left-0 max-w-[260px] rounded-[3px] border border-[#8f8f8f] bg-[#ffffe1] px-2 py-1 text-[11px] leading-[1.3] text-[#333] shadow-sm"
              style={{ display: "none", transformOrigin: "0 0" }}
            >
              {promptInfo.title ? <div className="font-bold">{promptInfo.title}</div> : null}
              {promptInfo.text ? (
                <div className="whitespace-pre-wrap">{promptInfo.text}</div>
              ) : null}
            </div>
          )}
          {/*
            The cell's text, lifted out of the painting so it can be selected
            and copied with the mouse, which painted text cannot be. It is a
            field only in looks: every way of changing it is refused below, so
            what is read is what the workbook holds. Escape, Enter and a press
            anywhere else put it back.
          */}
          {editing && (
            <div
              ref={editorRef}
              contentEditable
              suppressContentEditableWarning
              spellCheck={false}
              // Read-only, without `readOnly`: a contenteditable has no such
              // attribute, and its contents have to stay able to carry the
              // per-run colours a plain field could not show at all.
              //
              // The guard is a native `beforeinput` listener (below), not
              // React's `onBeforeInput`: React synthesises that one out of
              // `textInput` and `keypress`, which fire *after* the change has
              // been made, so preventing it changes nothing. Paste and drop are
              // stopped here as well, since they are the two that arrive with
              // their own events.
              onPaste={(event) => event.preventDefault()}
              onDrop={(event) => event.preventDefault()}
              onKeyDown={(event) => {
                if (event.key === "Escape" || event.key === "Enter") {
                  event.preventDefault();
                  setEditing(null);
                }
              }}
              // Put in place by the paint pass, which is the thing that knows
              // where the cell is; hidden until it has been, or it would show up
              // at the origin for a frame.
              style={{ display: "none" }}
              className="pointer-events-auto absolute top-0 left-0 overflow-auto break-words whitespace-pre-wrap outline-none select-text"
            >
              {runs ? (
                runs.map((run, index) => (
                  <span
                    key={index}
                    style={run.color ? { color: run.color } : undefined}
                  >
                    {run.text}
                  </span>
                ))
              ) : (
                <span>{editingText}</span>
              )}
            </div>
          )}
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
                frozenRows={layout.frozenRows ?? 0}
                frozenCols={layout.frozenCols ?? 0}
                zoom={zoom / 100}
                offsets={shapeOffsets}
                onMoveShape={moveShape}
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

        {/* Excel's leader from the note box to its cell, arrowhead on the cell. */}
        {noteTip &&
          (() => {
            const { from, to } = noteTip;
            const dx = to.x - from.x;
            const dy = to.y - from.y;
            const length = Math.hypot(dx, dy) || 1;
            const ux = dx / length;
            const uy = dy / length;
            // The head is six pixels long and three wide either side, in screen
            // pixels: Excel keeps the leader at 1px however far the sheet zooms.
            const baseX = to.x - ux * 6;
            const baseY = to.y - uy * 6;
            return (
              <svg
                className="pointer-events-none fixed inset-0 z-40 h-full w-full"
                aria-hidden
              >
                <line
                  x1={from.x}
                  y1={from.y}
                  x2={baseX}
                  y2={baseY}
                  stroke="#000"
                  strokeWidth={1}
                />
                <polygon
                  points={`${to.x},${to.y} ${baseX - uy * 3},${baseY + ux * 3} ${baseX + uy * 3},${baseY - ux * 3}`}
                  fill="#000"
                />
              </svg>
            );
          })()}

        {/* The note under the pointer, beside its cell. The box keeps the
            file's width but grows downward when the text needs more room than
            the file's box gives it (Excel's own line spacing is 1.25, tighter
            than a browser's default, and its text runs to the very edge). */}
        {noteTip && (
          <div
            className="pointer-events-none fixed z-40 overflow-hidden border border-black bg-[#ffffe1] text-black"
            style={{
              left: noteTip.x,
              top: noteTip.y,
              width: noteTip.w,
              minHeight: noteTip.h,
              lineHeight: 1.25,
              padding: 4 * (zoom / 100),
              fontFamily: noteTip.note.font?.family
                ? `"${noteTip.note.font.family}", ${CJK_FALLBACK}, system-ui, sans-serif`
                : fontFamily
                  ? `"${fontFamily}", ${CJK_FALLBACK}, system-ui, sans-serif`
                  : undefined,
              fontSize: (noteTip.note.font?.sizePx ?? 12) * (zoom / 100),
              fontWeight: noteTip.note.font?.bold ? 600 : 400,
              fontStyle: noteTip.note.font?.italic ? "italic" : undefined,
            }}
          >
            <div className="whitespace-pre-wrap break-words">
              {noteTip.note.text ?? ""}
            </div>
          </div>
        )}

        {/* The values a list validation allows, under its own arrow. */}
        {valueList && (
          <div
            role="listbox"
            className="fixed z-50 max-h-64 min-w-48 overflow-auto border border-border bg-background py-1 text-xs shadow-md"
            style={{ left: valueList.x, top: valueList.y }}
          >
            {valueList.items.map((item, index) => (
              <div
                key={`${index}-${item}`}
                role="option"
                aria-selected={false}
                className="cursor-default truncate px-2 py-1 hover:bg-accent"
                onPointerDown={(event) => event.stopPropagation()}
                onClick={() => setValueList(null)}
              >
                {item}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
