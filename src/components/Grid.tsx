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

import { FilterButton } from "@/components/FilterButton";
import { ShapeLayer, type ShapeDebugFacts } from "@/components/ShapeLayer";
import {
  shapeDebugEnabled,
  subscribeShapeDebug,
} from "@/lib/debugPanels";
import { Button } from "@/components/ui/button";
import { cellRef, columnLabel } from "@/lib/column";
import {
  axisValues,
  filteredRows,
  getCells,
  sheetLayout,
  type AutoFilter,
  type BorderEdge,
  type Cell,
  type CellStyle,
  type ColumnStyle,
  type FilterCriterion,
  type MergeRect,
  type SheetInfo,
  type SheetLayout,
} from "@/lib/excel";
import { copyText } from "@/lib/clipboard";
import { stepZoom } from "@/lib/zoom";
import {
  rangeBetween,
  selectionAnchor,
  selectionCovers,
  selectionIntersects,
  type Selection,
} from "@/lib/selection";
import { cn } from "@/lib/utils";

/** Used only when the file declares no size of its own. */
const FALLBACK_COL_W = 96;
const FALLBACK_ROW_H = 24;
const FALLBACK_FONT_PX = 12;
const MIN_COL_W = 40;
/** Excel's own default row height; also the floor for manual resizing. */
const MIN_ROW_H = 20;
const MAX_COL_W = 640;

/**
 * Excel's own line height. ClosedXML documents the formula Excel uses:
 * `(UnitsPerEm + WinDescent) / UnitsPerEm * fontSize`, which for Calibri
 * (2048, 550) is 1.2686. The slack this leaves above the glyphs is what insets
 * top-aligned text — no artificial padding needed.
 */
const LINE_HEIGHT = 1.269;

/** Fonts covering the Japanese glyphs design documents are full of. */
const CJK_FALLBACK = '"Yu Gothic UI", Meiryo, "MS PGothic", "MS UI Gothic"';

const ROW_HEADER_W = 56;
const COL_HEADER_H = 26;
const OVERSCAN_ROWS = 8;
const OVERSCAN_COLS = 2;

/** Fetch granularity. Cells render one by one, but the cache works in blocks. */
const BLOCK_ROWS = 128;
const BLOCK_COLS = 32;

const RESIZE_HANDLE = 6;

/** How far right a long value may reach before we stop looking. */
const OVERFLOW_COLUMNS = 40;

/**
 * Every cell is drawn one pixel wider and taller than its column and row, so
 * that a border on its right/bottom lands on the same pixel as the neighbour's
 * border on its left/top. CSS draws borders inside the box, so without this the
 * two lines sit side by side and a shared edge reads as double width.
 */
const CELL_BLEED = 1;

/**
 * Excel's in-cell padding, straight out of the accounting behind its column
 * width formula: `PP = 2 * CEILING(MDW/4) + 1`. Calibri and Meiryo UI both have
 * MDW 7, so that is 5px — but it is not 2.5 per side. Each side takes
 * `CEILING(MDW/4)` = 2, and the extra pixel *is the gridline*, so it belongs to
 * the left one: **3 on the left, 2 on the right**. Spreading the 5 evenly put
 * left-aligned text half a pixel in and right-aligned text half a pixel out.
 *
 * The total is unchanged, so the text area stays `列宽 - 5` = `chars * MDW` and
 * nothing that fits on one line starts wrapping. The cell's own 1px right border
 * lands exactly on that gridline pixel, which is why `padding-right` is 2 and not
 * 3.
 */
const CELL_PAD_LEFT = 3;
const CELL_PAD_RIGHT = 2;
const CELL_PAD_X = CELL_PAD_LEFT + CELL_PAD_RIGHT;
const CELL_PAD_CLASS = "pl-[3px] pr-[2px]";

/**
 * Excel marks a single cell with a thin green rule drawn inside it, and leaves
 * the cell unfilled. A whole row or column is the opposite: a faint wash across
 * the cells, with no per-cell outline — outlining every cell tiles the row with
 * green boxes instead of selecting it.
 */
const SELECTION_RING = "inset 0 0 0 1px #217346";
/** A gradient so it can be layered over a cell's own fill, as Excel does. */
const SELECTION_WASH =
  "linear-gradient(rgba(33, 115, 70, 0.07), rgba(33, 115, 70, 0.07))";

/**
 * Excel's print-area mark, as page-break preview draws it: a 3px `#0000D0`
 * rectangle straddling the boundary — measured as 1 screen pixel inside the
 * print area and 2 outside. It is markedly bluer and heavier than the
 * `#3333FF` a cell border uses, so one cannot stand in for the other.
 *
 * `PRINT_AREA_INSIDE` is part of the position and not decoration: a box drawn
 * on the boundary itself would, under `box-sizing: border-box`, spend the
 * whole line on the printed cells and take none of the margin Excel gives it.
 */
const PRINT_AREA_INK = "#0000D0";
const PRINT_AREA_LINE = 3;
const PRINT_AREA_INSIDE = 1;

/**
 * Gutters must stay opaque — the grid scrolls under them. Layering the tint as
 * a gradient over `bg-muted` keeps the base solid, whereas `bg-[#217346]/20`
 * alone let the cells show straight through the header.
 */
const GUTTER_TINT =
  "linear-gradient(rgba(33, 115, 70, 0.18), rgba(33, 115, 70, 0.18))";

/**
 * Excel shades the area outside `_xlnm.Print_Area`. It is opaque and takes
 * precedence over the cell's own fill, but the selection wash still layers on
 * top of it.
 */
/**
 * What page-break preview paints outside the print area. Measured off Excel at
 * 100% — it is a mid grey, noticeably darker than the `#D0D0D0` this started
 * with. It *replaces* a cell's own fill rather than sitting under it, which is
 * why it goes on the cell's background and not on a layer behind the grid.
 */
const OUTSIDE_PRINT = "#A0A0A0";

/** Ranges covering CJK, full-width forms and Hangul — roughly twice as wide. */
const WIDE_CHAR = /[ᄀ-ᅟ⺀-꓏ꥠ-꥿가-퟿豈-﫿︐-︙︰-﹯＀-｠￠-￦]/;

function estimateWidth(text: string): number {
  let width = 0;
  for (const ch of text) width += WIDE_CHAR.test(ch) ? 13.5 : 7.4;
  return width + 18;
}

interface Metrics {
  /** offsets[i] is the pixel position of index i; offsets[count] is the extent. */
  offsets: Float64Array;
  total: number;
}

/** Largest index whose offset is <= pos. */
function indexAt(offsets: Float64Array, pos: number): number {
  let lo = 0;
  let hi = offsets.length - 2;
  if (hi < 0) return 0;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (offsets[mid] <= pos) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

function spanAt(offsets: Float64Array, index: number): number {
  return offsets[index + 1] - offsets[index];
}

function buildMetrics(count: number, sizeOf: (index: number) => number): Metrics {
  const offsets = new Float64Array(count + 1);
  let acc = 0;
  for (let i = 0; i < count; i += 1) {
    offsets[i] = acc;
    acc += sizeOf(i);
  }
  offsets[count] = acc;
  return { offsets, total: acc };
}

/** One `indent` level is three spaces of the cell's font (ECMA-376 §18.8.1). */
const INDENT_SPACES = 3;
/** Only a last resort: this is Calibri's space, and these documents do not use
 * Calibri — ＭＳ ゴシック's is half an em, 游ゴシック's under a third. */
const SPACE_EM = 0.226;

/**
 * The width of a space in a given font, measured rather than assumed.
 *
 * An `indent` level is three spaces, and the font decides how wide one is: at
 * 9pt, ＭＳ ゴシック's space is 6px and 游ゴシック's is 3.4px, so the same
 * `indent="1"` is 18px in one cell and 10px in the next. No font metric is
 * stored in an xlsx, so the only way to know is to ask the renderer that is
 * about to draw it.
 */
const SPACE_ADVANCE = new Map<string, number>();

function spaceAdvancePx(family: string | null, fontPx: number): number {
  const key = `${family ?? ""}|${fontPx}`;
  const known = SPACE_ADVANCE.get(key);
  if (known !== undefined) return known;

  let width = fontPx * SPACE_EM;
  const context = measureContext();
  if (context) {
    context.font = `${fontPx}px "${family || "sans-serif"}"`;
    width = context.measureText(" ").width;
  }
  SPACE_ADVANCE.set(key, width);
  return width;
}

let sharedContext: CanvasRenderingContext2D | null | undefined;

function measureContext(): CanvasRenderingContext2D | null {
  if (sharedContext === undefined) {
    sharedContext = document.createElement("canvas").getContext("2d");
  }
  return sharedContext;
}

/**
 * Excel's column width: the file's own unit, in pixels.
 *
 * A column's width is stored as a multiple of the workbook font's **maximum
 * digit width**, with a fraction of a digit of padding rolled in and the answer
 * floored — `(chars * 256 + 128 / MDW) / 256 * MDW` in the file's 1/256ths. The
 * digit width is 8 for 游ゴシック and 7 for Calibri and nothing in the file says
 * which, so it has to be measured: `chars * 7 + 5` drew every one of these
 * documents' 2.664-character columns 5px too wide, which by column 40 had
 * pushed a callout 120px right of where Excel puts it.
 *
 * Verified against Excel's own `Shape.Left`/`Shape.Width` on a document whose
 * normal font is 游ゴシック 11pt: 860.60px and 238.40px, both exact.
 */
function charsToPx(chars: number, mdw: number): number {
  return Math.floor(((chars * 256 + Math.floor(128 / mdw)) / 256) * mdw);
}

/** The workbook font's maximum digit width, measured rather than assumed. */
const DIGIT_WIDTH = new Map<string, number>();

function digitWidthPx(family: string | null, fontPx: number): number {
  const key = `${family ?? ""}|${fontPx}`;
  const known = DIGIT_WIDTH.get(key);
  if (known !== undefined) return known;

  // Calibri's, as the last resort — `SPACE_EM` carries the same caveat.
  let width = 7;
  const context = measureContext();
  if (context) {
    context.font = `${fontPx}px "${family || "sans-serif"}"`;
    let widest = 0;
    for (let digit = 0; digit <= 9; digit += 1) {
      widest = Math.max(widest, context.measureText(String(digit)).width);
    }
    width = Math.round(widest) || width;
  }
  DIGIT_WIDTH.set(key, width);
  return width;
}

/**
 * The cell's horizontal placement: whatever the file asked for, otherwise
 * OOXML's own default for the kind of value in it.
 *
 * `centerContinuous` is kept apart from `center` because it places itself over
 * its neighbours rather than inside its own cell — see `centerSpans`.
 */
function horizontalOf(
  style: CellStyle | null,
  kind: Cell["kind"] | undefined,
): "left" | "center" | "centerContinuous" | "right" {
  return (
    style?.horizontal ??
    (kind === "number" || kind === "date"
      ? "right"
      : kind === "bool"
        ? "center"
        : "left")
  );
}

/**
 * The padding an `indent` asks for, on the side the text is pinned to — the
 * right for a right-aligned cell, the left otherwise, and neither for a centred
 * one, which Excel leaves alone.
 *
 * `text-indent` cannot express this: it only reaches the first line's *start*
 * edge, where Excel's indent moves a right-aligned cell away from the right.
 * The value returned is the whole side's padding, not just the extra, because an
 * inline style replaces `CELL_PAD_CLASS` rather than adding to it.
 */
function indentPadding(
  style: CellStyle | null,
  kind: Cell["kind"] | undefined,
  /** The workbook's own font, for a cell that states none of its own. */
  fallback: { fontPx: number; family: string | null },
): { paddingLeft?: number; paddingRight?: number } {
  const levels = style?.indent ?? 0;
  if (levels === 0) return {};
  const horizontal = horizontalOf(style, kind);
  if (horizontal === "center" || horizontal === "centerContinuous") return {};

  const fontPx = style?.sizePx ?? fallback.fontPx;
  const family = style?.family ?? fallback.family;
  const extra = Math.round(levels * INDENT_SPACES * spaceAdvancePx(family, fontPx));
  return horizontal === "right"
    ? { paddingRight: CELL_PAD_RIGHT + extra }
    : { paddingLeft: CELL_PAD_LEFT + extra };
}

/**
 * Placement for a cell: whatever the file asked for, otherwise the OOXML
 * defaults — text left, numbers right, and vertically *bottom*. Bottom is the
 * spec default and what Excel shows, so a tall row with one short line puts it
 * at the foot of the cell rather than floating in the middle.
 */
function placementClasses(
  style: CellStyle | null,
  kind: Cell["kind"] | undefined,
): string {
  const horizontal = horizontalOf(style, kind);
  const vertical = style?.vertical ?? "bottom";

  return cn(
    horizontal === "right" && "justify-end text-right",
    (horizontal === "center" || horizontal === "centerContinuous") &&
      "justify-center text-center",
    horizontal === "left" && "justify-start text-left",
    vertical === "top" && "items-start",
    vertical === "middle" && "items-center",
    vertical === "bottom" && "items-end",
    style?.wrap
      ? "whitespace-pre-wrap break-words"
      : "whitespace-pre text-ellipsis",
  );
}

function borderCss(edge: BorderEdge | null): string | undefined {
  if (!edge) return undefined;
  return `${edge.width}px ${edge.line} ${edge.color ?? "var(--border)"}`;
}

/**
 * Two adjacent cells describe the same edge and Excel draws one line for it.
 *
 * The rule, which Excel and Google Sheets share:
 *   1. a border beats no border,
 *   2. a thicker border beats a thinner one,
 *   3. at equal width the line style decides — double, solid, dashed, dotted,
 *   4. everything equal, the *later* cell in reading order wins — which is `b`,
 *      the right/bottom neighbour.
 *
 * `rank` folds 1–3 into one number in Excel's own precedence order.
 */
function heavier(a: BorderEdge | null, b: BorderEdge | null): BorderEdge | null {
  if (!a) return b;
  if (!b) return a;
  return a.rank > b.rank ? a : b;
}

/**
 * Which sides a box paints, and how a shared edge is resolved.
 *
 * Excel's rule for an edge two cells share, as Excel and Google Sheets both
 * behave: a border beats none, a thicker border beats a thinner one, at equal
 * width double > solid > dashed > dotted, and otherwise the later cell in
 * reading order wins. `rank` folds the first three into one number.
 *
 * The resolution is deliberately SYMMETRIC — both cells compute the same
 * `heavier(x, y)` for the edge they share — and each box is drawn one pixel
 * wider and taller than its column and row (see `CELL_BLEED`), so the two lines
 * land on the same pixel. Two identical lines over one pixel are one line, so a
 * shared edge cannot come out double-width whichever cell paints it; the later
 * cell merely paints over the earlier, which is exactly Excel's tie-break.
 */
function edgeStyles(
  styleAt: (row: number, col: number) => CellStyle | null,
  box: { row: number; col: number; rows: number; cols: number },
  bounds: { rows: number; cols: number },
  /**
   * The sides that face another cell of the same merged region (or Center-
   * Across run). Those are not edges at all. Excel draws such a region's
   * outline one *boundary cell* at a time, so its border is not a single line
   * around the region but whatever each of those cells resolves against the
   * neighbour it actually faces — which is why a 2-row merge whose right
   * neighbour carries a border on its top row only gets a line on that row and
   * nothing beside the second one.
   */
  interior?: { top?: boolean; right?: boolean; bottom?: boolean; left?: boolean },
) {
  const own = styleAt(box.row, box.col)?.border ?? null;

  // Off-sheet neighbours must read as "no border". `styleAt` falls through to
  // the column default, which would otherwise invent an edge past the last row.
  const facing = (row: number, col: number) =>
    row < 0 || col < 0 || row >= bounds.rows || col >= bounds.cols
      ? null
      : (styleAt(row, col)?.border ?? null);

  const above = facing(box.row - 1, box.col)?.bottom ?? null;
  const below = facing(box.row + box.rows, box.col)?.top ?? null;
  const left = facing(box.row, box.col - 1)?.right ?? null;
  const right = facing(box.row, box.col + box.cols)?.left ?? null;

  return {
    borderTop: interior?.top
      ? undefined
      : borderCss(heavier(own?.top ?? null, above)),
    borderRight: interior?.right
      ? undefined
      : borderCss(heavier(own?.right ?? null, right)),
    borderBottom: interior?.bottom
      ? undefined
      : borderCss(heavier(own?.bottom ?? null, below)),
    borderLeft: interior?.left
      ? undefined
      : borderCss(heavier(own?.left ?? null, left)),
  };
}

/**
 * Sparse index -> size overrides, as they arrive from the workbook layout.
 *
 * `convert` carries the file's unit across to pixels where the two differ:
 * column widths arrive in characters of the workbook's font, row heights
 * already in pixels.
 */
function toOverrides(
  source: Record<string, number>,
  convert?: (value: number) => number,
): Map<number, number> {
  const out = new Map<number, number>();
  for (const [key, value] of Object.entries(source)) {
    const index = Number(key);
    const size = convert ? convert(value) : value;
    if (Number.isFinite(index) && Number.isFinite(size) && size > 0) {
      out.set(index, size);
    }
  }
  return out;
}

function toFilters(applied: Record<string, string[]> | undefined): Map<number, string[]> {
  const out = new Map<number, string[]>();
  if (!applied) return out;
  for (const [key, values] of Object.entries(applied)) {
    const col = Number(key);
    if (Number.isFinite(col) && values.length > 0) out.set(col, values);
  }
  return out;
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

interface ResizeDrag {
  kind: "col" | "row";
  index: number;
  start: number;
  size: number;
}

/**
 * Excel resolves a cell's style down three levels: the cell's own, then its
 * row's default, then its column's. Reading only the cell level left holes
 * wherever a row style was doing the work.
 */
function styleOfColumn(runs: ColumnStyle[], col: number): number | undefined {
  for (const run of runs) {
    if (col >= run.from && col <= run.to) return run.style;
  }
  return undefined;
}

/** The cell's own style, if a run on its row covers it. */
function styleOfRow(
  runs: Record<string, [number, number, number][]>,
  row: number,
  col: number,
): number | undefined {
  const rowRuns = runs[row];
  if (!rowRuns) return undefined;
  for (const [from, to, style] of rowRuns) {
    if (col >= from && col <= to) return style;
  }
  return undefined;
}

interface GridProps {
  path: string;
  sheet: SheetInfo;
  selection: Selection;
  onSelect: (selection: Selection) => void;
  onDetail?: (cell: Cell | null) => void;
  /** The extent actually on screen, which can exceed calamine's range. */
  onExtent?: (rows: number, cols: number) => void;
  /** View scale as a percentage, the way Excel names it. */
  zoom: number;
  onZoom: (percent: number) => void;
  /**
   * Sits between the filter banner and the grid. Excel stacks them filter →
   * formula bar → cells, so the bar cannot live above the whole view.
   */
  bar?: ReactNode;
}

export function Grid({
  path,
  sheet,
  selection,
  onSelect,
  onDetail,
  onExtent,
  zoom,
  onZoom,
  bar,
}: GridProps) {
  const cache = useRef(new Map<string, Cell[][]>());
  const pending = useRef(new Set<string>());
  const frame = useRef(0);
  const drag = useRef<ResizeDrag | null>(null);
  /** A drag-select in progress: its anchor cell, and whether it has moved. */
  const rangeDrag = useRef<{ row: number; col: number; dragged: boolean } | null>(
    null,
  );
  /** Set when such a drag ends, so its closing click cannot collapse the block. */
  const suppressClick = useRef(false);

  const viewportRef = useRef<HTMLDivElement>(null);
  /** Read inside async callbacks, which cannot close over the current render. */
  const selectionRef = useRef(selection);
  selectionRef.current = selection;

  /**
   * Zoom is a scale on the *rendering*, not on the sheet's own units: column
   * widths, row heights, font sizes and shape geometry stay in the pixels the
   * file gave them, and only the scroller works in zoomed ones. That keeps one
   * coordinate space for everything the grid computes — virtualisation, hit
   * testing, autofit — and converts at the boundary instead.
   *
   * A CSS `transform` was the other way to do this and is not usable here: it
   * takes `position: sticky` out of the scroller, and the two gutters have to
   * stay pinned inside it.
   */
  const zoomFactor = zoom / 100;
  /** Read inside async callbacks, like `selectionRef`. */
  const zoomRef = useRef(zoomFactor);
  zoomRef.current = zoomFactor;

  const [version, setVersion] = useState(0);
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [size, setSize] = useState({ width: 0, height: 0 });
  const [baseColW, setBaseColW] = useState(FALLBACK_COL_W);
  const [baseRowH, setBaseRowH] = useState(FALLBACK_ROW_H);
  const [colOverrides, setColOverrides] = useState<Map<number, number>>(
    () => new Map(),
  );
  const [rowOverrides, setRowOverrides] = useState<Map<number, number>>(
    () => new Map(),
  );
  const [layout, setLayout] = useState<SheetLayout | null>(null);
  /** Which sheet `layout` describes, so a switch never shows the previous one. */
  const [layoutSheet, setLayoutSheet] = useState<string | null>(null);
  const layoutReady = layout !== null && layoutSheet === sheet.name;
  /** Which sheet has all of its on-screen cells in hand. */
  const [painted, setPainted] = useState<string | null>(null);
  /** True once the incoming sheet can replace the one on screen in one step. */
  const paintReady = layoutReady && painted === sheet.name;
  /** Read inside async callbacks, which cannot close over the current render. */
  const paintReadyRef = useRef(paintReady);
  paintReadyRef.current = paintReady;

  /**
   * Dev-only, and only for the readout: how long a sheet switch takes from the
   * click to the frame being on screen, and how much of that was building the
   * frame. The three pieces happen at different points — the switch starts in
   * an effect, the build during a render, the wait ends at a paint — so they
   * are gathered in a ref and published when there is something to say.
   */
  const switchTiming = useRef({ startedAt: 0, buildMs: 0 });
  const [switchWaitMs, setSwitchWaitMs] = useState<number | null>(null);

  useEffect(() => {
    const timing = switchTiming.current;
    if (!paintReady || timing.startedAt === 0) return;
    setSwitchWaitMs(performance.now() - timing.startedAt);
    timing.startedAt = 0;
  }, [paintReady]);

  /** Outside the sheet's print area, which Excel covers with its own shade. */
  const printArea = layout?.printArea ?? null;
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
  /**
   * calamine's range stops at the last cell holding a value, so a column kept
   * alive purely by styling or a merge disappears from it. The layout pass
   * reads the sheet's own extent, and the larger of the two wins.
   */
  const sheetRows = Math.max(sheet.rows, layout?.rows ?? 0);
  const sheetCols = Math.max(sheet.cols, layout?.cols ?? 0);

  /**
   * Which gutter entries are lit. Excel shades every header a selection
   * reaches, so a whole row lights its number, and a block lights the numbers
   * and letters it spans.
   */
  const litRow = useCallback(
    (row: number): boolean => {
      switch (selection.kind) {
        case "row":
          return selection.row === row;
        case "range":
          return row >= selection.row && row < selection.row + selection.rows;
        case "cell":
          return selection.row === row;
        default:
          return false;
      }
    },
    [selection],
  );
  const litColumn = useCallback(
    (col: number): boolean => {
      switch (selection.kind) {
        case "column":
          return selection.col === col;
        case "range":
          return col >= selection.col && col < selection.col + selection.cols;
        case "cell":
          return selection.col === col;
        default:
          return false;
      }
    },
    [selection],
  );
  /** Where keyboard movement starts from, whatever the selection shape. */
  const cursor = selectionAnchor(selection);
  const [fontFamily, setFontFamily] = useState<string | null>(null);
  const [fontSize, setFontSize] = useState(FALLBACK_FONT_PX);
  const [showGrid, setShowGrid] = useState(true);
  const [filters, setFilters] = useState<Map<number, string[]>>(() => new Map());
  const [visibleRows, setVisibleRows] = useState<number[] | null>(null);

  /**
   * Merges clipped to the cells the parser actually reported. A file can
   * declare regions beyond the used range, and those have no offsets to render
   * at — arithmetic on an out-of-bounds entry yields NaN.
   */
  const merges = useMemo(() => {
    const raw = layout?.merges ?? [];
    const out: MergeRect[] = [];
    for (const merge of raw) {
      if (merge.row >= sheetRows || merge.col >= sheetCols) continue;
      out.push({
        row: merge.row,
        col: merge.col,
        rows: Math.min(merge.rows, sheetRows - merge.row),
        cols: Math.min(merge.cols, sheetCols - merge.col),
      });
    }
    return out;
  }, [layout, sheetRows, sheetCols]);

  const autoFilter = layout?.autoFilter ?? null;

  /**
   * Rows the sheet actually shows. A filter hides data rows but never the
   * header or anything above the filter range, so those are kept in front.
   */
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

  const displayRowCount = rowView ? rowView.rows.length : sheetRows;

  /** A row must still fit one line at the workbook's own line height. */
  const minRowH = Math.max(MIN_ROW_H, Math.ceil(fontSize * LINE_HEIGHT));

  const sheetRowOf = useCallback(
    (display: number): number =>
      rowView ? (rowView.rows[display] ?? -1) : display,
    [rowView],
  );

  const displayOfRow = useCallback(
    (sheetRow: number): number =>
      rowView ? (rowView.index.get(sheetRow) ?? -1) : sheetRow,
    [rowView],
  );

  const cols = useMemo(
    () => buildMetrics(sheetCols, (index) => colOverrides.get(index) ?? baseColW),
    [sheetCols, colOverrides, baseColW],
  );

  const rows = useMemo(
    () =>
      buildMetrics(
        displayRowCount,
        (display) => rowOverrides.get(sheetRowOf(display)) ?? baseRowH,
      ),
    [displayRowCount, rowOverrides, baseRowH, sheetRowOf],
  );

  /** A column's left edge, for placing the shapes that float above the cells. */
  const shapeColX = useCallback(
    (col: number) => (col >= 0 && col <= sheetCols ? cols.offsets[col] : cols.total),
    [cols, sheetCols],
  );

  /**
   * A sheet row's top edge, for the same reason. A row a filter has hidden takes
   * the place of the next visible one, so a shape anchored to it shrinks rather
   * than disappearing.
   */
  const shapeRowY = useCallback(
    (row: number) => {
      const last = Math.min(sheetRows, Math.max(0, row) + 4096);
      for (let probe = Math.max(0, row); probe <= last; probe += 1) {
        const display = displayOfRow(probe);
        if (display >= 0 && display < displayRowCount) return rows.offsets[display];
      }
      return rows.total;
    },
    [displayOfRow, displayRowCount, rows, sheetRows],
  );

  const applyFilters = useCallback(
    async (next: Map<number, string[]>, target: AutoFilter | null) => {
      setFilters(next);
      try {
        setVisibleRows(await computeVisibleRows(path, sheet.name, target, next));
      } catch {
        setVisibleRows(null);
      }
    },
    [path, sheet.name],
  );

  useEffect(() => {
    // Bumping on every sheet change would throw away blocks that are still
    // valid: the cache key carries the sheet name, so nothing can collide.
    drag.current = null;
    setPainted(null);
    setVersion((n) => n + 1);
    switchTiming.current.startedAt = performance.now();
    // Nothing else is reset here, on purpose. The previous sheet is still on
    // screen — held frame and all — so anything reset at this moment is a
    // visible change to a drawing that has not been replaced yet. Nulling the
    // layout dropped the print-area shading back to white for as long as the
    // incoming layout took to arrive, which is the blank the eye catches when
    // switching sheets; the column, row and font fallbacks re-laid the held
    // sheet out under the cursor for the same stretch. The incoming sheet's own
    // values arrive with its layout, and `layoutReady` keeps them out of the
    // frame until that sheet's cells are in hand.
    // Deliberately no scrollTo here either: the previous sheet is still on
    // screen, and jogging it would be a visible jolt. The swap re-scrolls to
    // the remembered cursor once the incoming frame is actually up.

    let cancelled = false;
    sheetLayout(path, sheet.name)
      .then((parsed) => {
        if (cancelled) return;
        setLayout(parsed);
        setLayoutSheet(sheet.name);
        // The file's column widths are in characters of this workbook's font,
        // which is measured here — the one place that knows which font it is.
        const colPx = (chars: number) =>
          charsToPx(chars, digitWidthPx(parsed.fontFamily, parsed.fontSizePx));
        setBaseColW(colPx(parsed.defaultColWidthChars));
        setBaseRowH(parsed.defaultRowHeightPx);
        setFontFamily(parsed.fontFamily);
        setFontSize(parsed.fontSizePx);
        setShowGrid(parsed.showGridLines);

        // Aim the viewport at the remembered cursor *before* anything is
        // fetched. Leaving it on the previous sheet's offset would ask the
        // backend for the wrong rows, and the first paint would come up blank.
        const cursor = selectionRef.current;
        if (cursor.kind === "cell") {
          const rowMetrics = buildMetrics(
            parsed.rows,
            (index) => parsed.rowHeightsPx[String(index)] ?? parsed.defaultRowHeightPx,
          );
          const colMetrics = buildMetrics(parsed.cols, (index) =>
            colPx(parsed.colWidthsChars[String(index)] ?? parsed.defaultColWidthChars),
          );
          // No gutter offsets here. The gutters live inside this scroller, so
          // scrolling to `offsets[n]` is what puts cell n just clear of them —
          // adding ROW_HEADER_W / COL_HEADER_H shoved the view 56px right and
          // 26px down even at A1.
          setScroll({
            top: rowMetrics.offsets[cursor.row] ?? 0,
            left: colMetrics.offsets[cursor.col] ?? 0,
          });
        } else {
          setScroll({ top: 0, left: 0 });
        }
        setColOverrides(toOverrides(parsed.colWidthsChars, colPx));
        setRowOverrides(toOverrides(parsed.rowHeightsPx));

        // Excel opens a workbook with its saved filters still applied.
        const saved = toFilters(parsed.autoFilter?.applied);
        if (saved.size > 0) {
          setFilters(saved);
          computeVisibleRows(path, sheet.name, parsed.autoFilter, saved)
            .then((found) => {
              if (!cancelled) setVisibleRows(found);
            })
            .catch(() => {});
        } else {
          // Cleared here rather than when the switch starts, so the filter
          // banner does not collapse and reflow the grid while the previous
          // sheet is still the one on screen.
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
  }, [path, sheet.name]);

  useEffect(() => {
    const el = viewportRef.current;
    if (!el) return;
    const measure = () =>
      setSize({ width: el.clientWidth, height: el.clientHeight });
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  useEffect(() => () => cancelAnimationFrame(frame.current), []);

  /**
   * Ctrl+wheel zooms, as it does in Excel. React's own `wheel` listener rides
   * the root and is passive, and this one has to be able to `preventDefault` —
   * otherwise the webview zooms the whole page as well as the sheet.
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

  const handleScroll = useCallback(() => {
    if (frame.current) return;
    frame.current = requestAnimationFrame(() => {
      frame.current = 0;
      const node = viewportRef.current;
      // The scroller reports zoomed pixels; `scroll` is kept in the sheet's own.
      if (node) {
        const z = zoomRef.current;
        setScroll({ top: node.scrollTop / z, left: node.scrollLeft / z });
      }
    });
  }, []);

  const range = useMemo(() => {
    const top = Math.max(0, scroll.top - COL_HEADER_H);
    const left = Math.max(0, scroll.left - ROW_HEADER_W);
    // The viewport is measured in zoomed pixels and the offsets are not, so the
    // window has to come back across — at 50% a viewport otherwise asked for
    // half the rows it can show.
    const width = size.width / zoomFactor;
    const height = size.height / zoomFactor;
    return {
      rowStart: Math.max(0, indexAt(rows.offsets, top) - OVERSCAN_ROWS),
      rowEnd: Math.min(
        displayRowCount,
        indexAt(rows.offsets, top + height) + 1 + OVERSCAN_ROWS,
      ),
      colStart: Math.max(0, indexAt(cols.offsets, left) - OVERSCAN_COLS),
      colEnd: Math.min(
        sheetCols,
        indexAt(cols.offsets, left + width) + 1 + OVERSCAN_COLS,
      ),
    };
  }, [scroll, size, rows, cols, displayRowCount, sheetCols, zoomFactor]);

  useEffect(() => {
    // Nothing to draw against until the real column widths and font are in.
    if (!layoutReady) return;
    const { rowStart, rowEnd, colStart, colEnd } = range;
    if (rowEnd <= rowStart || colEnd <= colStart) {
      setPainted(sheet.name);
      return;
    }

    const firstCol = Math.floor(colStart / BLOCK_COLS);
    const lastCol = Math.floor((colEnd - 1) / BLOCK_COLS);

    // Blocks are keyed by sheet row, so a filtered view just maps each visible
    // row back before asking for its block.
    const needed = new Set<string>();
    for (let display = rowStart; display < rowEnd; display += 1) {
      const sheetRow = sheetRowOf(display);
      if (sheetRow < 0) continue;
      const block = Math.floor(sheetRow / BLOCK_ROWS);
      for (let bc = firstCol; bc <= lastCol; bc += 1) {
        needed.add(`${sheet.name}#${block}#${bc}`);
      }
    }

    // Only once every block on screen is in hand may the sheet be swapped in —
    // otherwise it lands as bare geometry and the text arrives a beat later.
    const settle = () => {
      if ([...needed].every((key) => cache.current.has(key))) setPainted(sheet.name);
    };

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
          // Repaint as each block lands, so text fills in a beat after a
          // scroll — but not while a sheet switch is holding the previous
          // frame: that build is thrown away, and a switch fetches several
          // blocks, so most of the switch's drawing was being done and
          // discarded. The swap rebuilds once, when `paintReady` turns over.
          if (paintReadyRef.current) setVersion((n) => n + 1);
        })
        .catch(() => {
          cache.current.set(key, []);
        })
        .finally(() => {
          pending.current.delete(key);
          settle();
        });
    }
    settle();
  }, [range, path, sheet.name, sheetRowOf, layoutReady]);

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
      const index =
        styleOfRow(layout.cellRuns, row, col) ??
        layout.rowStyles[row] ??
        styleOfColumn(layout.colStyles, col);
      return index === undefined ? null : (layout.styles[index] ?? null);
    },
    [layout],
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

  useEffect(() => {
    onDetail?.(
      selection.kind === "cell"
        ? cellAt(selection.row, selection.col)
        : null,
    );
  }, [onDetail, cellAt, selection, version]);

  useEffect(() => {
    onExtent?.(sheetRows, sheetCols);
  }, [onExtent, sheetRows, sheetCols]);

  useEffect(() => {
    const el = viewportRef.current;
    // Only chase the cursor once the incoming sheet is the one on screen.
    if (!el || !paintReady || selection.kind !== "cell") return;
    const display = displayOfRow(selection.row);
    if (display < 0) return;
    // Scrolling *up* to a cell means matching its offset, not the offset plus
    // the gutter: the gutter scrolls with the content and would leave the cell
    // tucked behind the sticky header.
    const top = rows.offsets[display] ?? 0;
    const bottom =
      COL_HEADER_H + (rows.offsets[Math.min(display + 1, displayRowCount)] ?? 0);
    const left = cols.offsets[selection.col] ?? 0;
    const right =
      ROW_HEADER_W + (cols.offsets[Math.min(selection.col + 1, sheetCols)] ?? 0);

    // Those targets are the sheet's own pixels and the scroller's are zoomed,
    // so bring the target across rather than converting the scroller's three
    // numbers one by one.
    const topPx = top * zoomFactor;
    const bottomPx = bottom * zoomFactor;
    const leftPx = left * zoomFactor;
    const rightPx = right * zoomFactor;

    if (topPx < el.scrollTop) el.scrollTop = topPx;
    else if (bottomPx > el.scrollTop + el.clientHeight)
      el.scrollTop = bottomPx - el.clientHeight;
    if (leftPx < el.scrollLeft) el.scrollLeft = leftPx;
    else if (rightPx > el.scrollLeft + el.clientWidth)
      el.scrollLeft = rightPx - el.clientWidth;
  }, [
    selection,
    rows,
    cols,
    displayOfRow,
    displayRowCount,
    sheetCols,
    paintReady,
    zoomFactor,
  ]);

  useEffect(() => {
    const onMove = (event: PointerEvent) => {
      const active = drag.current;
      if (!active) return;
      // Pointer deltas arrive in zoomed pixels and the remembered sizes are the
      // sheet's own, so a drag at 50% would otherwise shrink twice as fast.
      const moved =
        (active.kind === "col"
          ? event.clientX - active.start
          : event.clientY - active.start) / zoomRef.current;
      if (active.kind === "col") {
        const next = Math.max(
          MIN_COL_W,
          Math.min(MAX_COL_W, Math.round(active.size + moved)),
        );
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

  const startResize = (
    kind: "col" | "row",
    index: number,
    event: React.PointerEvent<HTMLElement>,
  ) => {
    event.stopPropagation();
    drag.current = {
      kind,
      index,
      start: kind === "col" ? event.clientX : event.clientY,
      size:
        kind === "col" ? spanAt(cols.offsets, index) : spanAt(rows.offsets, index),
    };
    event.currentTarget.setPointerCapture(event.pointerId);
  };

  /** Sizes a column to the widest text currently loaded for it. */
  const autofitColumn = (index: number) => {
    let widest = MIN_COL_W;
    for (const [key, block] of cache.current) {
      const [name, , rawCol] = key.split("#");
      if (name !== sheet.name) continue;
      const blockColStart = Number(rawCol) * BLOCK_COLS;
      if (index < blockColStart || index >= blockColStart + BLOCK_COLS) continue;
      const local = index - blockColStart;
      for (const line of block) {
        const text = line[local]?.text;
        if (text) widest = Math.max(widest, estimateWidth(text));
      }
    }
    setColOverrides((prev) =>
      new Map(prev).set(index, Math.min(MAX_COL_W, Math.ceil(widest))),
    );
  };

  /** The cell under a viewport point, or null over a gutter or off the sheet. */
  const cellAtPoint = useCallback(
    (clientX: number, clientY: number): { row: number; col: number } | null => {
      const el = viewportRef.current;
      if (!el || displayRowCount === 0 || sheetCols === 0) return null;
      // The container's rect does not move with scroll, so add the scroll
      // offset back to translate a viewport point into content coordinates,
      // then drop the two gutters. The point and the scroll offset are both
      // zoomed and the sheet's offsets are not.
      const rect = el.getBoundingClientRect();
      const x = (clientX - rect.left + el.scrollLeft) / zoomFactor - ROW_HEADER_W;
      const y = (clientY - rect.top + el.scrollTop) / zoomFactor - COL_HEADER_H;
      if (x < 0 || y < 0 || x >= cols.total || y >= rows.total) return null;

      const row = sheetRowOf(
        Math.min(displayRowCount - 1, indexAt(rows.offsets, y)),
      );
      if (row < 0) return null;
      return { row, col: Math.min(sheetCols - 1, indexAt(cols.offsets, x)) };
    },
    [cols, rows, sheetCols, displayRowCount, sheetRowOf, zoomFactor],
  );

  const handleClick = (event: React.MouseEvent<HTMLDivElement>) => {
    // A drag that selected a block ends in a click; letting it through would
    // collapse the block back to the cell the pointer stopped on.
    if (suppressClick.current) {
      suppressClick.current = false;
      return;
    }
    const cell = cellAtPoint(event.clientX, event.clientY);
    if (cell) select(cell.row, cell.col);
  };

  const handlePointerDown = (event: React.PointerEvent<HTMLDivElement>) => {
    suppressClick.current = false;
    if (event.button !== 0) return;
    // A gutter is not a cell: its own click picks the row or column.
    if (
      event.target instanceof HTMLElement &&
      event.target.closest("[data-gutter]")
    ) {
      return;
    }
    const cell = cellAtPoint(event.clientX, event.clientY);
    if (!cell) return;
    rangeDrag.current = { row: cell.row, col: cell.col, dragged: false };
    // Otherwise the browser paints a text selection across the whole grid as
    // the pointer sweeps over it.
    document.body.style.userSelect = "none";
  };

  /** Dragging across cells selects the block it spans, as Excel's does. */
  useEffect(() => {
    const onMove = (event: PointerEvent) => {
      const anchor = rangeDrag.current;
      if (!anchor) return;
      const cell = cellAtPoint(event.clientX, event.clientY);
      if (!cell) return;
      if (cell.row === anchor.row && cell.col === anchor.col) return;
      anchor.dragged = true;
      onSelect(rangeBetween(anchor, cell));
    };
    const onUp = () => {
      suppressClick.current = rangeDrag.current?.dragged ?? false;
      rangeDrag.current = null;
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
  }, [cellAtPoint, onSelect]);

  /** Ctrl+C: the cell's text, or the whole row/column, tab or newline joined. */
  const copySelection = useCallback(async () => {
    try {
      if (selection.kind === "cell") {
        await copyText(cellAt(selection.row, selection.col)?.text ?? "");
        return;
      }
      if (selection.kind === "row") {
        const values = await axisValues(
          path,
          sheet.name,
          { row: selection.row },
          0,
          sheetCols - 1,
        );
        await copyText(values.join("\t"));
        return;
      }
      if (selection.kind === "range") {
        // A block is fetched as a block: it need not be on screen, so the
        // cache the grid paints from may not hold it.
        const block = await getCells(
          path,
          sheet.name,
          selection.row,
          selection.col,
          selection.rows,
          selection.cols,
        );
        // Excel puts a tab between columns and a newline between rows.
        await copyText(
          block
            .map((line) => line.map((cell) => cell?.text ?? "").join("\t"))
            .join("\n"),
        );
        return;
      }
      const values = await axisValues(
        path,
        sheet.name,
        { col: selection.col },
        0,
        sheetRows - 1,
      );
      await copyText(values.join("\n"));
    } catch {
      // Nothing sensible to do if the backend is gone; leave the clipboard be.
    }
  }, [selection, cellAt, path, sheet.name, sheetCols, sheetRows]);

  const handleKeyDown = (event: React.KeyboardEvent) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
      event.preventDefault();
      void copySelection();
      return;
    }

    const clampRow = (r: number) => Math.min(sheetRows - 1, Math.max(0, r));
    const clampCol = (c: number) => Math.min(sheetCols - 1, Math.max(0, c));
    const page = Math.max(1, range.rowEnd - range.rowStart - 2 * OVERSCAN_ROWS);

    // A whole row or column moves along its own axis only.
    if (selection.kind === "row") {
      if (event.key === "ArrowUp" || event.key === "ArrowDown") {
        const step = event.key === "ArrowUp" ? -1 : 1;
        onSelect({ kind: "row", row: clampRow(selection.row + step) });
        event.preventDefault();
      }
      return;
    }
    if (selection.kind === "column") {
      if (event.key === "ArrowLeft" || event.key === "ArrowRight") {
        const step = event.key === "ArrowLeft" ? -1 : 1;
        onSelect({ kind: "column", col: clampCol(selection.col + step) });
        event.preventDefault();
      }
      return;
    }

    switch (event.key) {
      case "ArrowUp":
        select(clampRow(cursor.row - 1), cursor.col);
        break;
      case "ArrowDown":
        select(clampRow(cursor.row + 1), cursor.col);
        break;
      case "ArrowLeft":
        select(cursor.row, clampCol(cursor.col - 1));
        break;
      case "ArrowRight":
        select(cursor.row, clampCol(cursor.col + 1));
        break;
      case "PageUp":
        select(clampRow(cursor.row - page), cursor.col);
        break;
      case "PageDown":
        select(clampRow(cursor.row + page), cursor.col);
        break;
      case "Home":
        select(cursor.row, 0);
        break;
      case "End":
        select(cursor.row, clampCol(sheetCols - 1));
        break;
      default:
        return;
    }
    event.preventDefault();
  };

  /** Merged blocks whose anchor row is currently on screen. */
  const mergeView = useMemo(() => {
    if (merges.length === 0) return null;
    const first = sheetRowOf(range.rowStart);
    const last = sheetRowOf(Math.max(range.rowStart, range.rowEnd - 1));
    const lo = Math.min(first, last);
    const hi = Math.max(first, last);

    const shown: MergeRect[] = [];
    const covered = new Set<string>();
    for (const merge of merges) {
      const end = merge.row + merge.rows - 1;
      if (end < lo || merge.row > hi) continue;
      if (displayOfRow(merge.row) < 0) continue;
      shown.push(merge);
      for (let r = Math.max(merge.row, lo); r <= Math.min(end, hi); r += 1) {
        for (let c = merge.col; c < merge.col + merge.cols; c += 1) {
          covered.add(`${r}:${c}`);
        }
      }
    }
    return { shown, covered };
  }, [merges, range, sheetRowOf, displayOfRow]);

  /**
   * Center Across Selection. Excel does not centre such a cell inside itself:
   * it shows the run of neighbours carrying the same `centerContinuous`
   * alignment as ONE block — the text centred across the whole run and the
   * borders between those cells dropped, so a button drawn as five cells reads
   * as a single box. That is the same picture a merge draws, so these runs go
   * through the merge pass too; unlike a merge they stay separate cells, which
   * is why they are collected apart from `merges` rather than folded into it.
   */
  const centerRuns = useMemo(() => {
    const first = sheetRowOf(range.rowStart);
    const last = sheetRowOf(Math.max(range.rowStart, range.rowEnd - 1));
    const lo = Math.min(first, last);
    const hi = Math.max(first, last);

    const shown: (MergeRect & { text: string | null })[] = [];
    const covered = new Set<string>();
    /** Part of a run the grid can draw: formatted so, and not inside a merge. */
    const joinable = (row: number, col: number) =>
      col >= 0 &&
      col < sheetCols &&
      styleAt(row, col)?.horizontal === "centerContinuous" &&
      mergeAt(row, col) === null;

    for (let row = lo; row <= hi; row += 1) {
      if (row < 0 || row >= sheetRows || displayOfRow(row) < 0) continue;
      for (let col = range.colStart; col < range.colEnd; col += 1) {
        if (!joinable(row, col)) continue;
        if (joinable(row, col - 1)) continue; // built once, from its first cell
        let end = col;
        while (joinable(row, end + 1)) end += 1;
        if (end === col) continue; // a lone cell is an ordinary centred cell

        // The run can carry its text in any of its cells, so take the first
        // that has some; the block itself is drawn once, from the first cell.
        let text: string | null = null;
        for (let c = col; c <= end && text === null; c += 1) {
          text = cellAt(row, c)?.text ?? null;
        }

        shown.push({ row, col, rows: 1, cols: end - col + 1, text });
        for (let c = col; c <= end; c += 1) covered.add(`${row}:${c}`);
      }
    }
    return { shown, covered };
  }, [
    range,
    sheetRows,
    sheetCols,
    styleAt,
    mergeAt,
    cellAt,
    displayOfRow,
    sheetRowOf,
  ]);

  const cells = useMemo(() => {
    void version;
    // Also a dependency on purpose: a switch holds this frame until the
    // incoming one can replace it in a single step, and that step is the only
    // build worth making while a switch is in flight.
    void paintReady;
    const buildStartedAt = performance.now();
    const out: ReactNode[] = [];

    // Merged blocks and Center-Across runs are the same picture: one box over
    // several cells, so both are drawn here.
    const spans: { rect: MergeRect; text: string | null }[] = [
      ...(mergeView?.shown ?? []).map((rect) => ({ rect, text: null })),
      ...centerRuns.shown.map(({ text, ...rect }) => ({ rect, text })),
    ];

    for (const { rect: merge, text: runText } of spans) {
      const display = displayOfRow(merge.row);
      if (display < 0) continue;
      let width = 0;
      for (let c = merge.col; c < merge.col + merge.cols; c += 1) {
        width += spanAt(cols.offsets, c);
      }
      let height = 0;
      for (let r = merge.row; r < merge.row + merge.rows; r += 1) {
        const at = displayOfRow(r);
        if (at >= 0) height += spanAt(rows.offsets, at);
      }
      // A merge lights up whole when the selection touches any of its cells.
      const highlights = selectionIntersects(selection, merge);
      const outlined = highlights && selection.kind === "cell";
      const washed = highlights && selection.kind !== "cell";
      const anchor = cellAt(merge.row, merge.col);
      const style = styleAt(merge.row, merge.col);
      const text = runText ?? anchor?.text ?? "";
      out.push(
        <div
          key={`merge:${merge.row}:${merge.col}`}
          title={text || undefined}
          className={cn(
            "absolute flex overflow-hidden",
            CELL_PAD_CLASS,
            placementClasses(style, anchor?.kind),
            outlined
              ? "z-10"
              : showGrid && "border-r border-b border-border/60",
          )}
          style={{
            top: COL_HEADER_H + rows.offsets[display],
            left: ROW_HEADER_W + cols.offsets[merge.col],
            width: width + CELL_BLEED,
            height: height + CELL_BLEED,
            boxShadow: outlined ? SELECTION_RING : undefined,
            backgroundColor: outsidePrint(merge.row, merge.col)
              ? OUTSIDE_PRINT
              : (styleAt(merge.row, merge.col)?.fill ?? undefined),
            backgroundImage: washed ? SELECTION_WASH : undefined,
            // The cell's own font wins over the workbook default.
            fontSize: style?.sizePx ?? undefined,
            fontFamily: style?.family ? `"${style.family}", ${CJK_FALLBACK}` : undefined,
            fontWeight: style?.bold ? 600 : undefined,
            color: style?.color ?? undefined,
            ...indentPadding(style, anchor?.kind, { fontPx: fontSize, family: fontFamily }),
          }}
        >
          {text}
        </div>,
      );

      /**
       * The region's outline, resolved one boundary cell at a time. A merged
       * region (or a Center-Across run) is one block to look at but not to
       * Excel's border rules: each cell on its edge settles its own side against
       * the neighbour outside the region, so the line can stop part-way along
       * it. Sides facing another cell of the same region are not edges and are
       * left out.
       */
      const lastRow = merge.row + merge.rows - 1;
      const lastCol = merge.col + merge.cols - 1;
      for (let r = merge.row; r <= lastRow; r += 1) {
        const at = displayOfRow(r);
        if (at < 0) continue;
        for (let c = merge.col; c <= lastCol; c += 1) {
          if (c < range.colStart || c >= range.colEnd) continue;
          const edges = edgeStyles(
            styleAt,
            { row: r, col: c, rows: 1, cols: 1 },
            { rows: sheetRows, cols: sheetCols },
            {
              top: r > merge.row,
              bottom: r < lastRow,
              left: c > merge.col,
              right: c < lastCol,
            },
          );
          if (
            !edges.borderTop &&
            !edges.borderRight &&
            !edges.borderBottom &&
            !edges.borderLeft
          )
            continue;
          out.push(
            <div
              key={`span-edge:${r}:${c}`}
              aria-hidden
              className="pointer-events-none absolute"
              style={{
                top: COL_HEADER_H + rows.offsets[at],
                left: ROW_HEADER_W + cols.offsets[c],
                width: spanAt(cols.offsets, c) + CELL_BLEED,
                height: spanAt(rows.offsets, at) + CELL_BLEED,
                ...edges,
              }}
            />,
          );
        }
      }
    }

    for (let display = range.rowStart; display < range.rowEnd; display += 1) {
      const sheetRow = sheetRowOf(display);
      if (sheetRow < 0) continue;

      for (let c = range.colStart; c < range.colEnd; c += 1) {
        if (
          mergeView?.covered.has(`${sheetRow}:${c}`) ||
          centerRuns.covered.has(`${sheetRow}:${c}`)
        )
          continue;

        const cell = cellAt(sheetRow, c);
        const highlights = selectionCovers(selection, sheetRow, c);
        const outlined = highlights && selection.kind === "cell";
        const washed = highlights && selection.kind !== "cell";
        const style = styleAt(sheetRow, c);
        const text = cell?.text ?? "";
        const ownWidth = spanAt(cols.offsets, c);

        /**
         * Excel lets a left-aligned, non-wrapped cell's text run on into the
         * cells to its right while those stay empty, stopping at the first one
         * that holds something. Everything else is cut at the cell edge.
         */
        let overflow: number | null = null;
        const align =
          style?.horizontal ??
          (cell?.kind === "number"
            ? "right"
            : cell?.kind === "bool"
              ? "center"
              : "left");
        if (text && !style?.wrap && align === "left" && estimateWidth(text) > ownWidth) {
          let span = ownWidth;
          const limit = Math.min(sheetCols, c + 1 + OVERFLOW_COLUMNS);
          for (let next = c + 1; next < limit; next += 1) {
            if (mergeView?.covered.has(`${sheetRow}:${next}`)) break;
            if ((cellAt(sheetRow, next)?.text ?? "") !== "") break;
            span += spanAt(cols.offsets, next);
          }
          if (span > ownWidth) overflow = span;
        }

        const isFilterHeader =
          autoFilter !== null &&
          sheetRow === autoFilter.row &&
          c >= autoFilter.col &&
          c < autoFilter.col + autoFilter.cols;

        out.push(
          <div
            key={`${sheetRow}:${c}`}
            title={cell?.text || undefined}
            className={cn(
              "absolute flex",
              // The filter button is pinned right, so the header text keeps
              // clear of it the way Excel reserves that corner.
              isFilterHeader ? "pl-[3px] pr-[18px]" : CELL_PAD_CLASS,
              overflow === null && "overflow-hidden",
              placementClasses(style, cell?.kind),
              cell?.kind === "number" && "tabular-nums text-foreground",
              cell?.kind === "bool" && "text-muted-foreground",
              cell?.kind === "error" && "text-destructive",
              outlined
                ? "z-10"
                : showGrid && "border-r border-b border-border/60",
            )}
            style={{
              top: COL_HEADER_H + rows.offsets[display],
              left: ROW_HEADER_W + cols.offsets[c],
              width: spanAt(cols.offsets, c) + CELL_BLEED,
              height: spanAt(rows.offsets, display) + CELL_BLEED,
              boxShadow: outlined ? SELECTION_RING : undefined,
              backgroundColor: outsidePrint(sheetRow, c)
                ? OUTSIDE_PRINT
                : (style?.fill ?? undefined),
              backgroundImage: washed ? SELECTION_WASH : undefined,
              // Text spilling into a neighbour has to paint over that
              // neighbour's borders, the way Excel draws it.
              zIndex: !outlined && overflow !== null ? 1 : undefined,
              fontSize: style?.sizePx ?? undefined,
              fontFamily: style?.family ? `"${style.family}", ${CJK_FALLBACK}` : undefined,
              fontWeight: style?.bold ? 600 : undefined,
              color: style?.color ?? undefined,
              ...indentPadding(style, cell?.kind, { fontPx: fontSize, family: fontFamily }),
              ...edgeStyles(
                styleAt,
                { row: sheetRow, col: c, rows: 1, cols: 1 },
                { rows: sheetRows, cols: sheetCols },
              ),
            }}
          >
            {overflow === null ? (
              text
            ) : (
              // The span outruns its cell so the text can reach the last free
              // column, then clips there — the cell's own box is untouched, so
              // selection and borders stay its own size.
              <span
                className="block shrink-0 overflow-hidden text-ellipsis whitespace-pre"
                style={{ width: Math.max(0, overflow - CELL_PAD_X) }}
              >
                {text}
              </span>
            )}
            {isFilterHeader && autoFilter && (
              <FilterButton
                path={path}
                sheet={sheet.name}
                col={c}
                firstRow={autoFilter.row + 1}
                lastRow={autoFilter.row + autoFilter.rows - 1}
                selected={filters.get(c) ?? []}
                onApply={(values) => {
                  const next = new Map(filters);
                  if (values && values.length > 0) next.set(c, values);
                  else next.delete(c);
                  void applyFilters(next, autoFilter);
                }}
              />
            )}
          </div>,
        );
      }
    }

    // Excel draws one rule around the whole block rather than box per cell.
    if (selection.kind === "range") {
      const top = displayOfRow(selection.row);
      const bottom = displayOfRow(selection.row + selection.rows - 1);
      const right = selection.col + selection.cols - 1;
      if (top >= 0 && bottom >= 0 && right < sheetCols) {
        const x = cols.offsets[selection.col] ?? 0;
        const y = rows.offsets[top] ?? 0;
        out.push(
          <div
            key="selection-range"
            aria-hidden
            className="pointer-events-none absolute z-10"
            style={{
              top: COL_HEADER_H + y,
              left: ROW_HEADER_W + x,
              width: (cols.offsets[right + 1] ?? cols.total) - x,
              height: (rows.offsets[bottom + 1] ?? rows.total) - y,
              boxShadow: SELECTION_RING,
            }}
          />,
        );
      }
    }
    switchTiming.current.buildMs = performance.now() - buildStartedAt;
    return out;
  }, [
    range,
    cellAt,
    selection,
    version,
    paintReady,
    cols,
    rows,
    sheetRowOf,
    displayOfRow,
    mergeView,
    centerRuns,
    autoFilter,
    filters,
    applyFilters,
    styleAt,
    sheetRows,
    sheetCols,
    outsidePrint,
    // The indent fallback reads these, so they belong here even though they come
    // from the same layout `styleAt` already watches.
    fontSize,
    fontFamily,
    path,
    sheet.name,
  ]);

  const colHeaders = useMemo(() => {
    const out: ReactNode[] = [];
    for (let c = range.colStart; c < range.colEnd; c += 1) {
      const isSelected = litColumn(c);
      const width = spanAt(cols.offsets, c);
      out.push(
        <div
          key={c}
          data-gutter="column"
          onClick={(event) => {
            // The scroller's own click handler derives a cell from the click
            // position, and a sticky gutter keeps sitting under the cursor
            // after scrolling — letting the event through re-selected a cell
            // and threw the column selection away again, which is why only
            // single cells ever reached Ctrl+C.
            event.stopPropagation();
            onSelect({ kind: "column", col: c });
          }}
          onDoubleClick={() => autofitColumn(c)}
          className={cn(
            "absolute top-0 flex items-center justify-center border-r border-b border-border text-[11px] font-medium select-none",
            isSelected
              ? "bg-muted font-semibold text-[#217346]"
              : "bg-muted text-muted-foreground",
          )}
          style={{
            left: ROW_HEADER_W + cols.offsets[c],
            width,
            height: COL_HEADER_H,
            backgroundImage: isSelected ? GUTTER_TINT : undefined,
          }}
        >
          <span className="overflow-hidden px-1 text-ellipsis whitespace-nowrap">
            {columnLabel(c)}
          </span>
          <span
            role="separator"
            aria-orientation="vertical"
            aria-label={`Resize column ${columnLabel(c)}`}
            title="Drag to resize, double-click to fit"
            onPointerDown={(event) => startResize("col", c, event)}
            onClick={(event) => event.stopPropagation()}
            // A generous hit area, but nothing painted: Excel signals the
            // boundary with the cursor alone. Filling it drew a 6px dark bar
            // that then stayed put for the whole drag.
            className="absolute top-0 right-0 h-full cursor-col-resize"
            style={{ width: RESIZE_HANDLE }}
          />
        </div>,
      );
    }
    return out;
  }, [range, litColumn, cols, onSelect]);

  const rowHeaders = useMemo(() => {
    const out: ReactNode[] = [];
    for (let display = range.rowStart; display < range.rowEnd; display += 1) {
      const sheetRow = sheetRowOf(display);
      if (sheetRow < 0) continue;
      const isSelected = litRow(sheetRow);
      const height = spanAt(rows.offsets, display);
      out.push(
        <div
          key={sheetRow}
          data-gutter="row"
          onClick={(event) => {
            // See the column header: the scroller would otherwise turn this
            // click back into a cell selection.
            event.stopPropagation();
            onSelect({ kind: "row", row: sheetRow });
          }}
          className={cn(
            "absolute left-0 flex items-center justify-center border-r border-b border-border text-[11px] font-medium tabular-nums select-none",
            isSelected
              ? "bg-muted font-semibold text-[#217346]"
              : "bg-muted text-muted-foreground",
          )}
          style={{
            top: rows.offsets[display],
            width: ROW_HEADER_W,
            height,
            backgroundImage: isSelected ? GUTTER_TINT : undefined,
          }}
        >
          {sheetRow + 1}
          <span
            role="separator"
            aria-orientation="horizontal"
            aria-label={`Resize row ${sheetRow + 1}`}
            title="Drag to resize"
            onPointerDown={(event) => startResize("row", sheetRow, event)}
            onClick={(event) => event.stopPropagation()}
            className="absolute bottom-0 left-0 w-full cursor-row-resize"
            style={{ height: RESIZE_HANDLE }}
          />
        </div>,
      );
    }
    return out;
  }, [range, litRow, rows, sheetRowOf, onSelect]);

  /**
   * A switch has to fetch the incoming sheet's layout before it can draw at
   * all, and that takes long enough to paint an empty grid first. Holding the
   * last complete frame keeps the previous sheet on screen until the new one
   * can replace it in a single step.
   */
  const lastFrame = useRef<{
    width: number;
    height: number;
    content: ReactNode;
    /**
     * The print area the drawn sheet has, held with the frame rather than read
     * live: the shading outside it belongs to the drawing on screen, and taking
     * it from the incoming sheet made it blink white the moment a switch began.
     */
    printArea: MergeRect | null;
  } | null>(null);

  // Move the real viewport only once the new frame is the one on screen. A zoom
  // change re-runs it as well, which is what holds the same cells in view: the
  // remembered offset is the sheet's and the scroller wants it scaled.
  useLayoutEffect(() => {
    if (!paintReady) return;
    const el = viewportRef.current;
    if (!el) return;
    el.scrollTop = scroll.top * zoomFactor;
    el.scrollLeft = scroll.left * zoomFactor;
    // eslint-disable-next-line react-hooks/exhaustive-deps -- only on the swap
  }, [paintReady, zoomFactor]);

  /** Whether the hidden readout is on; Ctrl+Shift+D flips it. */
  const readoutOn = useSyncExternalStore(subscribeShapeDebug, shapeDebugEnabled);

  /**
   * The selected cell, described the way the hidden readout prints it.
   *
   * On most sheets the question is about a cell rather than a shape, and this is
   * where the resolved answer lives: `styleAt` applies the same cell → row →
   * column rule the drawing does, so what is printed here is what is on screen.
   */
  const cellReadout = useMemo(() => {
    const { row, col } = selectionAnchor(selection);
    const display = displayOfRow(row);
    const style = styleAt(row, col);
    const value = cellAt(row, col);
    const edge = (found: BorderEdge | null | undefined) =>
      found
        ? `${found.line} ${found.width}px ${found.color ?? "auto"} r${found.rank}`
        : "none";
    const merge = merges.find(
      (m) =>
        row >= m.row && row < m.row + m.rows && col >= m.col && col < m.col + m.cols,
    );
    return [
      `${cellRef(row, col)}   ${spanAt(cols.offsets, col)}×${spanAt(rows.offsets, display)} @ (${cols.offsets[col] ?? "?"}, ${rows.offsets[display] ?? "?"})`,
      `  value: ${value ? `${value.kind} ${JSON.stringify(value.text)}` : "not loaded"}`,
      style
        ? `  style: font ${style.family ?? "(workbook)"} ${style.sizePx ?? fontSize}px${style.bold ? " bold" : ""}  colour ${style.color ?? "auto"}  fill ${style.fill ?? "none"}  align ${style.horizontal ?? "auto"}/${style.vertical ?? "bottom"}  indent ${style.indent}  wrap ${style.wrap ? "yes" : "no"}`
        : "  style: none",
      `  borders: top ${edge(style?.border?.top)}  right ${edge(style?.border?.right)}  bottom ${edge(style?.border?.bottom)}  left ${edge(style?.border?.left)}`,
      `  merge: ${merge ? `r${merge.row} c${merge.col} ${merge.rows}×${merge.cols}` : "no"}   outside print area: ${outsidePrint(row, col) ? "yes" : "no"}`,
    ].join("\n");
  }, [
    selection,
    styleAt,
    cellAt,
    cols,
    rows,
    displayOfRow,
    merges,
    outsidePrint,
    fontSize,
  ]);

  /**
   * What the hidden readout says about the sheet itself, for the panel
   * `debugPanels.ts` toggles. Gathered here because this is where the layout is
   * known; the digit width is measured again in the loader rather than reused,
   * since that one has to run before the state it would come from is set — and
   * the measurement is cached by font either way.
   */
  const debugFacts: ShapeDebugFacts = useMemo(
    () => ({
      sheet: sheet.name,
      rows: sheetRows,
      cols: sheetCols,
      fontFamily,
      fontSizePx: fontSize,
      digitWidth: digitWidthPx(fontFamily, fontSize),
      defaultColChars: layout?.defaultColWidthChars ?? null,
      defaultColPx: baseColW,
      defaultRowPx: baseRowH,
      showGrid,
      merges: merges.length,
      colWidths: colOverrides,
      printArea: layout?.printArea ?? null,
      cellReadout,
      timings:
        switchWaitMs === null
          ? "no sheet switched yet"
          : `last switch: ${switchWaitMs.toFixed(0)}ms to on screen, of which this frame's build ${switchTiming.current.buildMs.toFixed(0)}ms`,
    }),
    [
      sheet.name,
      sheetRows,
      sheetCols,
      fontFamily,
      fontSize,
      layout,
      baseColW,
      baseRowH,
      showGrid,
      merges,
      colOverrides,
      cellReadout,
      switchWaitMs,
    ],
  );

  if (paintReady) {
    /*
     * Where the print-area mark goes, in the sheet's own pixels. The line
     * straddles the boundary rather than sitting on it: `border-box` spends a
     * border inwards, so a box laid on the boundary would eat the printed edge
     * and take none of the margin Excel gives the mark.
     */
    const area = layout?.printArea ?? null;
    const markInside = PRINT_AREA_INSIDE / zoomFactor;
    const markLine = PRINT_AREA_LINE / zoomFactor;
    const mark = area
      ? {
          // Shifted out by the line's inner share, and widened by the whole
          // line, so the border lands as measured: 1 inside, the rest outside.
          left: ROW_HEADER_W + shapeColX(area.col) - markInside,
          top: COL_HEADER_H + shapeRowY(area.row) - markInside,
          width:
            shapeColX(area.col + area.cols) - shapeColX(area.col) + markLine,
          height:
            shapeRowY(area.row + area.rows) - shapeRowY(area.row) + markLine,
        }
      : null;

    /*
     * The page itself, painted under everything else. The shading outside the
     * print area goes on the cells, so the page has to be laid in first, or the
     * part of a print area that runs past the data would be shaded too — and a
     * print area can reach past the last cell.
     */
    const page = area
      ? {
          left: ROW_HEADER_W + shapeColX(area.col),
          top: COL_HEADER_H + shapeRowY(area.row),
          width: shapeColX(area.col + area.cols) - shapeColX(area.col),
          height: shapeRowY(area.row + area.rows) - shapeRowY(area.row),
        }
      : null;

    lastFrame.current = {
      width: ROW_HEADER_W + cols.total,
      height: COL_HEADER_H + rows.total,
      printArea: area,
      content: (
        <>
          {page && (
            <div
              aria-hidden
              className="absolute bg-background"
              style={page}
            />
          )}

          <div
            className="sticky top-0 z-20"
            style={{ height: COL_HEADER_H, width: ROW_HEADER_W + cols.total }}
          >
            {colHeaders}
          </div>

          <div
            className="sticky left-0 z-20"
            style={{ width: ROW_HEADER_W, height: rows.total }}
          >
            {rowHeaders}
          </div>

          {cells}

          {/*
            The sheet's shapes and pictures, laid over the cells in the same
            coordinate space. `pointer-events-none` because a shape floats above
            the grid without being clickable, so clicks must reach the cells.
          */}
          {layoutReady && (
            <div
              className="pointer-events-none absolute z-[15]"
              style={{ left: ROW_HEADER_W, top: COL_HEADER_H }}
            >
              <ShapeLayer
                shapes={layout.drawings}
                colX={shapeColX}
                rowY={shapeRowY}
                width={cols.total}
                height={rows.total}
                path={path}
                // Null unless the readout is on, so the layer's `memo` still
                // holds in normal use — the facts change with the selection.
                debug={readoutOn ? debugFacts : null}
              />
            </div>
          )}

          {/*
            The print area's own mark. Excel shows the page's extent this way —
            a heavy blue rule on the boundary — and the cells outside it are
            shaded to match, so without the rule the shading has no edge and a
            sheet looks like it is missing part of its border. It rides above
            the shapes and below the gutters, in the same coordinate space.
          */}
          {mark && (
            <div
              aria-hidden
              className="pointer-events-none absolute z-[16]"
              style={{
                ...mark,
                // The mark is a view overlay, so it keeps its weight on screen
                // when the sheet is zoomed rather than growing with the cells —
                // and its position is in those same screen pixels.
                border: `${PRINT_AREA_LINE / zoomFactor}px solid ${PRINT_AREA_INK}`,
              }}
            />
          )}
        </>
      ),
    };
  }
  const drawn = lastFrame.current;

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
              setVisibleRows(null);
            }}
          >
            Clear filters
          </Button>
        </div>
      )}

      {bar}

      <div className="relative min-h-0 flex-1">
        <div
          ref={viewportRef}
          tabIndex={0}
          onScroll={handleScroll}
          onClick={handleClick}
          onPointerDown={handlePointerDown}
          onKeyDown={handleKeyDown}
          // `scrollbar-gutter` keeps the space a scrollbar will occupy, so a
          // switch between a short sheet and a long one cannot shove the whole
          // grid sideways as the bar appears.
          className="h-full overflow-auto outline-none [scrollbar-gutter:stable]"
        style={{
          fontFamily: fontFamily
            ? `"${fontFamily}", ${CJK_FALLBACK}, system-ui, sans-serif`
            : undefined,
          fontSize,
          lineHeight: LINE_HEIGHT,
          /*
           * Excel shades the whole window outside the print area, not just the
           * cells the sheet happens to use: scrolling past the last column or
           * row still lands on grey. The cells carry the shade where they
           * exist, and this picks up everything past them — taken from the
           * frame on screen, so it changes exactly when that does.
           */
          backgroundColor: drawn?.printArea ? OUTSIDE_PRINT : undefined,
        }}
      >
        {/*
          The gutters live inside this one scroller and are held in place with
          position: sticky, so the browser keeps them locked to the data.
          Syncing two separate scrollers through JS drifts as soon as their
          widths differ.
        */}
        <div
          className="relative"
          style={
            drawn
              ? { width: drawn.width, height: drawn.height, zoom: zoomFactor }
              : undefined
          }
        >
          {drawn?.content}
        </div>
        </div>

        {/*
          The corner block is pinned to the scroller's own corner rather than
          riding inside the column-header strip as a nested sticky element —
          that is easy to get subtly wrong, and an overlay cannot drift.
        */}
        <div
          aria-hidden
          className="pointer-events-none absolute top-0 left-0 z-40 border-r border-b border-border bg-muted"
          style={{
            // Outside the zoomed content, so it has to be scaled by hand to
            // stay the size of the gutters it covers.
            width: ROW_HEADER_W * zoomFactor,
            height: COL_HEADER_H * zoomFactor,
          }}
        />
      </div>
    </div>
  );
}
