/**
 * The grid's pure geometry and measurement, lifted out of `Grid.tsx` so the
 * canvas renderer can share it. Nothing here touches React or the DOM beyond
 * an offscreen 2D context used to measure text.
 *
 * The DOM grid keeps its own private copies of all of this until it is deleted;
 * the two must stay in step in the meantime.
 */

import type { BorderEdge, Cell, CellStyle, ColumnStyle } from "@/lib/excel";

/** Used only when the file declares no size of its own. */
export const FALLBACK_COL_W = 96;
export const FALLBACK_ROW_H = 24;
export const FALLBACK_FONT_PX = 12;
export const MIN_COL_W = 40;
/** Excel's own default row height; also the floor for manual resizing. */
export const MIN_ROW_H = 20;
export const MAX_COL_W = 640;

/**
 * Excel's own line height. ClosedXML documents the formula Excel uses:
 * `(UnitsPerEm + WinDescent) / UnitsPerEm * fontSize`, which for Calibri
 * (2048, 550) is 1.2686. The slack this leaves above the glyphs is what insets
 * top-aligned text — no artificial padding needed.
 */
export const LINE_HEIGHT = 1.269;

/** Fonts covering the Japanese glyphs design documents are full of. */
export const CJK_FALLBACK = '"Yu Gothic UI", Meiryo, "MS PGothic", "MS UI Gothic"';

export const ROW_HEADER_W = 56;
export const COL_HEADER_H = 26;
export const OVERSCAN_ROWS = 8;
export const OVERSCAN_COLS = 2;

/** Fetch granularity. Cells render one by one, but the cache works in blocks. */
export const BLOCK_ROWS = 128;
export const BLOCK_COLS = 32;

export const RESIZE_HANDLE = 6;

/** How far right a long value may reach before we stop looking. */
export const OVERFLOW_COLUMNS = 40;

/**
 * Excel's in-cell padding, straight out of the accounting behind its column
 * width formula: `PP = 2 * CEILING(MDW/4) + 1`. Calibri and Meiryo UI both have
 * MDW 7, so that is 5px — but it is not 2.5 per side. Each side takes
 * `CEILING(MDW/4)` = 2, and the extra pixel *is the gridline*, so it belongs to
 * the left one: **3 on the left, 2 on the right**. Spreading the 5 evenly put
 * left-aligned text half a pixel in and right-aligned text half a pixel out.
 *
 * The total is unchanged, so the text area stays `列宽 - 5` = `chars * MDW` and
 * nothing that fits on one line starts wrapping.
 */
export const CELL_PAD_LEFT = 3;
export const CELL_PAD_RIGHT = 2;

/**
 * Excel's selection ink. A single cell is marked with a thin green rule drawn
 * inside it and left unfilled; a whole row or column is the opposite — a faint
 * wash across the cells with no per-cell outline, since outlining every cell
 * tiles the row with green boxes instead of selecting it.
 */
export const ACCENT_INK = "#217346";
export const SELECTION_WASH = "rgba(33, 115, 70, 0.07)";

/**
 * How thick the selection's rule is, in the sheet's own pixels.
 *
 * Excel marks a selection with a heavy line rather than a hairline, and the
 * weight goes *inward*: the outer edge of the box stays on the cell's boundary —
 * the pixels `paintEdges` puts the cell's own border on — so a thicker rule
 * never spills onto the neighbouring cells. The grid's in-cell field wears the
 * same rule while it is open, which is why this lives here and not inside the
 * painter.
 */
export const RING_W = 2;

/**
 * The tint on a lit gutter entry. Gutters must stay opaque — the grid scrolls
 * under them — so the tint is layered over an opaque base rather than being
 * drawn on its own.
 */
export const GUTTER_TINT = "rgba(33, 115, 70, 0.18)";

/**
 * What page-break preview paints outside the print area. Measured off Excel at
 * 100% — it is a mid grey, noticeably darker than the `#D0D0D0` this started
 * with. It *replaces* a cell's own fill rather than sitting under it.
 */
export const OUTSIDE_PRINT = "#A0A0A0";

/**
 * Excel's print-area mark, as page-break preview draws it: a 3px `#0000D0`
 * rectangle straddling the boundary — measured as 1 screen pixel inside the
 * print area and 2 outside. It is markedly bluer and heavier than the
 * `#3333FF` a cell border uses, so one cannot stand in for the other.
 *
 * `PRINT_AREA_INSIDE` is part of the position and not decoration: a box drawn
 * on the boundary itself would spend the whole line on the printed cells and
 * take none of the margin Excel gives it.
 */
export const PRINT_AREA_INK = "#0000D0";
export const PRINT_AREA_LINE = 3;
export const PRINT_AREA_INSIDE = 1;

/** Ranges covering CJK, full-width forms and Hangul — roughly twice as wide. */
const WIDE_CHAR = /[ᄀ-ᅟ⺀-꓏ꥠ-꥿가-퟿豈-﫿︐-︙︰-﹯＀-｠￠-￦]/;

export function estimateWidth(text: string): number {
  let width = 0;
  for (const ch of text) width += WIDE_CHAR.test(ch) ? 13.5 : 7.4;
  return width + 18;
}

export interface Metrics {
  /** offsets[i] is the pixel position of index i; offsets[count] is the extent. */
  offsets: Float64Array;
  total: number;
}

/** Largest index whose offset is <= pos. */
export function indexAt(offsets: Float64Array, pos: number): number {
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

export function spanAt(offsets: Float64Array, index: number): number {
  return offsets[index + 1] - offsets[index];
}

export function buildMetrics(
  count: number,
  sizeOf: (index: number) => number,
): Metrics {
  const offsets = new Float64Array(count + 1);
  let acc = 0;
  for (let i = 0; i < count; i += 1) {
    offsets[i] = acc;
    acc += sizeOf(i);
  }
  offsets[count] = acc;
  return { offsets, total: acc };
}

/**
 * A cell's key in the maps one paint pass builds.
 *
 * A number rather than `"row:col"`, because these are read once per visible
 * cell and the surface and text passes both walk every cell: building the
 * string was costing more than the lookup it fed. Whatever builds a map and
 * whatever reads it have to agree on `cols` — the sheet's column count.
 */
export function cellKey(cols: number, row: number, col: number): number {
  return row * cols + col;
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

export function spaceAdvancePx(family: string | null, fontPx: number): number {
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

/**
 * A 1×1 offscreen context, used only to measure text. Creating a canvas is not
 * free and this is asked often, so it is made once for the whole app.
 */
export function measureContext(): CanvasRenderingContext2D | null {
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
export function charsToPx(chars: number, mdw: number): number {
  return Math.floor(((chars * 256 + Math.floor(128 / mdw)) / 256) * mdw);
}

/** The workbook font's maximum digit width, measured rather than assumed. */
const DIGIT_WIDTH = new Map<string, number>();

export function digitWidthPx(family: string | null, fontPx: number): number {
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
 * its neighbours rather than inside its own cell — see the center-across run
 * pass.
 */
export function horizontalOf(
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
 * The value returned is the whole side's padding, not just the extra, so a
 * caller must fall back to `CELL_PAD_LEFT`/`CELL_PAD_RIGHT` for the sides this
 * leaves out. An absent side is "no indent", not "no padding".
 */
export function indentPadding(
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
export function heavier(
  a: BorderEdge | null,
  b: BorderEdge | null,
): BorderEdge | null {
  if (!a) return b;
  if (!b) return a;
  return a.rank > b.rank ? a : b;
}

export interface BoxEdges {
  top: BorderEdge | null;
  right: BorderEdge | null;
  bottom: BorderEdge | null;
  left: BorderEdge | null;
}

/**
 * The answer for a cell with no border of its own and none reaching in from a
 * neighbour, which is nearly every cell on a full screen. Shared rather than
 * rebuilt, because a fresh object per cell was the biggest source of garbage a
 * paint pass produced. Callers only read it.
 */
const NO_EDGES: BoxEdges = { top: null, right: null, bottom: null, left: null };


/**
 * Which sides a box paints, and how a shared edge is resolved — the canvas
 * counterpart of the DOM renderer's `edgeStyles`, returning the winning
 * `BorderEdge` per side rather than a CSS border string.
 *
 * The resolution is deliberately SYMMETRIC — both cells compute the same
 * `heavier(x, y)` for the edge they share — so a shared edge cannot come out
 * double-width whichever cell paints it. A canvas caller is expected to paint
 * only the top and left sides of an ordinary cell and let the neighbour handle
 * the rest, which is what keeps one line per shared edge without the DOM
 * renderer's one-pixel bleed trick.
 *
 * `interior` marks the sides that face another cell of the same merged region
 * (or Center-Across run). Those are not edges at all. Excel draws such a
 * region's outline one *boundary cell* at a time, so its border is not a single
 * line around the region but whatever each of those cells resolves against the
 * neighbour it actually faces — which is why a 2-row merge whose right
 * neighbour carries a border on its top row only gets a line on that row and
 * nothing beside the second one.
 */
export function resolvedEdges(
  styleAt: (row: number, col: number) => CellStyle | null,
  box: { row: number; col: number; rows: number; cols: number },
  bounds: { rows: number; cols: number },
  interior?: { top?: boolean; right?: boolean; bottom?: boolean; left?: boolean },
): BoxEdges {
  const own = styleAt(box.row, box.col)?.border ?? null;

  // Off-sheet neighbours must read as "no border". `styleAt` falls through to
  // the column default, which would otherwise invent an edge past the last row.
  //
  // Spelled out per side rather than through a `facing` helper: this runs once
  // per visible cell per frame, so the helper was a closure allocated per call —
  // thousands a frame, for a four-way bounds check.
  const { row, col, rows, cols } = box;
  const { rows: maxRows, cols: maxCols } = bounds;
  const below = row + rows;
  const rightOf = col + cols;
  const above = row > 0 && col < maxCols ? (styleAt(row - 1, col)?.border ?? null) : null;
  const under = below < maxRows && col < maxCols ? (styleAt(below, col)?.border ?? null) : null;
  const beside = col > 0 && row < maxRows ? (styleAt(row, col - 1)?.border ?? null) : null;
  const beyond = rightOf < maxCols && row < maxRows ? (styleAt(row, rightOf)?.border ?? null) : null;

  const top = interior?.top ? null : heavier(own?.top ?? null, above?.bottom ?? null);
  const right = interior?.right ? null : heavier(own?.right ?? null, beyond?.left ?? null);
  const bottom = interior?.bottom ? null : heavier(own?.bottom ?? null, under?.top ?? null);
  const left = interior?.left ? null : heavier(own?.left ?? null, beside?.right ?? null);

  if (top === null && right === null && bottom === null && left === null) return NO_EDGES;
  return { top, right, bottom, left };
}

/**
 * Sparse index -> size overrides, as they arrive from the workbook layout.
 *
 * `convert` carries the file's unit across to pixels where the two differ:
 * column widths arrive in characters of the workbook's font, row heights
 * already in pixels.
 */
export function toOverrides(
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

export function toFilters(
  applied: Record<string, string[]> | undefined,
): Map<number, string[]> {
  const out = new Map<number, string[]>();
  if (!applied) return out;
  for (const [key, values] of Object.entries(applied)) {
    const col = Number(key);
    if (Number.isFinite(col) && values.length > 0) out.set(col, values);
  }
  return out;
}

/**
 * Excel resolves a cell's style down three levels: the cell's own, then its
 * row's default, then its column's. Reading only the cell level left holes
 * wherever a row style was doing the work.
 */
export function styleOfColumn(runs: ColumnStyle[], col: number): number | undefined {
  for (const run of runs) {
    if (col >= run.from && col <= run.to) return run.style;
  }
  return undefined;
}

/** The cell's own style, if a run on its row covers it. */
export function styleOfRow(
  runs: Record<string, [number, number, number][]>,
  row: number,
  col: number,
): number | undefined {
  const rowRuns = runs[row];
  if (!rowRuns) return undefined;
  // Bisection, because the runs arrive sorted by first column and never
  // overlap — the parser sorts the row's cells and merges neighbours into runs
  // (see `cell_runs` in layout.rs). This is the paint pass's hottest lookup: a
  // full screen resolves a style per visible cell per frame, and these
  // documents style every cell, so each row carries a run per column. Scanning
  // that list from the top cost a hundred comparisons per call.
  let lo = 0;
  let hi = rowRuns.length - 1;
  while (lo <= hi) {
    const middle = (lo + hi) >> 1;
    const run = rowRuns[middle];
    if (col < run[0]) hi = middle - 1;
    else if (col > run[1]) lo = middle + 1;
    else return run[2];
  }
  return undefined;
}
