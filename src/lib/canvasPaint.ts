/**
 * The canvas painter.
 *
 * It is a pair of pure functions over a plain view model — nothing here reads
 * React state or the DOM — and it is split in two on purpose:
 *
 *   `paintBase`    everything that depends on the scroll position, the zoom,
 *                  the size, and the data: the page and print-area shading,
 *                  gridlines, cell fills, borders, text, merges, and the
 *                  gutter bands behind the row numbers and column letters.
 *   `paintOverlay` the selection, and the gutter labels with their tint.
 *
 * The split is what makes a click cheap. Changing the selected cell repaints
 * only the overlay, leaving the whole cell layer — every string in it — alone.
 * That is the reason this renderer exists at all: the DOM version rebuilt about
 * a thousand elements to move the selection by one cell.
 *
 * `paintBase` deliberately does NOT read the selection. An outlined cell loses
 * its right and bottom gridlines in the DOM renderer, and reproducing that here
 * would make the cell layer depend on the selection and defeat the split;
 * instead the gridlines are drawn as usual and the overlay's ring, which is on
 * the same pixels, covers them.
 *
 * Every line is drawn in the sheet's own pixel space, on the boundary it
 * belongs to and extending away from the cell. Two cells sharing an edge
 * therefore compute the same `heavier()` winner, draw it on the same pixels,
 * and the later of the two simply paints over the earlier — one line per shared
 * edge, with none of the DOM renderer's one-pixel bleed.
 */

import type {
  AutoFilter,
  BorderEdge,
  Cell,
  CellStyle,
  MergeRect,
  RichRun,
} from "@/lib/excel";
import type { CanvasTheme } from "@/lib/canvasTheme";
import { columnLabel } from "@/lib/column";
import {
  ACCENT_INK,
  CELL_PAD_LEFT,
  CELL_PAD_RIGHT,
  COL_HEADER_H,
  GUTTER_TINT,
  OUTSIDE_PRINT,
  ROW_HEADER_W,
  SELECTION_WASH,
  indentPadding,
  horizontalOf,
  resolvedEdges,
  cellKey,
  spanAt,
  type Metrics,
} from "@/lib/gridGeometry";
import { selectionCovers, selectionIntersects, type Selection } from "@/lib/selection";
import {
  cellFontStack,
  fontString,
  layoutLines,
  lineMetrics,
  runSpans,
  textWidth,
} from "@/lib/canvasText";

/** One visible row: its sheet row, and the display index its offset comes from. */
interface ShownRow {
  sheetRow: number;
  display: number;
}

export interface GridView {
  // --- geometry, all in the sheet's own pixels ---
  cols: Metrics;
  rows: Metrics;
  sheetRows: number;
  sheetCols: number;
  /** The overscanned window: display rows, and sheet columns. */
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
  /** The client box in CSS pixels — the canvas's own size. */
  width: number;
  height: number;
  zoom: number;
  /** Scroll offset in the sheet's own pixels, not zoomed. */
  scrollTop: number;
  scrollLeft: number;

  // --- lookups ---
  sheetRowOf: (display: number) => number;
  displayOfRow: (sheetRow: number) => number;
  cellAt: (row: number, col: number) => Cell | null;
  styleAt: (row: number, col: number) => CellStyle | null;
  /**
   * How many times `styleAt` was called this pass, and how many of those the
   * pass-scoped memo answered without doing the work. Owned by the view because
   * the memo lives there; read by the readout to say whether resolving styles
   * is what the surface pass is actually spending its time on.
   */
  styleCalls: { resolved: number; memo: number };
  mergeAt: (row: number, col: number) => MergeRect | null;

  // --- sheet facts ---
  workbookFontFamily: string | null;
  workbookFontSizePx: number;
  showGrid: boolean;
  selection: Selection;
  merges: MergeRect[];
  mergeView: { shown: MergeRect[]; covered: Set<number> };
  centerRuns: { shown: (MergeRect & { text: string | null })[]; covered: Set<number> };
  autoFilter: AutoFilter | null;
  printArea: MergeRect | null;
  /** Whether a cell lies outside the print area, and so takes the grey shade. */
  outsidePrint: (row: number, col: number) => boolean;
  /** Coloured runs, keyed by the text they spell out. Empty for most files. */
  richText: Record<string, RichRun[]>;
  theme: CanvasTheme;
}

/** The room a filter header reserves for its button, so the text keeps clear. */
const FILTER_HEADER_PAD_RIGHT = 18;

const GUTTER_FONT_PX = 11;

/** Draw-space origin: sheet coordinates shifted by the scroll and the gutters. */
function originX(view: GridView): number {
  return ROW_HEADER_W - view.scrollLeft;
}
function originY(view: GridView): number {
  return COL_HEADER_H - view.scrollTop;
}

function fill(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  color: string,
): void {
  if (w <= 0 || h <= 0) return;
  ctx.fillStyle = color;
  ctx.fillRect(x, y, w, h);
}

/** A line as one filled rectangle — see the module note on why not a stroke. */
function line(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  color: string,
): void {
  fill(ctx, x, y, w, h, color);
}

/**
 * A border edge in whatever style the file asked for. Dashes and dots have to
 * be stroked — a dash pattern cannot be faked with rectangles — and `double`,
 * which canvas has no primitive for, is drawn as the two parallel strokes CSS
 * would.
 */
function strokeEdge(
  ctx: CanvasRenderingContext2D,
  edge: BorderEdge,
  orientation: "h" | "v",
  x: number,
  y: number,
  length: number,
  fallback: string,
): void {
  const color = edge.color ?? fallback;
  const w = Math.max(1, edge.width);

  if (edge.line === "solid") {
    if (orientation === "h") line(ctx, x, y, length, w, color);
    else line(ctx, x, y, w, length, color);
    return;
  }
  if (edge.line === "double") {
    // Two strokes of a third the width, with a third of a gap between them.
    const part = Math.max(1, w / 3);
    const offset = w - part;
    if (orientation === "h") {
      line(ctx, x, y, length, part, color);
      line(ctx, x, y + offset, length, part, color);
    } else {
      line(ctx, x, y, part, length, color);
      line(ctx, x + offset, y, part, length, color);
    }
    return;
  }

  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = w;
  ctx.lineCap = "butt";
  // CSS's own patterns, near enough: dashes about three widths, dots one.
  ctx.setLineDash(edge.line === "dotted" ? [w, w] : [3 * w, 2 * w]);
  ctx.beginPath();
  if (orientation === "h") {
    ctx.moveTo(x, y + w / 2);
    ctx.lineTo(x + length, y + w / 2);
  } else {
    ctx.moveTo(x + w / 2, y);
    ctx.lineTo(x + w / 2, y + length);
  }
  ctx.stroke();
  ctx.restore();
}

/**
 * A cell's four sides. Each is drawn on its boundary and extends away from the
 * cell, so the neighbour sharing it lands on the same pixels and the two agree
 * whichever of them paints second.
 */
function paintEdges(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  edges: ReturnType<typeof resolvedEdges>,
  x: number,
  y: number,
  w: number,
  h: number,
): void {
  const fallback = view.theme.border;
  if (edges.top) strokeEdge(ctx, edges.top, "h", x, y, w, fallback);
  if (edges.bottom) strokeEdge(ctx, edges.bottom, "h", x, y + h, w, fallback);
  if (edges.left) strokeEdge(ctx, edges.left, "v", x, y, h, fallback);
  if (edges.right) strokeEdge(ctx, edges.right, "v", x + w, y, h, fallback);
}

/** A selection ring: the four sides drawn just inside the box, as an inset shadow is. */
function ring(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  color: string,
): void {
  if (w <= 0 || h <= 0) return;
  fill(ctx, x, y, w, 1, color);
  fill(ctx, x, y + h - 1, w, 1, color);
  fill(ctx, x, y, 1, h, color);
  fill(ctx, x + w - 1, y, 1, h, color);
}

/** The text colour a cell's kind asks for, when the style states none. */
function textColor(
  style: CellStyle | null,
  kind: Cell["kind"] | undefined,
  theme: CanvasTheme,
): string {
  if (style?.color) return style.color;
  if (kind === "bool") return theme.mutedForeground;
  if (kind === "error") return theme.destructive;
  return theme.foreground;
}

/** The padding an indent adds, with the defaults for the sides it leaves out. */
function pads(
  view: GridView,
  style: CellStyle | null,
  kind: Cell["kind"] | undefined,
): { left: number; right: number } {
  const indent = indentPadding(style, kind, {
    fontPx: view.workbookFontSizePx,
    family: view.workbookFontFamily,
  });
  return {
    left: indent.paddingLeft ?? CELL_PAD_LEFT,
    right: indent.paddingRight ?? CELL_PAD_RIGHT,
  };
}

/** The cell's font, as the paint passes and the fitting both need it. */
function cellFont(view: GridView, style: CellStyle | null): string {
  return fontString(
    cellFontStack(style?.family ?? null, view.workbookFontFamily),
    style?.sizePx ?? view.workbookFontSizePx,
    style?.bold ?? false,
  );
}

/**
 * A cell's or a block's text, placed by the file's alignment, clipped to its
 * text area, and ellipsised or wrapped to fit.
 *
 * `box` is where the text may draw — for a value spilling into empty
 * neighbours it is the wider span, while the cell's own box keeps its size for
 * the fill, the border and the selection.
 */
function paintText(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  box: { x: number; y: number; w: number; h: number },
  text: string,
  style: CellStyle | null,
  kind: Cell["kind"] | undefined,
  padLeft: number,
  padRight: number,
  /** The string's coloured runs, when the file split it into any. */
  runs: RichRun[] | null,
): void {
  if (text === "") return;
  const available = box.w - padLeft - padRight;
  if (available <= 0) return;

  const font = cellFont(view, style);
  const sizePx = style?.sizePx ?? view.workbookFontSizePx;
  const wrap = style?.wrap ?? false;
  const lines = layoutLines(font, text, available, wrap);
  const metrics = lineMetrics(font, sizePx);
  const blockHeight = lines.length * metrics.lineHeight;

  const vertical = style?.vertical ?? "bottom";
  const top =
    vertical === "top"
      ? box.y
      : vertical === "middle"
        ? box.y + (box.h - blockHeight) / 2
        : box.y + box.h - blockHeight;

  const innerX = box.x + padLeft;
  const horizontal = horizontalOf(style, kind);
  let x: number;
  let align: CanvasTextAlign;
  if (horizontal === "right") {
    x = innerX + available;
    align = "right";
  } else if (horizontal === "center" || horizontal === "centerContinuous") {
    x = innerX + available / 2;
    align = "center";
  } else {
    x = innerX;
    align = "left";
  }

  ctx.font = font;
  ctx.textAlign = align;
  ctx.textBaseline = "alphabetic";

  /**
   * Excel clips a value with no room left at the edge of the room it does have —
   * it does not mark the cut with an ellipsis. The clip is only set up when
   * something actually overflows: `clip()` is not cheap and this runs for every
   * cell on screen.
   */
  const tooWide = wrap
    ? false
    : lines.some((line) => line !== "" && textWidth(font, line) > available);
  const needsClip = blockHeight > box.h || tooWide;
  if (needsClip) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(innerX, box.y, available, box.h);
    ctx.clip();
  }

  // CSS centres the leading, so the baseline sits below the half-leading.
  const halfLeading = (metrics.lineHeight - (metrics.ascent + metrics.descent)) / 2;
  const base = textColor(style, kind, view.theme);
  const spans = runs && runs.length > 0 ? runSpans(runs) : null;
  /**
   * Where the next line's text sits in the source. The lines come back in order
   * and contiguous, so the next occurrence from here is the right one — a wrapped
   * break is a prefix of what it cut, and the workspace the layout trimmed off a
   * line's end is simply skipped over.
   */
  let cursor = 0;

  for (let i = 0; i < lines.length; i += 1) {
    const source = lines[i];
    const start = source === "" ? cursor : Math.max(cursor, text.indexOf(source, cursor));
    cursor = start + source.length;
    if (source === "") continue;
    const baseline = top + i * metrics.lineHeight + halfLeading + metrics.ascent;

    if (!spans) {
      ctx.fillStyle = base;
      ctx.fillText(source, x, baseline);
      continue;
    }

    // A line with runs is drawn piece by piece, from its own left edge — the
    // alignment positions the line, the runs only colour it.
    const width = textWidth(font, source);
    let at = align === "right" ? x - width : align === "center" ? x - width / 2 : x;
    ctx.textAlign = "left";
    for (const span of spans) {
      const from = Math.max(start, span.start);
      const to = Math.min(start + source.length, span.end);
      if (to <= from) continue;
      const piece = text.slice(from, to);
      ctx.fillStyle = span.color ?? base;
      ctx.fillText(piece, at, baseline);
      at += textWidth(font, piece);
    }
    ctx.textAlign = align;
  }

  if (needsClip) ctx.restore();
}

/** The page's rectangle in draw coordinates, or null when there is no print area. */
function printRect(
  view: GridView,
  ox: number,
  oy: number,
): { x0: number; y0: number; x1: number; y1: number } | null {
  const area = view.printArea;
  if (!area) return null;
  const first = view.displayOfRow(area.row);
  if (first < 0) return null;
  const last = view.displayOfRow(area.row + area.rows - 1);
  /**
   * Clamped, because a page is allowed to run past the last row and column the
   * sheet formats — that is how a document gives a page room to breathe, and
   * `$A$1:$BC$53` on a 52-row sheet is exactly that. Reading `offsets[53]` off
   * the end of the array gave `undefined`, `y1` came out NaN, and a NaN height
   * turns the page cut-out in `paintBase` into a silent no-op: the flood of
   * OUTSIDE_PRINT stayed, and the whole sheet came up grey.
   */
  const rowAt = (display: number) =>
    view.rows.offsets[Math.min(view.rows.offsets.length - 1, Math.max(0, display))];
  const colAt = (index: number) =>
    view.cols.offsets[Math.min(view.cols.offsets.length - 1, Math.max(0, index))];
  return {
    x0: ox + colAt(area.col),
    x1: ox + colAt(area.col + area.cols),
    y0: oy + rowAt(first),
    y1: last >= 0 ? oy + rowAt(last + 1) : oy + rowAt(first),
  };
}

/**
 * Paints a block's background, shading whatever part of it lies outside the
 * print area.
 *
 * The split matters for a block that straddles the boundary — a Center-Across
 * run of sixty cells starts inside the page and ends outside it. Asking "is
 * this block outside?" once, of its anchor, painted the whole run in the
 * anchor's own fill and left the tail looking like a normal filled cell where
 * Excel's preview shows the page edge.
 */
function fillBlock(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  page: { x0: number; y0: number; x1: number; y1: number } | null,
  x: number,
  y: number,
  w: number,
  h: number,
  fillColor: string | null,
): void {
  if (!page) {
    if (fillColor) fill(ctx, x, y, w, h, fillColor);
    return;
  }
  fill(ctx, x, y, w, h, OUTSIDE_PRINT);
  const ix = Math.max(x, page.x0);
  const iy = Math.max(y, page.y0);
  const ix1 = Math.min(x + w, page.x1);
  const iy1 = Math.min(y + h, page.y1);
  if (ix1 > ix && iy1 > iy) {
    fill(ctx, ix, iy, ix1 - ix, iy1 - iy, fillColor ?? view.theme.background);
  }
}

/** The visible rows, as sheet rows, for a paint pass. */
function shownRows(view: GridView): ShownRow[] {
  const out: ShownRow[] = [];
  for (let display = view.rowStart; display < view.rowEnd; display += 1) {
    const sheetRow = view.sheetRowOf(display);
    if (sheetRow >= 0) out.push({ sheetRow, display });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Base layer
// ---------------------------------------------------------------------------

/**
 * What one base pass had to put on the canvas, for the hidden readout.
 *
 * The counts say how much of the visible grid was real work rather than cells
 * the pass walked and left untouched, and the two timings split the pass where
 * its cost actually divides.
 */
export interface PaintStats {
  /** Display rows the pass drew, and columns — their product is the cell count. */
  rows: number;
  cols: number;
  /** Merged regions and Center-Across runs drawn as a single box. */
  spans: number;
  /** Cells whose value was non-empty, so had to be measured and drawn. */
  texts: number;
  /** Milliseconds over the surface pass: fills, gridlines, borders. */
  surfaceMs: number;
  /** Milliseconds over the text pass. */
  textMs: number;
  /** Style lookups that did the work, and those the memo answered. */
  styleResolved: number;
  styleMemoHits: number;
}

export function paintBase(ctx: CanvasRenderingContext2D, view: GridView): PaintStats {
  const stats: PaintStats = {
    rows: 0,
    cols: 0,
    spans: 0,
    texts: 0,
    surfaceMs: 0,
    textMs: 0,
    styleResolved: 0,
    styleMemoHits: 0,
  };
  const ox = originX(view);
  const oy = originY(view);

  ctx.clearRect(0, 0, view.width, view.height);

  // The page, and the grey Excel shades everything outside the print area with.
  const cellX = ROW_HEADER_W;
  const cellY = COL_HEADER_H;
  const cellW = view.width - ROW_HEADER_W;
  const cellH = view.height - COL_HEADER_H;

  const page = printRect(view, ox, oy);
  fill(ctx, cellX, cellY, cellW, cellH, page ? OUTSIDE_PRINT : view.theme.background);
  if (page) {
    fill(ctx, page.x0, page.y0, page.x1 - page.x0, page.y1 - page.y0, view.theme.background);
  }

  const rows = shownRows(view);
  const spans = collectSpans(view);
  stats.rows = rows.length;
  stats.cols = view.colEnd - view.colStart;
  stats.spans = spans.length;

  /**
   * A merged region's covered cells are not cells — the region is one box. A
   * Center-Across run's are: Excel does not merge them, it only centres one
   * cell's text across the run, so each keeps its own fill, gridline and
   * border. Drawing a run as one box hid the author's own box edges inside it —
   * a `medium` right border on the run's last cell is exactly where such a
   * table ends.
   */
  const merged = (sheetRow: number, col: number) =>
    view.mergeView.covered.has(cellKey(view.sheetCols, sheetRow, col));
  const inRun = (sheetRow: number, col: number) =>
    view.centerRuns.covered.has(cellKey(view.sheetCols, sheetRow, col));

  // Everything the text sits on: fills, print shading, gridlines, borders.
  const surfaceStarted = performance.now();
  for (const span of spans) {
    if (span.kind === "merge") paintSpanSurface(ctx, view, ox, oy, page, span);
  }
  for (const { sheetRow, display } of rows) {
    const y = oy + view.rows.offsets[display];
    const h = spanAt(view.rows.offsets, display);
    for (let c = view.colStart; c < view.colEnd; c += 1) {
      if (merged(sheetRow, c)) continue;
      paintCellSurface(ctx, view, ox, sheetRow, c, y, h, page);
    }
  }
  stats.surfaceMs = performance.now() - surfaceStarted;

  // Then every string, over all of it. Excel lets a left-aligned value run on
  // into the empty cells beside it and draws that spill across their fills and
  // gridlines. Painting the text together with the cell it belongs to put the
  // neighbour's background over the spill instead, which cut such a value off
  // at its own edge whenever the neighbouring cell carried a fill.
  //
  // A run's cells are skipped here and drawn by the run itself, because its
  // text is centred across the whole run rather than inside one of its cells.
  const textStarted = performance.now();
  for (const span of spans) paintSpanText(ctx, view, span);
  for (const { sheetRow, display } of rows) {
    const y = oy + view.rows.offsets[display];
    const h = spanAt(view.rows.offsets, display);
    for (let c = view.colStart; c < view.colEnd; c += 1) {
      if (merged(sheetRow, c) || inRun(sheetRow, c)) continue;
      if (paintCellText(ctx, view, ox, sheetRow, c, y, h)) stats.texts += 1;
    }
  }
  stats.textMs = performance.now() - textStarted;

  paintGutterBands(ctx, view, ox, oy, rows);
  stats.styleResolved = view.styleCalls.resolved;
  stats.styleMemoHits = view.styleCalls.memo;
  return stats;
}

/** One merged region or Center-Across run, measured out for both passes. */
interface Span {
  /** A merge is one box; a run is a text placement over otherwise ordinary cells. */
  kind: "merge" | "run";
  rect: MergeRect;
  /** The run's own text; a merge takes its anchor cell's. */
  text: string | null;
  style: CellStyle | null;
  kindOfCell: Cell["kind"] | undefined;
  /** Where the block lands, in draw coordinates. */
  x: number;
  y: number;
  w: number;
  h: number;
}

function collectSpans(view: GridView): Span[] {
  const ox = originX(view);
  const oy = originY(view);
  const spans: Span[] = [];

  const add = (kind: "merge" | "run", rect: MergeRect, text: string | null) => {
    const display = view.displayOfRow(rect.row);
    if (display < 0) return;
    let w = 0;
    for (let c = rect.col; c < rect.col + rect.cols; c += 1) w += spanAt(view.cols.offsets, c);
    let h = 0;
    for (let r = rect.row; r < rect.row + rect.rows; r += 1) {
      const at = view.displayOfRow(r);
      if (at >= 0) h += spanAt(view.rows.offsets, at);
    }
    spans.push({
      kind,
      rect,
      text,
      style: view.styleAt(rect.row, rect.col),
      kindOfCell: view.cellAt(rect.row, rect.col)?.kind,
      x: ox + view.cols.offsets[rect.col],
      y: oy + view.rows.offsets[display],
      w,
      h,
    });
  };

  for (const rect of view.mergeView.shown) add("merge", rect, null);
  for (const run of view.centerRuns.shown) add("run", run, run.text);
  return spans;
}

/** A span's fill, gridline and the outline Excel resolves cell by cell. */
function paintSpanSurface(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  ox: number,
  oy: number,
  page: { x0: number; y0: number; x1: number; y1: number } | null,
  span: Span,
): void {
  const { x, y, w, h, rect, style } = span;

  fillBlock(ctx, view, page, x, y, w, h, style?.fill ?? null);

  if (view.showGrid) {
    line(ctx, x + w - 1, y, 1, h, view.theme.gridline);
    line(ctx, x, y + h - 1, w, 1, view.theme.gridline);
  }

  // The region's outline, resolved one boundary cell at a time, exactly as the
  // DOM renderer did — sides facing another cell of the same region are not
  // edges and are left out.
  const lastRow = rect.row + rect.rows - 1;
  const lastCol = rect.col + rect.cols - 1;
  for (let r = rect.row; r <= lastRow; r += 1) {
    const at = view.displayOfRow(r);
    if (at < 0) continue;
    const cy = oy + view.rows.offsets[at];
    const ch = spanAt(view.rows.offsets, at);
    for (let c = rect.col; c <= lastCol; c += 1) {
      if (c < view.colStart || c >= view.colEnd) continue;
      paintEdges(
        ctx,
        view,
        resolvedEdges(
          view.styleAt,
          { row: r, col: c, rows: 1, cols: 1 },
          { rows: view.sheetRows, cols: view.sheetCols },
          { top: r > rect.row, bottom: r < lastRow, left: c > rect.col, right: c < lastCol },
        ),
        ox + view.cols.offsets[c],
        cy,
        spanAt(view.cols.offsets, c),
        ch,
      );
    }
  }
}

function paintSpanText(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  span: Span,
): void {
  const text = span.text ?? view.cellAt(span.rect.row, span.rect.col)?.text ?? "";
  const pad = pads(view, span.style, span.kindOfCell);
  paintText(
    ctx,
    view,
    { x: span.x, y: span.y, w: span.w, h: span.h },
    text,
    span.style,
    span.kindOfCell,
    pad.left,
    pad.right,
    view.richText[text] ?? null,
  );
}

/** One ordinary cell's fill, gridline and borders — the pass before the text. */
function paintCellSurface(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  ox: number,
  sheetRow: number,
  col: number,
  y: number,
  h: number,
  page: { x0: number; y0: number; x1: number; y1: number } | null,
): void {
  const x = ox + view.cols.offsets[col];
  const w = spanAt(view.cols.offsets, col);
  const style = view.styleAt(sheetRow, col);
  const own = style?.fill ?? null;

  if (view.outsidePrint(sheetRow, col)) {
    // Already OUTSIDE_PRINT: paintBase flooded the whole cell area with it and
    // then cut the page rect back out in the sheet's own background, so shading
    // the cell again would only repaint the colour that is underneath it. The
    // one page rect that cannot be measured — its first row hidden — leaves no
    // flood to lean on, and that cell does have to shade itself.
    if (!page) fill(ctx, x, y, w, h, OUTSIDE_PRINT);
  } else if (own) {
    fill(ctx, x, y, w, h, own);
  }

  // The hairline grid, before the borders so a border wins where both land.
  if (view.showGrid) {
    line(ctx, x + w - 1, y, 1, h, view.theme.gridline);
    line(ctx, x, y + h - 1, w, 1, view.theme.gridline);
  }

  paintEdges(
    ctx,
    view,
    resolvedEdges(
      view.styleAt,
      { row: sheetRow, col, rows: 1, cols: 1 },
      { rows: view.sheetRows, cols: view.sheetCols },
    ),
    x,
    y,
    w,
    h,
  );
}

function paintCellText(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  ox: number,
  sheetRow: number,
  col: number,
  y: number,
  h: number,
): boolean {
  const cell = view.cellAt(sheetRow, col);
  const text = cell?.text ?? "";
  if (text === "") return false;

  const x = ox + view.cols.offsets[col];
  const w = spanAt(view.cols.offsets, col);
  const style = view.styleAt(sheetRow, col);

  const isFilterHeader =
    view.autoFilter !== null &&
    sheetRow === view.autoFilter.row &&
    col >= view.autoFilter.col &&
    col < view.autoFilter.col + view.autoFilter.cols;

  let padLeft = CELL_PAD_LEFT;
  let padRight = CELL_PAD_RIGHT;
  if (isFilterHeader) padRight = FILTER_HEADER_PAD_RIGHT;
  else {
    const pad = pads(view, style, cell?.kind);
    padLeft = pad.left;
    padRight = pad.right;
  }

  // A left-aligned value with no wrap runs on into the empty cells to its
  // right, stopping at the first that holds something.
  //
  // Only two things stop it: a merged region's cells, which are not cells at
  // all, and a neighbour that actually holds a value. A run of Center-Across
  // cells does NOT stop it — those are ordinary empty cells that merely share
  // an alignment, and Excel keeps spilling through them. Treating the run as
  // the barrier the DOM renderer never made it cut the value off at its own
  // edge on every label followed by an empty centred field.
  let box = { x, y, w, h };
  if (!style?.wrap && horizontalOf(style, cell?.kind) === "left") {
    if (textWidth(cellFont(view, style), text) > w - padLeft - padRight) {
      let reached = w;
      for (let next = col + 1; next < view.sheetCols; next += 1) {
        if (view.mergeView.covered.has(cellKey(view.sheetCols, sheetRow, next))) break;
        if ((view.cellAt(sheetRow, next)?.text ?? "") !== "") break;
        reached += spanAt(view.cols.offsets, next);
      }
      if (reached > w) box = { x, y, w: reached, h };
    }
  }

  paintText(ctx, view, box, text, style, cell?.kind, padLeft, padRight, view.richText[text] ?? null);
  return true;
}

/** The opaque strips behind the row numbers and column letters. */
function paintGutterBands(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  ox: number,
  oy: number,
  rows: ShownRow[],
): void {
  const muted = view.theme.muted;
  const grid = view.theme.gridline;

  fill(ctx, 0, COL_HEADER_H, ROW_HEADER_W, view.height - COL_HEADER_H, muted);
  fill(ctx, ROW_HEADER_W, 0, view.width - ROW_HEADER_W, COL_HEADER_H, muted);
  fill(ctx, 0, 0, ROW_HEADER_W, COL_HEADER_H, muted);

  for (const { display } of rows) {
    const y = oy + view.rows.offsets[display];
    const h = spanAt(view.rows.offsets, display);
    line(ctx, 0, y + h - 1, ROW_HEADER_W, 1, grid);
  }
  for (let c = view.colStart; c < view.colEnd; c += 1) {
    const x = ox + view.cols.offsets[c];
    const w = spanAt(view.cols.offsets, c);
    line(ctx, x + w - 1, 0, 1, COL_HEADER_H, grid);
  }
  // The two sides separating the gutters from the cells, and the corner's own.
  line(ctx, ROW_HEADER_W - 1, 0, 1, view.height, grid);
  line(ctx, 0, COL_HEADER_H - 1, view.width, 1, grid);
}

// ---------------------------------------------------------------------------
// Overlay layer
// ---------------------------------------------------------------------------

export function paintOverlay(ctx: CanvasRenderingContext2D, view: GridView): void {
  ctx.clearRect(0, 0, view.width, view.height);
  const ox = originX(view);
  const oy = originY(view);
  const rows = shownRows(view);

  // Excel marks one cell with an outline and anything larger with a fill,
  // never both.
  const single = view.selection.kind === "cell";
  for (const { sheetRow, display } of rows) {
    const y = oy + view.rows.offsets[display];
    const h = spanAt(view.rows.offsets, display);
    for (let c = view.colStart; c < view.colEnd; c += 1) {
      if (!selectionCovers(view.selection, sheetRow, c)) continue;
      // A merged region's cells are not cells and are handled below; a
      // Center-Across run's are, so a selected row lights every one of them.
      if (view.mergeView.covered.has(cellKey(view.sheetCols, sheetRow, c))) continue;
      const x = ox + view.cols.offsets[c];
      const w = spanAt(view.cols.offsets, c);
      if (single) ring(ctx, x, y, w, h, ACCENT_INK);
      else fill(ctx, x, y, w, h, SELECTION_WASH);
    }
  }

  // A merged region lights up whole when the selection touches any of it. A
  // Center-Across run does not: it is not merged, and its cells were washed
  // one by one above.
  for (const rect of view.mergeView.shown) {
    if (!selectionIntersects(view.selection, rect)) continue;
    const display = view.displayOfRow(rect.row);
    if (display < 0) continue;
    const x = ox + view.cols.offsets[rect.col];
    const y = oy + view.rows.offsets[display];
    let w = 0;
    for (let c = rect.col; c < rect.col + rect.cols; c += 1) w += spanAt(view.cols.offsets, c);
    let h = 0;
    for (let r = rect.row; r < rect.row + rect.rows; r += 1) {
      const at = view.displayOfRow(r);
      if (at >= 0) h += spanAt(view.rows.offsets, at);
    }
    if (single) ring(ctx, x, y, w, h, ACCENT_INK);
    else fill(ctx, x, y, w, h, SELECTION_WASH);
  }

  // Excel draws one rule around the whole block rather than a box per cell.
  if (view.selection.kind === "range") {
    const top = view.displayOfRow(view.selection.row);
    const bottom = view.displayOfRow(view.selection.row + view.selection.rows - 1);
    const right = Math.min(view.sheetCols - 1, view.selection.col + view.selection.cols - 1);
    if (top >= 0 && bottom >= 0) {
      const x = ox + view.cols.offsets[view.selection.col];
      const y = oy + view.rows.offsets[top];
      ring(
        ctx,
        x,
        y,
        view.cols.offsets[right + 1] - view.cols.offsets[view.selection.col],
        view.rows.offsets[bottom + 1] - view.rows.offsets[top],
        ACCENT_INK,
      );
    }
  }

  paintGutterLabels(ctx, view, ox, oy, rows);
}

/**
 * The row numbers and column letters, and the tint on the ones a selection
 * reaches. They live on the overlay rather than with the gutter bands so the
 * tint can lie under the label and not over it — a translucent green across
 * dark text would read as a smudge.
 */
function paintGutterLabels(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  ox: number,
  oy: number,
  rows: ShownRow[],
): void {
  const stack = cellFontStack(null, view.workbookFontFamily);
  ctx.font = fontString(stack, GUTTER_FONT_PX, false);
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  for (const { sheetRow, display } of rows) {
    const y = oy + view.rows.offsets[display];
    const h = spanAt(view.rows.offsets, display);
    const lit = litRow(view.selection, sheetRow);
    if (lit) {
      fill(ctx, 0, y, ROW_HEADER_W - 1, h, GUTTER_TINT);
      ctx.font = fontString(stack, GUTTER_FONT_PX, true);
    } else {
      ctx.font = fontString(stack, GUTTER_FONT_PX, false);
    }
    ctx.fillStyle = lit ? ACCENT_INK : view.theme.mutedForeground;
    ctx.fillText(String(sheetRow + 1), ROW_HEADER_W / 2, y + h / 2);
  }
  ctx.font = fontString(stack, GUTTER_FONT_PX, false);

  for (let c = view.colStart; c < view.colEnd; c += 1) {
    const x = ox + view.cols.offsets[c];
    const w = spanAt(view.cols.offsets, c);
    const lit = litColumn(view.selection, c);
    if (lit) {
      fill(ctx, x, 0, w, COL_HEADER_H - 1, GUTTER_TINT);
      ctx.font = fontString(stack, GUTTER_FONT_PX, true);
    } else {
      ctx.font = fontString(stack, GUTTER_FONT_PX, false);
    }
    ctx.fillStyle = lit ? ACCENT_INK : view.theme.mutedForeground;
    ctx.fillText(columnLabel(c), x + w / 2, COL_HEADER_H / 2);
  }
}

function litRow(selection: Selection, row: number): boolean {
  switch (selection.kind) {
    case "row":
    case "cell":
      return selection.row === row;
    case "range":
      return row >= selection.row && row < selection.row + selection.rows;
    default:
      return false;
  }
}

function litColumn(selection: Selection, col: number): boolean {
  switch (selection.kind) {
    case "column":
    case "cell":
      return selection.col === col;
    case "range":
      return col >= selection.col && col < selection.col + selection.cols;
    default:
      return false;
  }
}
