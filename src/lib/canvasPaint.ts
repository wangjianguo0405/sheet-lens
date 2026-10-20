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
 * would make the cell layer depend on the selection and defeat the split, so the
 * gridlines are drawn as usual and the ring is left to land where it lands. It
 * lands on the border, because `ring` and `paintEdges` share one convention —
 * which is the case these files are full of. Where a sheet shows gridlines
 * instead, the rule under the ring's right and bottom sides stays visible one
 * pixel inside it.
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
  CellPattern,
  CellStyle,
  DataBarCell,
  DataValidation,
  IconCell,
  MergeRect,
  RichRun,
} from "@/lib/excel";
import { ICON_INSET, ICON_SIZE, paintConditionalIcon } from "@/lib/condIcons";
import type { CanvasTheme } from "@/lib/canvasTheme";
import { columnLabel } from "@/lib/column";
import {
  ACCENT_INK,
  CELL_PAD_LEFT,
  CELL_PAD_RIGHT,
  COL_HEADER_H,
  decorationFor,
  FILTERED_INK,
  GUTTER_TINT,
  OUTSIDE_PRINT,
  RING_W,
  ROW_HEADER_W,
  SELECTION_WASH,
  indentPadding,
  horizontalOf,
  resolvedEdges,
  cellKey,
  spanAt,
  type Decoration,
  type Metrics,
} from "@/lib/gridGeometry";
import {
  activeArea,
  areasOf,
  selectionBandCovers,
  selectionBandCrosses,
  selectionCovers,
  type Selection,
} from "@/lib/selection";
import {
  cellFontStack,
  flatText,
  fontString,
  layoutLines,
  lineMetrics,
  runSpans,
  textWidth,
  type RunSpan,
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
  /**
   * The rows and columns the sheet hides. Excel draws none of them and skips
   * the number, so the gap is the only sign it gives; a row that is hidden
   * takes no room at all, which is what puts the next one where Excel has it.
   */
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
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
  /**
   * The review fill a cell carries, or null when it has none. It is the app's
   * own mark rather than the file's, so it wins over the cell's own fill — a
   * mark is only there because the reader put it there.
   */
  markAt: (row: number, col: number) => string | null;
  /**
   * The data-validation rule governing a cell, if any. Only the active cell's
   * is drawn — Excel's in-cell dropdown appears on the cell the cursor is on.
   */
  validationAt: (row: number, col: number) => DataValidation | null;

  // --- sheet facts ---
  workbookFontFamily: string | null;
  workbookFontSizePx: number;
  showGrid: boolean;
  selection: Selection;
  /**
   * What Ctrl+C took on this sheet, or an empty list when the clipboard holds
   * nothing from it. Drawn as a dashed rule, which is the only sign a copy
   * happened at all: Excel's copy shows nothing else. A Ctrl+clicked selection
   * is copied as several blocks, and Excel's marquee rings each of them.
   */
  copied: MergeRect[];
  /** How far the marquee's dashes have crawled, in pixels. */
  ants: number;
  merges: MergeRect[];
  mergeView: {
    shown: MergeRect[];
    covered: Set<number>;
    /** Each covered cell back to the region holding it, for its fill. */
    at: Map<number, MergeRect>;
  };
  centerRuns: {
    shown: (MergeRect & { text: string | null })[];
    covered: Set<number>;
    /** Each covered cell back to the run holding it, for its dropped edges. */
    at: Map<number, MergeRect>;
  };
  autoFilter: AutoFilter | null;
  /**
   * Whether a filter is hiding rows on this sheet. It decides one thing here:
   * the row numbers of the filtered range are drawn in Excel's blue.
   */
  filtering: boolean;
  printArea: MergeRect | null;
  /** Whether a cell lies outside the print area, and so takes the grey shade. */
  outsidePrint: (row: number, col: number) => boolean;
  /**
   * Cells that draw part of their text in runs of their own, keyed by cell —
   * the text alone cannot identify them, since two cells may spell the same
   * text with different runs. Empty for most files.
   */
  cellRich: Map<number, RichRun[]>;
  /** Data bars by cell, for the bars Excel draws over the fill. Sparse. */
  dataBars: Map<number, DataBarCell>;
  /** Conditional icon marks by cell. Sparse. */
  icons: Map<number, IconCell>;
  /** Cells carrying a note, keyed the same way. Sparse. */
  notes: Set<number>;
  theme: CanvasTheme;
}

/** The room a filter header reserves for its button, so the text keeps clear. */
const FILTER_HEADER_PAD_RIGHT = 18;

/** Excel's note mark: a 6×6 pure-red triangle in the cell's top-right corner. */
const NOTE_MARK = "#ff0000";

const GUTTER_FONT_PX = 11;

/** Draw-space origin: sheet coordinates shifted by the scroll and the gutters. */
function originX(view: GridView): number {
  return ROW_HEADER_W - view.scrollLeft;
}
function originY(view: GridView): number {
  return COL_HEADER_H - view.scrollTop;
}

/**
 * How many rectangles this module has filled. Counted rather than inferred,
 * because the surface pass walks the same cells whichever way the timings move:
 * what separates "resolving a cell's style and borders is expensive" from
 * "drawing them is expensive" is how much the pass actually drew, and dividing
 * its total time by one of those counts assumes the other is free.
 *
 * A snapshot delta, so the overlay's own fills do not land in the base pass.
 */
let fillRects = 0;

function fill(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  color: string,
): void {
  if (w <= 0 || h <= 0) return;
  fillRects += 1;
  ctx.fillStyle = color;
  ctx.fillRect(x, y, w, h);
}

/**
 * A `#RRGGBB` colour mixed toward white, which is the light end of a data bar.
 * Measured off Excel's own render: the bar's right edge is about four fifths of
 * the way to white, whichever colour the rule names.
 */
function lighten(hex: string, amount: number): string {
  if (hex.length !== 7) return hex;
  const channel = (at: number) => parseInt(hex.slice(at, at + 2), 16);
  const mix = (value: number) => Math.round(value + (255 - value) * amount);
  return `rgb(${mix(channel(1))}, ${mix(channel(3))}, ${mix(channel(5))})`;
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
 * Excel's 18 hatch patterns as 8×8 masks, measured off a 1:1 `CopyPicture` of
 * every pattern in `fills.xlsx`: `#` is the hatch colour, `.` the ground.
 *
 * The tile belongs to the sheet's coordinate space, not the cell's: two cells
 * carrying the same pattern 19 rows apart line up exactly when their distance
 * is taken modulo 8, so a hatched band runs unbroken across cell edges — as it
 * does in Excel, whose brush origin is the sheet.
 */
const PATTERN_TILES: Record<string, string[]> = {
  lightGray: [
    "#...#...",
    "..#...#.",
    "#...#...",
    "..#...#.",
    "#...#...",
    "..#...#.",
    "#...#...",
    "..#...#.",
  ],
  mediumGray: [
    "#.#.#.#.",
    ".#.#.#.#",
    "#.#.#.#.",
    ".#.#.#.#",
    "#.#.#.#.",
    ".#.#.#.#",
    "#.#.#.#.",
    ".#.#.#.#",
  ],
  darkGray: [
    "###.###.",
    "#.###.##",
    "###.###.",
    "#.###.##",
    "###.###.",
    "#.###.##",
    "###.###.",
    "#.###.##",
  ],
  gray125: [
    "#...#...",
    "........",
    "..#...#.",
    "........",
    "#...#...",
    "........",
    "..#...#.",
    "........",
  ],
  gray0625: [
    "#.......",
    "........",
    "....#...",
    "........",
    "#.......",
    "........",
    "....#...",
    "........",
  ],
  darkHorizontal: [
    "########",
    "########",
    "........",
    "........",
    "########",
    "########",
    "........",
    "........",
  ],
  darkVertical: [
    "##..##..",
    "##..##..",
    "##..##..",
    "##..##..",
    "##..##..",
    "##..##..",
    "##..##..",
    "##..##..",
  ],
  darkDown: [
    "##..##..",
    ".##..##.",
    "..##..##",
    "#..##..#",
    "##..##..",
    ".##..##.",
    "..##..##",
    "#..##..#",
  ],
  darkUp: [
    "..##..##",
    ".##..##.",
    "##..##..",
    "#..##..#",
    "..##..##",
    ".##..##.",
    "##..##..",
    "#..##..#",
  ],
  lightGrid: [
    "########",
    "#...#...",
    "#...#...",
    "#...#...",
    "########",
    "#...#...",
    "#...#...",
    "#...#...",
  ],
  lightTrellis: [
    "#...#...",
    ".#.#.#.#",
    "..#...#.",
    ".#.#.#.#",
    "#...#...",
    ".#.#.#.#",
    "..#...#.",
    ".#.#.#.#",
  ],
  darkTrellis: [
    "########",
    ".##..##.",
    "########",
    "#..##..#",
    "########",
    ".##..##.",
    "########",
    "#..##..#",
  ],
  darkGrid: [
    "#..##..#",
    ".##..##.",
    ".##..##.",
    "#..##..#",
    "#..##..#",
    ".##..##.",
    ".##..##.",
    "#..##..#",
  ],
  lightHorizontal: [
    "########",
    "........",
    "........",
    "........",
    "########",
    "........",
    "........",
    "........",
  ],
  lightVertical: [
    "#...#...",
    "#...#...",
    "#...#...",
    "#...#...",
    "#...#...",
    "#...#...",
    "#...#...",
    "#...#...",
  ],
  lightDown: [
    "#...#...",
    ".#...#..",
    "..#...#.",
    "...#...#",
    "#...#...",
    ".#...#..",
    "..#...#.",
    "...#...#",
  ],
  lightUp: [
    "...#...#",
    "..#...#.",
    ".#...#..",
    "#...#...",
    "...#...#",
    "..#...#.",
    ".#...#..",
    "#...#...",
  ],
};

/** One tile per colour pair; the palette in a workbook is small. */
const patternTiles = new Map<string, CanvasPattern | null>();

function patternFor(pattern: CellPattern): CanvasPattern | null {
  const key = `${pattern.kind}|${pattern.fg}|${pattern.bg}`;
  const known = patternTiles.get(key);
  if (known !== undefined) return known;
  const rows = PATTERN_TILES[pattern.kind];
  if (!rows) return null;
  const tile = document.createElement("canvas");
  tile.width = 8;
  tile.height = 8;
  const ctx = tile.getContext("2d");
  if (!ctx) return null;
  ctx.fillStyle = pattern.bg;
  ctx.fillRect(0, 0, 8, 8);
  ctx.fillStyle = pattern.fg;
  rows.forEach((row, y) => {
    for (let x = 0; x < row.length; x += 1) {
      if (row[x] === "#") ctx.fillRect(x, y, 1, 1);
    }
  });
  const made = ctx.createPattern(tile, "repeat");
  if (patternTiles.size > 64) patternTiles.clear();
  patternTiles.set(key, made);
  return made;
}

/**
 * A patterned cell's background. The tile is anchored at the sheet's origin, so
 * the hatch continues across cell edges exactly as Excel's does.
 */
function paintPattern(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  x: number,
  y: number,
  w: number,
  h: number,
  pattern: CellPattern,
): void {
  const tile = patternFor(pattern);
  if (!tile || w <= 0 || h <= 0) return;
  const ox = originX(view);
  const oy = originY(view);
  ctx.save();
  ctx.translate(ox, oy);
  ctx.fillStyle = tile;
  ctx.fillRect(x - ox, y - oy, w, h);
  ctx.restore();
  fillRects += 1;
}

/**
 * A border edge in whatever style the file asked for. Dashes and dots have to
 * be stroked — a dash pattern cannot be faked with rectangles — and `double`,
 * which canvas has no primitive for, is drawn as the two parallel strokes CSS
 * would.
 */function strokeEdge(
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
 * A cell's diagonal rule, corner to corner inside the cell. Excel keeps one
 * style for the pair and two flags for the direction, so both set means two
 * crossing lines. Off `CopyPicture` of a merged region: the line spans the
 * whole region, which is why the caller passes the region's rectangle rather
 * than one cell's.
 */
function paintDiagonal(
  ctx: CanvasRenderingContext2D,
  edge: BorderEdge | null | undefined,
  up: boolean,
  down: boolean,
  x: number,
  y: number,
  w: number,
  h: number,
  fallback: string,
): void {
  if (!edge || (!up && !down) || w <= 0 || h <= 0) return;
  const color = edge.color ?? fallback;
  const width = Math.max(1, edge.width);
  // The cell's own pixels: a line drawn on x+w would land on the neighbour's
  // edge, where Excel's diagonal never reaches.
  const x1 = x;
  const y1 = y;
  const x2 = x + w - 1;
  const y2 = y + h - 1;

  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.strokeStyle = color;
  // A double diagonal has no CSS twin; two thin strokes to either side of the
  // line is the same construction `strokeEdge` uses.
  const double = edge.line === "double";
  ctx.lineWidth = double ? Math.max(1, width / 3) : width;
  ctx.lineCap = "butt";
  ctx.setLineDash(
    edge.line === "dotted"
      ? [width, width]
      : edge.line === "dashed"
        ? [3 * width, 2 * width]
        : [],
  );
  const offsets = double ? [-width / 3, width / 3] : [0];
  const segment = (ax: number, ay: number, bx: number, by: number, off: number) => {
    const length = Math.hypot(bx - ax, by - ay) || 1;
    const nx = (-(by - ay) / length) * off;
    const ny = ((bx - ax) / length) * off;
    ctx.beginPath();
    ctx.moveTo(ax + nx, ay + ny);
    ctx.lineTo(bx + nx, by + ny);
    ctx.stroke();
  };
  for (const off of offsets) {
    if (down) segment(x1, y1, x2, y2, off);
    if (up) segment(x1, y2, x2, y1, off);
  }
  ctx.restore();
}

/**
 * Excel's note mark: a 6×6 pure-red triangle whose right angle sits in the
 * cell's top-right corner, one pixel down from its top. Measured off a 1:1
 * `CopyPicture` — the mark is part of what Excel copies, and it scales with the
 * zoom like the cell it belongs to.
 */
function paintNoteMark(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
): void {
  // One run per row rather than a path: the 45° hypotenuse then lands on whole
  // pixels, as Excel's does, instead of being spread over an antialiased edge.
  for (let row = 0; row < 6; row += 1) {
    fill(ctx, x + w - 6 + row, y + 1 + row, 6 - row, 1, NOTE_MARK);
  }
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

/**
 * A selection ring, `RING_W` pixels thick, on the four pixels a cell's own
 * border would be drawn on — left and top at the cell's own coordinates, right
 * and bottom one past its last pixel, which is the neighbour's first.
 *
 * That is `paintEdges`'s convention, and it has to be the ring's too: inset so
 * that the right and bottom sides landed on the cell's last pixel, a bordered
 * cell showed the green line *and* the border beside it — those two sides read
 * as twice the thickness of the left and top, where the ring sits exactly on
 * the border and covers it.
 */
function ring(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  color: string,
): void {
  if (w <= 0 || h <= 0) return;
  const t = RING_W;
  fill(ctx, x, y, w + 1, t, color);
  fill(ctx, x, y + h - t + 1, w + 1, t, color);
  fill(ctx, x, y, t, h + 1, color);
  fill(ctx, x + w - t + 1, y, t, h + 1, color);
}

/** Whether two rectangles share a cell. */
function overlapsRect(a: MergeRect, b: MergeRect): boolean {
  return (
    a.row < b.row + b.rows &&
    b.row < a.row + a.rows &&
    a.col < b.col + b.cols &&
    b.col < a.col + a.cols
  );
}

/**
 * The dashed rule around a copied range, crawling the way Excel's does.
 *
 * Dashed rather than solid because the selection's own rule is drawn in the same
 * ink and the two can be on screen at once. A marquee that stood still would
 * read as one more border, and a dasher drawn straight over a solid line reads
 * as a solid one — the range just copied is usually the selection as well, so
 * the gaps are painted out with the casing first. That is also what keeps the
 * ants legible over a filled or shaded cell. The stroke state is the only thing
 * here that has to be undone, so it is saved and restored rather than reset by
 * hand.
 */
function marquee(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  phase: number,
  casing: string,
): void {
  if (w <= 0 || h <= 0) return;
  // The same box `ring` fills, struck as one centred path instead of four
  // rectangles: a stroke spreads half its width either side of the path, so the
  // path is inset half a width from the box's outer edge and the dashes land on
  // exactly the pixels the selection's own rule covers.
  const t = RING_W;
  const px = x + t / 2;
  const py = y + t / 2;
  const pw = w + 1 - t;
  const ph = h + 1 - t;
  ctx.save();
  ctx.lineWidth = t;
  ctx.strokeStyle = casing;
  ctx.strokeRect(px, py, pw, ph);
  ctx.strokeStyle = ACCENT_INK;
  ctx.setLineDash([4, 3]);
  ctx.lineDashOffset = -phase;
  ctx.strokeRect(px, py, pw, ph);
  ctx.restore();
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

/** The stack a cell's text draws in, which its font is built from. */
function cellFontStackFor(view: GridView, style: CellStyle | null): string {
  return cellFontStack(style?.family ?? null, view.workbookFontFamily);
}

/** The cell's font, as the paint passes and the fitting both need it. */
function cellFont(view: GridView, style: CellStyle | null): string {
  return fontString(
    cellFontStackFor(view, style),
    style?.sizePx ?? view.workbookFontSizePx,
    style?.bold ?? false,
    style?.italic ?? false,
  );
}

/** The family the cell's stack starts with, whose metrics the decorations use. */
function cellFontFamily(view: GridView, style: CellStyle | null): string | null {
  return style?.family ?? view.workbookFontFamily;
}

/**
 * The underline and strikethrough a cell — or one run of its text — draws.
 *
 * Canvas has no `text-decoration`, so the geometry comes from the font's own
 * metrics — measured off Excel's own PDF, see `decorationFor`. Both rules span
 * the ink they belong to rather than the cell, which is what lets a run inside a
 * sentence be underlined on its own.
 */
function paintDecoration(
  ctx: CanvasRenderingContext2D,
  metrics: Decoration,
  left: number,
  width: number,
  baseline: number,
  sizePx: number,
  color: string,
  /** The cell's style, or a run's own font — whichever is drawing this piece. */
  source: { underline?: string | null; strike?: boolean } | null,
): void {
  const underline = source?.underline ?? null;
  const strike = source?.strike ?? false;
  if (underline !== "single" && underline !== "double" && !strike) return;
  ctx.fillStyle = color;
  if (strike) {
    const thickness = metrics.strikeThickness * sizePx;
    // The measured offset is the struck line's centre.
    ctx.fillRect(left, baseline - metrics.strike * sizePx - thickness / 2, width, thickness);
  }
  if (underline === "single") {
    ctx.fillRect(
      left,
      baseline + metrics.underline * sizePx,
      width,
      metrics.underlineThickness * sizePx,
    );
  } else if (underline === "double") {
    const thickness = metrics.doubleThickness * sizePx;
    ctx.fillRect(left, baseline + metrics.double[0] * sizePx, width, thickness);
    ctx.fillRect(left, baseline + metrics.double[1] * sizePx, width, thickness);
  }
}

/** The gap a turned cell keeps from the edge it is measured against. */
const TURN_INSET = 2;

/**
 * Text that runs down its cell instead of across it: `textRotation` +90 and
 * −90, and 255's stacked form, where the characters stand one under the next.
 *
 * Where it sits came off Excel's own PDFs rather than out of the standard,
 * which says what the angle means and nothing about placement:
 *
 * - A turned line is centred down the cell's height, and its baseline sits
 *   against the edge its descenders point at — the right for a line read
 *   bottom-to-top, the left for one read top-to-bottom. `center` moves it to
 *   the middle of the cell's width instead.
 * - A stacked column is centred in the cell, and its pitch is the font's own
 *   line box, which is what Excel stacks by: 14.2pt at 10pt Meiryo UI against
 *   16.8pt at 10pt Yu Gothic, so no single constant fits both.
 *
 * The cell clips it, the way it clips a flat cell's text.
 */
function paintTurnedText(
  ctx: CanvasRenderingContext2D,
  view: GridView,
  box: { x: number; y: number; w: number; h: number },
  text: string,
  style: CellStyle | null,
  color: string,
): void {
  const content = flatText(text);
  if (content === "") return;
  const font = cellFont(view, style);
  const metrics = lineMetrics(font, style?.sizePx ?? view.workbookFontSizePx);
  const horizontal = style?.horizontal ?? null;

  ctx.save();
  ctx.beginPath();
  ctx.rect(box.x, box.y, box.w, box.h);
  ctx.clip();
  ctx.font = font;
  ctx.fillStyle = color;
  ctx.textBaseline = "alphabetic";

  if (style?.stacked) {
    const chars = [...content];
    /** One character per line box, the font's own. */
    const pitch = metrics.ascent + metrics.descent;
    const height = chars.length * pitch;
    const top =
      style?.vertical === "top"
        ? box.y + TURN_INSET
        : style?.vertical === "middle"
          ? box.y + (box.h - height) / 2
          : box.y + box.h - height;
    // General is centred here: a stacked column that hugged one edge of its
    // column read as a mistake next to Excel's own.
    ctx.textAlign =
      horizontal === "left" ? "left" : horizontal === "right" ? "right" : "center";
    const x =
      horizontal === "left"
        ? box.x + TURN_INSET
        : horizontal === "right"
          ? box.x + box.w - TURN_INSET
          : box.x + box.w / 2;
    chars.forEach((char, index) => {
      ctx.fillText(char, x, top + metrics.ascent + index * pitch);
    });
    ctx.restore();
    return;
  }

  const angle = style?.rotation ?? 0;
  const ccw = angle > 0;
  const rad = (-angle * Math.PI) / 180;
  /**
   * A second line sits beside the first, on the far side of its baseline — and
   * before it can be laid out, the first has to know how much room the cell
   * gives it along its own direction.
   */
  const lines = layoutLines(font, content, box.h - TURN_INSET * 2, style?.wrap ?? false);
  const pitch = metrics.ascent + metrics.descent;
  const baseline =
    horizontal === "center"
      ? box.x + box.w / 2 + (metrics.ascent - metrics.descent) / 2
      : ccw
        ? box.x + box.w - TURN_INSET - metrics.descent
        : box.x + TURN_INSET + metrics.descent;
  /**
   * Only the quarter turns have a measured anchor. Excel does write other
   * angles — 45 and −45 among them — but no workbook here uses one, so rather
   * than guess at the rule those are laid out as a block centred in the cell,
   * which is where Excel's own 45° came out to within a few points and which
   * cannot leave the text hanging off an edge.
   */
  const quarter = Math.abs(angle) === 90;

  ctx.textAlign = "left";
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    if (line === "") continue;
    const length = textWidth(font, line);
    let x: number;
    let y: number;
    if (quarter) {
      // The origin is where the text starts: the bottom of a centred line read
      // upwards, its top when read downwards.
      x = ccw ? baseline - index * pitch : baseline + index * pitch;
      y =
        style?.vertical === "top"
          ? ccw
            ? box.y + TURN_INSET + length
            : box.y + TURN_INSET
          : style?.vertical === "middle"
            ? ccw
              ? box.y + box.h / 2 + length / 2
              : box.y + box.h / 2 - length / 2
            : ccw
              ? box.y + box.h - TURN_INSET
              : box.y + box.h - TURN_INSET - length;
    } else {
      // The centre of the turned line's own box, put on the cell's centre.
      const cos = Math.cos(rad);
      const sin = Math.sin(rad);
      const across = (metrics.descent - metrics.ascent) / 2;
      x = box.x + box.w / 2 - (length / 2) * cos + across * sin;
      y = box.y + box.h / 2 - (length / 2) * sin - across * cos;
    }
    ctx.save();
    // Canvas turns clockwise and the pixel grid runs downwards, so a positive
    // Excel angle — a line read bottom-to-top — is a negative rotation here.
    ctx.translate(x, y);
    ctx.rotate(rad);
    ctx.fillText(line, 0, 0);
    ctx.restore();
  }
  ctx.restore();
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

  /**
   * A turned cell — `textRotation` 90 or −90, or 255's stacked form — runs
   * down the cell's height rather than across it, so it is laid out and drawn
   * by its own pass. Nearly every cell in every workbook is flat; the ones that
   * are not are the narrow column headings a Japanese document writes its
   * labels in.
   */
  if (style?.stacked || (style?.rotation ?? 0) !== 0) {
    paintTurnedText(ctx, view, box, text, style, textColor(style, kind, view.theme));
    return;
  }

  if (available <= 0) return;

  const wrap = style?.wrap ?? false;
  /**
   * An unwrapped value has no breaks in it as far as Excel is concerned, so it
   * is measured, laid out and drawn without them. The runs lose the same ones,
   * which keeps every offset below pointing at the character it colours.
   */
  const content = wrap ? text : flatText(text);
  const spans =
    runs && runs.length > 0
      ? runSpans(wrap ? runs : runs.map((run) => ({ ...run, text: flatText(run.text) })))
      : null;

  const cellStack = cellFontStackFor(view, style);
  const cellPx = style?.sizePx ?? view.workbookFontSizePx;
  const cellBold = style?.bold ?? false;
  const cellItalic = style?.italic ?? false;
  /**
   * Shrink to fit scales whatever the cell draws, so the factor stands above the
   * cell's own size and every run's at once — a run that keeps its size while the
   * rest shrinks would overflow by exactly the difference.
   */
  let scale = 1;
  /**
   * The size a run is laid out at. Excel stores the base size for a
   * super/subscript run and scales it when drawing, so line heights stay the
   * base font's — this is what wrapping and alignment measure with.
   */
  const layoutPx = (span: RunSpan | null): number => (span?.font?.sizePx ?? cellPx) * scale;
  /**
   * Superscript and subscript draw at two thirds of the size they are laid out
   * at. Measured off `CopyPicture` at 11pt and 20pt: the digit's cap height is
   * 7px where the base is 11px, and 13px where it is 20px.
   */
  const vertOf = (span: RunSpan | null): "superscript" | "subscript" | null =>
    span?.font ? (span.font.vertAlign ?? null) : (style?.vertAlign ?? null);
  const drawPx = (span: RunSpan | null): number =>
    vertOf(span) ? (layoutPx(span) * 2) / 3 : layoutPx(span);
  /**
   * So far the baseline moves for such a run. Same measurement: a superscript's
   * foot sits 7px above the base line at 11pt and 13px at 20 — half the size;
   * a subscript's sits a fifth below.
   */
  const riseOf = (span: RunSpan | null): number => {
    const vert = vertOf(span);
    if (vert === "superscript") return -0.5 * layoutPx(span);
    if (vert === "subscript") return 0.2 * layoutPx(span);
    return 0;
  };
  const piecePx = (span: RunSpan | null): number => drawPx(span);
  /**
   * The font one run draws in: what its `<rPr>` states, and the cell's font for
   * everything it leaves out — including the run that has no `<rPr>` at all, and
   * which therefore takes the cell's font whole.
   */
  const pieceFont = (span: RunSpan | null): string => {
    const run = span?.font ?? null;
    if (!run) return fontString(cellStack, piecePx(null), cellBold, cellItalic);
    return fontString(
      cellFontStack(run.family ?? style?.family ?? null, view.workbookFontFamily),
      piecePx(span),
      run.bold ?? false,
      run.italic ?? false,
    );
  };
  /**
   * The width of a slice of the source, each run measured in its own font. Lines
   * are laid out and placed with this rather than with the cell's font alone: a
   * word in a larger font has to wrap where Excel wraps it and, in a centred or
   * right-aligned cell, start where Excel starts it.
   */
  const widthOf = (from: number, to: number): number => {
    if (!spans) return textWidth(pieceFont(null), content.slice(from, to));
    let total = 0;
    for (const span of spans) {
      const start = Math.max(from, span.start);
      const end = Math.min(to, span.end);
      if (end > start) total += textWidth(pieceFont(span), content.slice(start, end));
    }
    return total;
  };

  // `widthOf` is the measure the layout uses, so the font here is only the one
  // it would have used — the cell's own, before any shrink.
  const lines = layoutLines(cellFont(view, style), content, available, wrap, widthOf);
  /**
   * Shrink to fit: Excel scales the whole font down until the cell's line fits.
   * That is the one line it has, because the standard makes the setting
   * inapplicable to a cell holding several. Measured off Excel's own PDF — 14pt
   * text in an 8-character column lands at 8.52pt, 7.44, 5.64, 4.20, 2.64 as it
   * gets longer — so the factor is a plain ratio of what is needed to what is
   * there, and a second pass covers the rounding.
   */
  if (style?.shrink && lines.length === 1 && lines[0] !== "") {
    for (let pass = 0; pass < 3; pass += 1) {
      const width = widthOf(0, lines[0].length);
      if (width <= available) break;
      scale *= available / width;
    }
  }
  const sizePx = layoutPx(null);
  const font = pieceFont(null);
  // The line is the base font's even when the whole cell is superscript: the
  // stored size is the layout size, and the two-thirds scale is a drawing rule.
  const metrics = lineMetrics(fontString(cellStack, sizePx, cellBold, cellItalic), sizePx);
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
   * Where each line's text sits in the source, and how wide its ink is.
   *
   * The lines come back in order and contiguous, so the next occurrence from
   * here is the right one — a wrapped break is a prefix of what it cut, and the
   * whitespace the layout trimmed off a line's end is simply skipped over. The
   * offsets are needed before anything is drawn, because a line wider than the
   * cell is what decides whether the whole block has to be clipped.
   */
  const lineStarts: number[] = [];
  const lineWidths: number[] = [];
  let cursor = 0;
  for (const source of lines) {
    const start = source === "" ? cursor : Math.max(cursor, content.indexOf(source, cursor));
    cursor = start + source.length;
    lineStarts.push(start);
    lineWidths.push(source === "" ? 0 : widthOf(start, cursor));
  }

  /**
   * Excel clips a value with no room left at the edge of the room it does have —
   * it does not mark the cut with an ellipsis. The clip is only set up when
   * something actually overflows: `clip()` is not cheap and this runs for every
   * cell on screen.
   */
  const tooWide = wrap ? false : lineWidths.some((width) => width > available);
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
  const cellDecoration = decorationFor(cellFontFamily(view, style));

  for (let i = 0; i < lines.length; i += 1) {
    const source = lines[i];
    if (source === "") continue;
    const start = lineStarts[i];
    const baseline = top + i * metrics.lineHeight + halfLeading + metrics.ascent;

    /**
     * The line's own ink — where it starts and how far it reaches — which is
     * what an underline and a strike span, not the cell.
     */
    const width = lineWidths[i];
    const left = align === "right" ? x - width : align === "center" ? x - width / 2 : x;

    if (!spans) {
      const drawBaseline = baseline + riseOf(null);
      ctx.fillStyle = base;
      ctx.fillText(source, x, drawBaseline);
      paintDecoration(
        ctx,
        cellDecoration,
        left,
        width,
        drawBaseline,
        piecePx(null),
        base,
        style,
      );
      continue;
    }

    /**
     * A line with runs is drawn piece by piece, from its own left edge — the
     * alignment positions the line, the runs only colour and set the font of
     * their own part. Each piece draws its own underline and strike, so a run
     * turned plain inside an underlined cell loses the line under itself only.
     */
    ctx.textAlign = "left";
    let at = left;
    for (const span of spans) {
      const from = Math.max(start, span.start);
      const to = Math.min(start + source.length, span.end);
      if (to <= from) continue;
      const piece = content.slice(from, to);
      const runFont = pieceFont(span);
      const pieceWidth = textWidth(runFont, piece);
      const runBaseline = baseline + riseOf(span);
      ctx.font = runFont;
      ctx.fillStyle = span.color ?? base;
      ctx.fillText(piece, at, runBaseline);
      paintDecoration(
        ctx,
        decorationFor(span.font?.family ?? style?.family ?? null),
        at,
        pieceWidth,
        runBaseline,
        piecePx(span),
        span.color ?? base,
        span.font ? span.font : style,
      );
      at += pieceWidth;
    }
    ctx.font = font;
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

  /**
   * The page's top and bottom edges, counted in the rows that are on screen.
   *
   * Not `displayOfRow(area.row + area.rows - 1)`: a page's last row is very often
   * one a filter has hidden — the total under a filtered table is the usual case
   * — and a row that is not on screen has no display index, so that lookup
   * answers -1 and `y1` came out equal to `y0`. A cut-out with no height is a
   * silent no-op, and the grey flood it exists to carve out of stayed put: a
   * sheet filtered anywhere near its last row came up grey corner to corner. The
   * rows on screen are in sheet order, so where a page's edges fall in that order
   * is a bisection.
   */
  const displays = view.rows.offsets.length - 1;
  const upTo = (sheetRow: number): number => {
    let lo = 0;
    let hi = displays;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (view.sheetRowOf(mid) <= sheetRow) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };
  const first = upTo(area.row - 1);
  const last = upTo(area.row + area.rows - 1);
  // Every row of the page hidden leaves no edge to draw a boundary at, and
  // saying so leaves the sheet unshaded rather than all grey.
  if (last <= first) return null;

  /**
   * Clamped, because a page is allowed to run past the last row and column the
   * sheet formats — that is how a document gives a page room to breathe, and
   * `$A$1:$BC$53` on a 52-row sheet is exactly that. Reading `offsets[53]` off
   * the end of the array gave `undefined` and a NaN height, which is the same
   * silent no-op by another route.
   */
  const rowAt = (display: number) =>
    view.rows.offsets[Math.min(view.rows.offsets.length - 1, Math.max(0, display))];
  const colAt = (index: number) =>
    view.cols.offsets[Math.min(view.cols.offsets.length - 1, Math.max(0, index))];
  return {
    x0: ox + colAt(area.col),
    x1: ox + colAt(area.col + area.cols),
    y0: oy + rowAt(first),
    y1: oy + rowAt(last),
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
  pattern: CellPattern | null,
): void {
  if (!page) {
    if (pattern) paintPattern(ctx, view, x, y, w, h, pattern);
    else if (fillColor) fill(ctx, x, y, w, h, fillColor);
    return;
  }
  fill(ctx, x, y, w, h, OUTSIDE_PRINT);
  const ix = Math.max(x, page.x0);
  const iy = Math.max(y, page.y0);
  const ix1 = Math.min(x + w, page.x1);
  const iy1 = Math.min(y + h, page.y1);
  if (ix1 > ix && iy1 > iy) {
    if (pattern) paintPattern(ctx, view, ix, iy, ix1 - ix, iy1 - iy, pattern);
    else fill(ctx, ix, iy, ix1 - ix, iy1 - iy, fillColor ?? view.theme.background);
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
  /** Rectangles the surface pass filled — a border side is one of these. */
  surfaceFills: number;
  /** Milliseconds over the text pass. */
  textMs: number;
  /** Style lookups that did the work, and those the memo answered. */
  styleResolved: number;
  styleMemoHits: number;
}

/**
 * Whether a region covers a cell: a merged region, or the run one cell's text is
 * centred across. Which set is asked is the caller's.
 *
 * Both sets are keyed by the sheet's own width, and this pass goes on drawing
 * past that width into the blank area — where a column would be read as a cell
 * of the row below and answer for it. That is what left a selected row's colour
 * missing from the columns a merge stood above, and what leaves those same
 * columns of the row above undrawn on a sheet whose gridlines are on. Nothing
 * out there belongs to a region, so the question is answered before the key is
 * built.
 */
function inRegion(
  view: GridView,
  covered: Set<number>,
  sheetRow: number,
  col: number,
): boolean {
  if (col >= view.sheetCols) return false;
  return covered.has(cellKey(view.sheetCols, sheetRow, col));
}

/**
 * Whether a cell paints a fill of its own. Gridlines are drawn under fills, so
 * an edge either side of which something is filled is not drawn at all — Excel's
 * own picture of a solid block has no rules inside it or along its edge, and the
 * empty cell beside a filled one loses that edge too.
 *
 * A merged cell answers for its whole region: the anchor's fill is the one
 * Excel paints, whatever the covered cells carry.
 */
function hasFill(view: GridView, row: number, col: number): boolean {
  if (row < 0 || row >= view.sheetRows || col < 0 || col >= view.sheetCols) {
    return false;
  }
  const merge = view.mergeView.at.get(cellKey(view.sheetCols, row, col));
  const at = merge ?? { row, col };
  const style = view.styleAt(at.row, at.col);
  return (style?.fill ?? null) !== null || (style?.pattern ?? null) !== null;
}

/** Whether any cell of a column/row stretch is filled, for a span's own edges. */
function columnFilled(view: GridView, col: number, from: number, to: number): boolean {
  for (let row = from; row <= to; row += 1) if (hasFill(view, row, col)) return true;
  return false;
}

function rowFilled(view: GridView, row: number, from: number, to: number): boolean {
  for (let col = from; col <= to; col += 1) if (hasFill(view, row, col)) return true;
  return false;
}

/**
 * A cell's key in one of the sparse per-cell maps, or null past the sheet's own
 * columns. The blank area is drawn as cells too, and a key built out there
 * collides with a real cell of the next row — the shape `inRegion` already
 * guards the merge and run sets against.
 */
function keyInSheet(view: GridView, row: number, col: number): number | null {
  if (col >= view.sheetCols) return null;
  return cellKey(view.sheetCols, row, col);
}

export function paintBase(ctx: CanvasRenderingContext2D, view: GridView): PaintStats {
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
    inRegion(view, view.mergeView.covered, sheetRow, col);
  const inRun = (sheetRow: number, col: number) =>
    inRegion(view, view.centerRuns.covered, sheetRow, col);

  // Everything the text sits on: fills, print shading, gridlines, borders.
  const surfaceStarted = performance.now();
  const fillsBefore = fillRects;
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
  stats.surfaceFills = fillRects - fillsBefore;

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

  // A merged region is one cell to Excel, and a selection always lands on its
  // anchor — so its mark is read from the anchor too.
  const mark = view.markAt(rect.row, rect.col);
  fillBlock(
    ctx,
    view,
    page,
    x,
    y,
    w,
    h,
    mark ?? style?.fill ?? null,
    mark ? null : (style?.pattern ?? null),
  );

  const lastRow = rect.row + rect.rows - 1;
  const lastCol = rect.col + rect.cols - 1;

  // Gridlines sit under fills: a region carrying one loses its own rules, and
  // so does an edge whose neighbour is filled.
  if (view.showGrid) {
    const filled =
      (style?.fill ?? null) !== null || (style?.pattern ?? null) !== null;
    if (!filled && !columnFilled(view, lastCol + 1, rect.row, lastRow)) {
      line(ctx, x + w - 1, y, 1, h, view.theme.gridline);
    }
    if (!filled && !rowFilled(view, lastRow + 1, rect.col, lastCol)) {
      line(ctx, x, y + h - 1, w, 1, view.theme.gridline);
    }
  }

  // The region's outline, resolved one boundary cell at a time, exactly as the
  // DOM renderer did — sides facing another cell of the same region are not
  // edges and are left out.
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

  // Excel draws the diagonal across the whole region, not just the anchor
  // cell — `misc-diag.xlsx`'s merged A3:C5 is one line corner to corner.
  if (span.kind === "merge") {
    paintDiagonal(
      ctx,
      style?.border?.diagonal ?? null,
      style?.border?.diagonalUp === true,
      style?.border?.diagonalDown === true,
      x,
      y,
      w,
      h,
      view.theme.border,
    );
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
    view.cellRich.get(cellKey(view.sheetCols, span.rect.row, span.rect.col)) ?? null,
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
  const pattern = style?.pattern ?? null;
  const mark = view.markAt(sheetRow, col);
  const key = keyInSheet(view, sheetRow, col);

  if (mark) {
    // A mark is the reader's own, so it takes the cell whole — over the file's
    // fill and over the print-area shade both.
    fill(ctx, x, y, w, h, mark);
  } else if (view.outsidePrint(sheetRow, col)) {
    // Already OUTSIDE_PRINT: paintBase flooded the whole cell area with it and
    // then cut the page rect back out in the sheet's own background, so shading
    // the cell again would only repaint the colour that is underneath it. The
    // one page rect that cannot be measured — its first row hidden — leaves no
    // flood to lean on, and that cell does have to shade itself.
    if (!page) fill(ctx, x, y, w, h, OUTSIDE_PRINT);
  } else if (own) {
    fill(ctx, x, y, w, h, own);
  } else if (pattern) {
    paintPattern(ctx, view, x, y, w, h, pattern);
  }

  // Excel draws a data bar over the cell's fill and under its gridline, as a
  // gradient that runs from the rule's colour to a much lighter one across the
  // bar's own width. Its dark end stays on the axis, so a negative bar — which
  // runs right to left — reads light at its far end.
  const bar = key === null ? undefined : view.dataBars.get(key);
  if (bar) {
    const from = x + w * Math.min(bar.from, bar.to);
    const to = x + w * Math.max(bar.from, bar.to);
    if (to > from) {
      const gradient = ctx.createLinearGradient(from, y, to, y);
      const light = lighten(bar.color, 0.83);
      gradient.addColorStop(0, bar.negative ? light : bar.color);
      gradient.addColorStop(1, bar.negative ? bar.color : light);
      ctx.fillStyle = gradient;
      ctx.fillRect(from, y, to - from, h);
    }
    if (bar.axis !== null && bar.axis !== undefined) {
      // A crisp pixel, the way Excel draws the axis; a fractional x would be
      // spread over two pixels and read grey.
      ctx.fillStyle = bar.axisColor ?? "#000000";
      ctx.fillRect(Math.round(x + w * bar.axis), y, 1, h);
    }
  }

  // An icon-set mark sits at the cell's left, centred vertically. The value is
  // still drawn, moved right past the mark — see paintCellText.
  const icon = key === null ? undefined : view.icons.get(key);
  if (icon) {
    paintConditionalIcon(
      ctx,
      x + CELL_PAD_LEFT,
      y + (h - ICON_SIZE) / 2,
      ICON_SIZE,
      icon.family,
      icon.level,
    );
  }

  // The hairline grid, before the borders so a border wins where both land. It
  // lies under fills as well: an edge with a filled cell on either side is not
  // drawn, which is what turns a filled block into a solid rectangle.
  if (view.showGrid) {
    if (!own && !pattern && !hasFill(view, sheetRow, col + 1)) {
      line(ctx, x + w - 1, y, 1, h, view.theme.gridline);
    }
    if (!own && !pattern && !hasFill(view, sheetRow + 1, col)) {
      line(ctx, x, y + h - 1, w, 1, view.theme.gridline);
    }
  }

  // Excel drops the borders between the cells of a Center-Across run: the run
  // is one box, so only its outer edges are drawn. Without this every cell kept
  // its own left and right rule, and a header centred across a dozen columns
  // came out as a dozen little boxes.
  const run = key === null ? undefined : view.centerRuns.at.get(key);
  paintEdges(
    ctx,
    view,
    resolvedEdges(
      view.styleAt,
      { row: sheetRow, col, rows: 1, cols: 1 },
      { rows: view.sheetRows, cols: view.sheetCols },
      run
        ? { left: col > run.col, right: col < run.col + run.cols - 1 }
        : undefined,
    ),
    x,
    y,
    w,
    h,
  );

  paintDiagonal(
    ctx,
    style?.border?.diagonal ?? null,
    style?.border?.diagonalUp === true,
    style?.border?.diagonalDown === true,
    x,
    y,
    w,
    h,
    view.theme.border,
  );

  // A note's mark goes over the fill and the borders, in the corner Excel puts
  // it in. A merged region's covered cells are not painted, so a note on one
  // of them — Excel does not show a mark there either — draws nothing.
  if (key !== null && view.notes.has(key)) paintNoteMark(ctx, x, y, w);
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
  const key = keyInSheet(view, sheetRow, col);

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
  // An icon-set mark takes the cell's left edge, so the value starts past it.
  if (key !== null && view.icons.has(key)) {
    padLeft += ICON_INSET;
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
  //
  // The sheet's own last column is not a barrier either. Excel spills past the
  // cells the file mentions into ones it has never heard of, which is how a
  // label sitting in the rightmost column of a narrow sheet — a rules page
  // three columns wide — is read at all; stopping at the data cut every such
  // label down to the two characters its own column holds. What stops it in
  // the end is the value running out, or the scroller's own edge — which for
  // such a label is off screen by then.
  const box = {
    x,
    y,
    w: valueWidth(view, sheetRow, col, w, padLeft, padRight),
    h,
  };

  paintText(
    ctx,
    view,
    box,
    text,
    style,
    cell?.kind,
    padLeft,
    padRight,
    key === null ? null : view.cellRich.get(key) ?? null,
  );
  return true;
}

/**
 * The width a value is drawn in: its own cell's, or — when it is left-aligned,
 * unwrapped and wider than that — as much of the room beside it as it needs,
 * never past the first thing that stops a spill.
 *
 * The field a double-click opens over a cell is this same box, which is why it
 * is here rather than inline: a column two pixels wide holds nothing readable,
 * and its value is read only because the field takes the room the value is
 * drawn in. A field the width of the column alone showed nothing at all.
 *
 * The insets are the caller's, since a filter header's are not a cell's.
 */
export function valueWidth(
  view: GridView,
  sheetRow: number,
  col: number,
  w: number,
  padLeft: number,
  padRight: number,
): number {
  const cell = view.cellAt(sheetRow, col);
  const text = cell?.text ?? "";
  const style = view.styleAt(sheetRow, col);
  // A cell that shrinks to fit never reaches past its own edge.
  if (text === "" || style?.wrap || style?.shrink) return w;
  if (horizontalOf(style, cell?.kind) !== "left") return w;
  // Only the reach of what is drawn: a break in an unwrapped value is not a
  // space, so it does not widen the room the value spills into.
  const needed = textWidth(cellFont(view, style), flatText(text)) + padLeft + padRight;
  if (needed <= w) return w;

  // Only as far as the value reaches: Excel runs the *drawing* box on to the
  // end of the room, and for a field this same number would be a white bar
  // across a row that is otherwise empty.
  let reached = w;
  for (let next = col + 1; reached < needed && next < view.colEnd; next += 1) {
    if (inRegion(view, view.mergeView.covered, sheetRow, next)) break;
    if ((view.cellAt(sheetRow, next)?.text ?? "") !== "") break;
    reached += spanAt(view.cols.offsets, next);
  }
  return Math.min(reached, needed);
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

  // Excel marks one cell with an outline and anything larger with a fill — and
  // inside a fill it leaves the active cell clear and outlines that one, which
  // is what makes the corner the arrow keys will leave from visible at all. That
  // cell is the area's anchor, so it stays where the block was pinned while the
  // block is grown away from it; a Ctrl+clicked selection has one such cell too,
  // in whichever of its areas is the active one. A run of whole rows or columns
  // leaves no cell clear at all.
  const onlyCell = view.selection.kind === "cell";
  const shown = activeArea(view.selection);
  const active =
    shown.kind === "range"
      ? shown.anchor
      : shown.kind === "cell"
        ? { row: shown.row, col: shown.col }
        : null;
  for (const { sheetRow, display } of rows) {
    const y = oy + view.rows.offsets[display];
    const h = spanAt(view.rows.offsets, display);
    for (let c = view.colStart; c < view.colEnd; c += 1) {
      if (!selectionCovers(view.selection, sheetRow, c)) continue;
      // A merged region's cells are not cells and are handled below — except
      // under a band of whole rows or columns, which Excel draws straight over
      // what it crosses. A Center-Across run's cells are cells either way, so a
      // selected row lights every one of them.
      const band = selectionBandCovers(view.selection, sheetRow, c);
      if (!band && inRegion(view, view.mergeView.covered, sheetRow, c)) continue;
      const x = ox + view.cols.offsets[c];
      const w = spanAt(view.cols.offsets, c);
      if (active && active.row === sheetRow && active.col === c) {
        ring(ctx, x, y, w, h, ACCENT_INK);
      } else {
        fill(ctx, x, y, w, h, SELECTION_WASH);
      }
    }
  }

  // A merged region is one cell, and a selection of cells keeps it at its
  // anchor: it lights whole when that top-left is held, and not when the
  // selection merely crosses the region. A band of whole rows or columns is not
  // that — it was drawn over its own part of the region above, and the rest of
  // the region is none of its business. A Center-Across run is not merged and
  // was washed cell by cell above as well.
  for (const rect of view.mergeView.shown) {
    if (selectionBandCrosses(view.selection, rect)) continue;
    if (!selectionCovers(view.selection, rect.row, rect.col)) continue;
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
    if (onlyCell) ring(ctx, x, y, w, h, ACCENT_INK);
    else fill(ctx, x, y, w, h, SELECTION_WASH);
  }

  // Excel draws one rule around the whole block rather than a box per cell —
  // and one around each block of a Ctrl+clicked selection, which is what says
  // where one area of it ends and the next begins.
  for (const area of areasOf(view.selection)) {
    if (area.kind !== "range") continue;
    const top = view.displayOfRow(area.row);
    const bottom = view.displayOfRow(area.row + area.rows - 1);
    const right = Math.min(view.sheetCols - 1, area.col + area.cols - 1);
    if (top >= 0 && bottom >= 0) {
      const x = ox + view.cols.offsets[area.col];
      const y = oy + view.rows.offsets[top];
      ring(
        ctx,
        x,
        y,
        view.cols.offsets[right + 1] - view.cols.offsets[area.col],
        view.rows.offsets[bottom + 1] - view.rows.offsets[top],
        ACCENT_INK,
      );
    }
  }

  // Excel's in-cell dropdown, on the active cell alone: a small button at the
  // right edge, which a list validation offers its choices through. It is not
  // drawn for a file that turned the dropdown off, and the attribute that says
  // so is inverted in the format.
  if (active) {
    const rule = view.validationAt(active.row, active.col);
    if (rule?.kind === "list" && !rule.hideDropdown) {
      const display = view.displayOfRow(active.row);
      if (display >= 0) {
        const x = ox + view.cols.offsets[active.col];
        const y = oy + view.rows.offsets[display];
        const w = spanAt(view.cols.offsets, active.col);
        const h = spanAt(view.rows.offsets, display);
        const size = Math.min(16, Math.max(9, h - 2));
        const bx = x + w - size - 1;
        const by = y + (h - size) / 2;
        fill(ctx, bx, by, size, size, view.theme.background);
        ring(ctx, bx, by, size, size, view.theme.border);
        const cx = bx + size / 2;
        const cy = by + size / 2;
        ctx.fillStyle = view.theme.mutedForeground;
        ctx.beginPath();
        ctx.moveTo(cx - 3.5, cy - 1.5);
        ctx.lineTo(cx + 3.5, cy - 1.5);
        ctx.lineTo(cx, cy + 2.5);
        ctx.closePath();
        ctx.fill();
      }
    }
  }

  // What Ctrl+C took, if it was taken here. Drawn over the selection's own
  // rule, so a copied block that is also selected keeps both.
  for (const taken of view.copied) {
    // A merged cell is one cell to the grid and a block to Excel, whose marquee
    // is drawn around the whole of it — so the copied rectangle is grown over any
    // merged region it reaches into before it is outlined. Merged regions never
    // overlap one another, so one pass is enough. The range ring above needs none
    // of this: it is drawn from the region itself.
    let rect = taken;
    for (const merge of view.merges) {
      if (!overlapsRect(rect, merge)) continue;
      const row = Math.min(rect.row, merge.row);
      const col = Math.min(rect.col, merge.col);
      rect = {
        row,
        col,
        rows: Math.max(rect.row + rect.rows, merge.row + merge.rows) - row,
        cols: Math.max(rect.col + rect.cols, merge.col + merge.cols) - col,
      };
    }
    const top = view.displayOfRow(rect.row);
    const bottom = view.displayOfRow(rect.row + rect.rows - 1);
    if (top >= 0 && bottom >= 0) {
      const right = Math.min(view.sheetCols - 1, rect.col + rect.cols - 1);
      marquee(
        ctx,
        ox + view.cols.offsets[rect.col],
        oy + view.rows.offsets[top],
        view.cols.offsets[right + 1] - view.cols.offsets[rect.col],
        view.rows.offsets[bottom + 1] - view.rows.offsets[top],
        view.ants,
        view.theme.background,
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

  // Each strip is clipped to itself, because both passes walk cells that begin
  // off screen: the first visible row is above the viewport once it is scrolled
  // at all, and its number is drawn at that row's centre — which is inside the
  // corner where the two headers meet. A column walked from the left does the
  // same to the same corner. The clip is what keeps each label on its own strip.
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, COL_HEADER_H, ROW_HEADER_W, view.height - COL_HEADER_H);
  ctx.clip();

  for (const { sheetRow, display } of rows) {
    // A hidden row carries no number in Excel, and has no room for one here.
    if (view.hiddenRows.has(sheetRow)) continue;
    const y = oy + view.rows.offsets[display];
    const h = spanAt(view.rows.offsets, display);
    const lit = litRow(view.selection, sheetRow);
    if (lit) {
      fill(ctx, 0, y, ROW_HEADER_W - 1, h, GUTTER_TINT);
      ctx.font = fontString(stack, GUTTER_FONT_PX, true);
    } else {
      ctx.font = fontString(stack, GUTTER_FONT_PX, false);
    }
    // A filter's own range keeps Excel's blue on the numbers of the rows it
    // shows — not on the row that carries the filter buttons, and not past its
    // last row. It is the sheet's only mark that rows are missing from it.
    const filtering =
      view.filtering &&
      view.autoFilter !== null &&
      sheetRow > view.autoFilter.row &&
      sheetRow < view.autoFilter.row + view.autoFilter.rows;
    ctx.fillStyle = lit
      ? ACCENT_INK
      : filtering
        ? FILTERED_INK
        : view.theme.mutedForeground;
    ctx.fillText(String(sheetRow + 1), ROW_HEADER_W / 2, y + h / 2);
  }
  ctx.restore();

  ctx.save();
  ctx.beginPath();
  ctx.rect(ROW_HEADER_W, 0, view.width - ROW_HEADER_W, COL_HEADER_H);
  ctx.clip();

  for (let c = view.colStart; c < view.colEnd; c += 1) {
    // The letter goes with the column, as it does in Excel's own header.
    if (view.hiddenCols.has(c)) continue;
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
  ctx.restore();
}

function litRow(selection: Selection, row: number): boolean {
  return areasOf(selection).some((area) => {
    switch (area.kind) {
      case "row":
      case "range":
        return row >= area.row && row < area.row + area.rows;
      case "cell":
        return area.row === row;
      // A run of whole columns is every row's, which says nothing about any
      // one of them.
      case "column":
        return false;
    }
  });
}

function litColumn(selection: Selection, col: number): boolean {
  return areasOf(selection).some((area) => {
    switch (area.kind) {
      case "column":
      case "range":
        return col >= area.col && col < area.col + area.cols;
      case "cell":
        return area.col === col;
      case "row":
        return false;
    }
  });
}
