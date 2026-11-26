/**
 * Where Excel breaks a sheet into pages, worked out the way Excel works it
 * out: the paper and its margins give the printable area, the print scale
 * divides it back into sheet pixels, and columns and rows are taken whole
 * until the next one would not fit.
 *
 * Nothing here is in the file. Excel stores only the breaks somebody placed by
 * hand; the automatic ones are recomputed from the page setup every time the
 * sheet is shown, which is why this has to exist on the reading side at all.
 */

import type { MergeRect, PageSetup } from "@/lib/excel";
import type { Metrics } from "@/lib/gridGeometry";

const PT_PER_IN = 72;
const PX_PER_IN = 96;

/**
 * Paper sizes in points, portrait, from ECMA-376's `ST_PaperSize`. Only the
 * sizes a document is likely to use are listed; anything else falls back to
 * Letter, which is the standard's own default.
 */
const PAPER_PT: Record<number, [number, number]> = {
  1: [612, 792], // Letter
  3: [792, 1224], // Tabloid
  4: [1224, 792], // Ledger
  5: [612, 1008], // Legal
  6: [396, 612], // Statement
  7: [522, 756], // Executive
  8: [842, 1191], // A3
  9: [595, 842], // A4
  10: [595, 842], // A4 Small
  11: [420, 595], // A5
  12: [729, 1032], // B4 (JIS)
  13: [516, 729], // B5 (JIS)
  14: [612, 936], // Folio
  27: [842, 1191], // A3, written by some writers under a second id
};

/** A page boundary in sheet pixels, and whether somebody placed it. */
export interface PageLine {
  at: number;
  manual: boolean;
}

/** One page, in sheet pixels, with the number Excel prints on it. */
export interface PageRect {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  number: number;
}

export interface Pagination {
  cols: PageLine[];
  rows: PageLine[];
  pages: PageRect[];
}

/** One page's printable area, in sheet pixels at the given scale. */
function printable(setup: PageSetup, scale: number, grid: number) {
  const [w, h] = PAPER_PT[setup.paperSize] ?? PAPER_PT[1];
  const width = ((setup.landscape ? h : w) * PX_PER_IN) / PT_PER_IN;
  const height = ((setup.landscape ? w : h) * PX_PER_IN) / PT_PER_IN;
  return {
    width: (width - (setup.margins.left + setup.margins.right) * PX_PER_IN) / scale / grid,
    // Excel's body is a shade taller than the margins leave it: six pixels,
    // measured as the only amount that reproduces every break on the design
    // documents this was checked against. Nothing in the file explains it, and
    // it does not scale with the paper (Letter, Legal and A3 each want a
    // different number), so it stays a measured constant rather than a rule.
    height:
      (height - (setup.margins.top + setup.margins.bottom) * PX_PER_IN + PAGE_SLACK) / scale / grid,
  };
}

/**
 * How much smaller than its stored size Excel prints the grid, by the workbook's
 * own font.
 *
 * A workbook set in ＭＳ Ｐゴシック comes out of Excel with every page holding
 * about a tenth more rows and columns than the file's sizes allow, while one set
 * in 游ゴシック holds exactly what they allow. It is the grid that is laid out
 * smaller and not the page that is bigger: an exported PDF draws a 12pt row
 * 10.8pt tall at 100% and the text at its full size, and the column pitch is
 * short by the same share. Measured across the corpus — 24 workbooks in
 * ＭＳ Ｐゴシック need 1.09× the page (p25 1.095, median 1.096, p75 1.105) and
 * 22 in 游ゴシック need 1.00× (p25 1.008, median 1.014) — with nothing between
 * the two groups, and the same split whether or not the pages repeat title rows.
 *
 * Nothing in the file says why, so this is the measurement and not a rule. 0.9
 * is the row pitch a PDF exported from Excel shows — 12pt drawn 10.8pt at 100%
 * — and it is also the constant that reproduces the most first breaks across
 * the corpus. Sheets in any other font keep 1.
 */
function gridPrintFactor(fontFamily: string | null): number {
  const family = (fontFamily ?? "").replace(/\s+/g, "");
  return family === "ＭＳＰゴシック" ? 0.9 : 1;
}

/** How much taller than paper-minus-margins Excel's page body is. */
const PAGE_SLACK = 6;

/** Sub-pixel slack, so a row that lands exactly on the edge is not pushed over. */
const EPSILON = 0.01;

function offsetAt(offsets: Float64Array, index: number): number {
  return offsets[Math.min(offsets.length - 1, Math.max(0, index))];
}

/** Pixels of `$A:$C`-style bounds the print titles take, or 0. */
function titleSize(
  titles: [number, number] | null | undefined,
  width: (index: number) => number,
): number {
  if (!titles) return 0;
  let total = 0;
  for (let index = titles[0]; index <= titles[1]; index += 1) total += width(index);
  return total;
}

/**
 * The manual breaks an axis is actually laid out with.
 *
 * Excel ignores the breaks somebody placed on an axis Fit to constrains: a
 * sheet fitted to one page wide prints every column on that page even though
 * the file still stores a vertical break, and the same holds for a fitted
 * height. The breaks stay in the file — `VPageBreaks` still reports them — but
 * neither the print nor Page Break Preview uses them. Measured on a document
 * fitted one page wide with a stored break after column 37 (its preview spans
 * the whole print area and its PDF is one page wide) and on hand-made books
 * fitted 1 × 1 (row break dropped) and 2 wide (column break dropped); a height
 * left automatic keeps its row break.
 */
function keptBreaks(breaks: number[], fitted: boolean): number[] {
  return fitted ? [] : breaks;
}

/**
 * Columns break whole, with no slack: a column that would run past the right
 * margin moves to the next page. Measured against Excel at 100% and 80%, and
 * on A4 landscape, where the arithmetic lands at 12.5 columns and Excel keeps
 * 12 — a floor, not a rounding.
 */
function cutColumns(setup: PageSetup, area: MergeRect, cols: Metrics, content: number): PageLine[] {
  const lines: PageLine[] = [];
  const last = area.col + area.cols - 1;
  const manual = new Set(keptBreaks(setup.colBreaks, setup.fitToPage && setup.fitToWidth > 0));
  const title = titleSize(
    setup.printTitles?.cols,
    (index) => offsetAt(cols.offsets, index + 1) - offsetAt(cols.offsets, index),
  );
  // The first page holds the title columns as part of its own content; every
  // page after that has them repeated at the left, so they come back out of
  // its room for the columns that continue.
  let used = 0;

  const add = (at: number, placed: boolean) => {
    const existing = lines.find((line) => line.at === at);
    if (existing) existing.manual = existing.manual || placed;
    else lines.push({ at, manual: placed });
  };

  for (let col = area.col; col <= last; col += 1) {
    const width = offsetAt(cols.offsets, col + 1) - offsetAt(cols.offsets, col);
    // A column wider than the page goes on one of its own rather than looping.
    if (used > 0 && used + width > content + EPSILON) {
      add(offsetAt(cols.offsets, col), false);
      used = title;
    }
    used += width;
    if (manual.has(col) && col < last) {
      add(offsetAt(cols.offsets, col + 1), true);
      used = title;
    }
  }
  return lines;
}

/**
 * Rows break with a quarter of a row of slack: the row that runs a little past
 * the bottom margin still goes on this page rather than the next.
 *
 * That is what Excel does, read off the first break at twelve print scales,
 * seven margin settings and five paper sizes, plus every break of two real
 * design documents. Half a row over-fits the odd boundary — on `画面項目定義`
 * it pulled a page one row short and pushed the row count off for every page
 * after it. Columns get no slack at all; the two axes really do disagree.
 */
function cutRows(
  setup: PageSetup,
  area: MergeRect,
  rows: Metrics,
  displayOfRow: (sheetRow: number) => number,
  content: number,
): PageLine[] {
  const lines: PageLine[] = [];
  const last = area.row + area.rows - 1;
  const manual = new Set(keptBreaks(setup.rowBreaks, setup.fitToPage && setup.fitToHeight > 0));
  const displays = rows.offsets.length - 1;
  const title = (() => {
    if (!setup.printTitles?.rows) return 0;
    let total = 0;
    for (let row = setup.printTitles.rows[0]; row <= setup.printTitles.rows[1]; row += 1) {
      const display = displayOfRow(row);
      if (display < 0) continue;
      total += offsetAt(rows.offsets, display + 1) - offsetAt(rows.offsets, display);
    }
    return total;
  })();
  // As with the columns: the first page carries the titles once, every later
  // page repeats them and has that much less room for what follows.
  let used = 0;

  const add = (at: number, placed: boolean) => {
    const existing = lines.find((line) => line.at === at);
    if (existing) existing.manual = existing.manual || placed;
    else lines.push({ at, manual: placed });
  };

  for (let row = area.row; row <= last; row += 1) {
    const display = displayOfRow(row);
    // A hidden row — or one a filter took out — takes no page space, and its
    // absence must not put a break where the next visible row starts.
    if (display < 0 || display >= displays) continue;
    const height = offsetAt(rows.offsets, display + 1) - offsetAt(rows.offsets, display);
    if (used > 0 && used + height > content + height / 4 + EPSILON) {
      add(offsetAt(rows.offsets, display), false);
      used = title;
    }
    used += height;
    if (manual.has(row) && row < last) {
      add(offsetAt(rows.offsets, display + 1), true);
      used = title;
    }
  }
  return lines;
}

/** How many pages the area takes at a scale, for the fit-to-page search. */
function pageCount(
  scale: number,
  setup: PageSetup,
  area: MergeRect,
  cols: Metrics,
  rows: Metrics,
  displayOfRow: (row: number) => number,
  grid: number,
) {
  const page = printable(setup, scale, grid);
  return {
    wide: cutColumns(setup, area, cols, page.width).length + 1,
    tall: cutRows(setup, area, rows, displayOfRow, page.height).length + 1,
  };
}

/** Columns the print area spans at 100%, in sheet pixels. */
function contentWidth(area: MergeRect, cols: Metrics): number {
  return offsetAt(cols.offsets, area.col + area.cols) - offsetAt(cols.offsets, area.col);
}

/**
 * The scale a fixed-percentage sheet really prints at.
 *
 * Excel will not print wider than the paper: when the print area at the
 * sheet's own scale is wider than the page, it shrinks the whole page to the
 * largest whole percentage that fits, and the row breaks move with it. Two
 * 54-column sheets measured on the design documents (93% and 94% of A4
 * landscape, 10 mm and 19 mm margins) print at 92% of their stated scale for
 * exactly this reason; their page breaks only line up once that is applied.
 * Sheets whose content already fits are left alone.
 */
function widthFit(
  scale: number,
  setup: PageSetup,
  area: MergeRect,
  cols: Metrics,
  grid: number,
): number {
  const width = contentWidth(area, cols);
  if (width <= 0) return scale;
  const pageWidth = printable(setup, 1, grid).width;
  if (width * scale <= pageWidth + EPSILON) return scale;
  const fit = Math.floor((pageWidth / width) * 100) / 100;
  return scale * Math.max(0.1, fit);
}

/**
 * The scale the sheet prints at.
 *
 * A fixed percentage is the file's own `scale`. Fit-to-page asks for a page
 * count instead, and Excel answers with the largest whole percentage that
 * meets it — 10% to 400% — which is what the search below does.
 */
function resolveScale(
  setup: PageSetup,
  area: MergeRect,
  cols: Metrics,
  rows: Metrics,
  displayOfRow: (row: number) => number,
  grid: number,
): number {
  const fixed = Math.min(4, Math.max(0.1, (setup.scale || 100) / 100));
  // Excel works in whole percents: the width fit is taken at the largest
  // integer percentage that fits, and the scale it leaves behind is whole
  // too. Keeping the fraction (93% x 93% = 86.49) lands a row short on the
  // design documents' fit sheets; 86 does not. Measured over the 04 営業支援
  // corpus: 573 of 1173 breaks follow this way, 549 without the floor.
  if (!setup.fitToPage) return Math.floor(widthFit(fixed, setup, area, cols, grid) * 100) / 100;
  // Both bounds automatic leaves the scale alone; Excel prints at 100%.
  if (setup.fitToWidth === 0 && setup.fitToHeight === 0) return 1;

  for (let percent = 400; percent >= 10; percent -= 1) {
    const scale = percent / 100;
    const { wide, tall } = pageCount(scale, setup, area, cols, rows, displayOfRow, grid);
    const fitsWide = setup.fitToWidth === 0 || wide <= setup.fitToWidth;
    const fitsTall = setup.fitToHeight === 0 || tall <= setup.fitToHeight;
    if (fitsWide && fitsTall) return scale;
  }
  return 0.1;
}

/**
 * The page grid over a print area: where each page ends, and each page's
 * rectangle and number.
 *
 * Page numbers run down a column of pages before moving right, unless the file
 * asked for `overThenDown`.
 *
 * `fontFamily` is the workbook's own font, which is what decides how much of
 * the sheet Excel fits on a page — see `gridPrintFactor`.
 */
export function paginate(
  setup: PageSetup,
  area: MergeRect,
  cols: Metrics,
  rows: Metrics,
  displayOfRow: (sheetRow: number) => number,
  fontFamily: string | null,
): Pagination {
  const grid = gridPrintFactor(fontFamily);
  const scale = resolveScale(setup, area, cols, rows, displayOfRow, grid);
  const page = printable(setup, scale, grid);
  const colLines = cutColumns(setup, area, cols, page.width);
  const rowLines = cutRows(setup, area, rows, displayOfRow, page.height);

  const xs = [
    offsetAt(cols.offsets, area.col),
    ...colLines.map((line) => line.at),
    offsetAt(cols.offsets, area.col + area.cols),
  ];
  // The area's first and last rows that are actually on screen: a hidden or
  // filtered row at either edge leaves no page edge to draw at.
  let firstDisplay = -1;
  let lastDisplay = -1;
  for (let row = area.row; row < area.row + area.rows; row += 1) {
    const display = displayOfRow(row);
    if (display < 0 || display >= rows.offsets.length - 1) continue;
    if (firstDisplay < 0) firstDisplay = display;
    lastDisplay = display;
  }
  const y0 = firstDisplay < 0 ? 0 : rows.offsets[firstDisplay];
  const y1 = lastDisplay < 0 ? 0 : rows.offsets[lastDisplay + 1];
  const ys = [y0, ...rowLines.map((line) => line.at), y1];

  const pages: PageRect[] = [];
  for (let i = 0; i < xs.length - 1; i += 1) {
    for (let j = 0; j < ys.length - 1; j += 1) {
      if (xs[i + 1] <= xs[i] || ys[j + 1] <= ys[j]) continue;
      const number = setup.pageOrderOverThenDown
        ? j * (xs.length - 1) + i + 1
        : i * (ys.length - 1) + j + 1;
      pages.push({
        x0: xs[i],
        y0: ys[j],
        x1: xs[i + 1],
        y1: ys[j + 1],
        number,
      });
    }
  }
  return { cols: colLines, rows: rowLines, pages };
}
