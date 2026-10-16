import { invoke } from "@tauri-apps/api/core";

export type CellKind = "empty" | "text" | "number" | "bool" | "date" | "error";

export interface Cell {
  kind: CellKind;
  text: string;
}

export interface SheetInfo {
  name: string;
  rows: number;
  cols: number;
  /** The tab's own colour, or null where Excel shows none. */
  tabColor?: string | null;
  /**
   * Whether the sheet defines a print area of its own. A workbook where none
   * does opens in the ordinary view rather than under the page shading, which
   * would otherwise be drawn around a page nobody asked for.
   */
  printArea: boolean;
}

export interface WorkbookInfo {
  path: string;
  fileName: string;
  sheets: SheetInfo[];
}

export interface SearchHit {
  sheet: string;
  row: number;
  col: number;
  text: string;
}

export interface SearchOptions {
  caseSensitive: boolean;
  wholeCell: boolean;
}

export interface BorderEdge {
  /** A CSS `border-style`. */
  line: "solid" | "dashed" | "dotted" | "double";
  width: number;
  /** Excel's border precedence — the heavier of two shared edges wins. */
  rank: number;
  color: string | null;
}

export interface CellBorder {
  top: BorderEdge | null;
  right: BorderEdge | null;
  bottom: BorderEdge | null;
  left: BorderEdge | null;
}

export interface CellStyle {
  /**
   * `centerContinuous` is Excel's Center Across Selection, which the grid draws
   * as one block over the run of cells carrying it — not centred in one cell.
   */
  horizontal: "left" | "center" | "centerContinuous" | "right" | null;
  vertical: "top" | "middle" | "bottom" | null;
  /** Indent levels; one is three spaces of the cell's font (ECMA-376 §18.8.1). */
  indent: number;
  wrap: boolean;
  /** Excel's shrink-to-fit: the font is scaled down until the cell's line fits. */
  shrink: boolean;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  /** `single` or `double`; null when the font carries no underline. */
  underline: "single" | "double" | null;
  /** Set only when the cell's font differs from the workbook default. */
  sizePx: number | null;
  /** Likewise, and it is what an `indent` is measured in. */
  family: string | null;
  color: string | null;
  /** The cell's background, from a solid fill. */
  fill: string | null;
  border: CellBorder | null;
}

export interface MergeRect {
  row: number;
  col: number;
  rows: number;
  cols: number;
}

/** A rectangle of cells, as the backend reports one back. */
export interface Block {
  row: number;
  col: number;
  rows: number;
  cols: number;
}

/**
 * One coloured run of a cell's rich text.
 *
 * A cell shows the concatenation of its runs, so the text says nothing new; what
 * a run adds is a colour for part of it, which is how a document makes one word
 * inside a sentence red.
 */
export interface RichRun {
  text: string;
  /** `#rrggbb`, or null when the run takes the cell's own colour. */
  color: string | null;
}

export interface AutoFilter {
  /** Sheet row holding the filter buttons. */
  row: number;
  col: number;
  /** Counts every row in the range, including the header row. */
  rows: number;
  cols: number;
  /** Absolute column index -> values the file already filters on. */
  applied: Record<string, string[]>;
}

export interface ColumnStyle {
  from: number;
  to: number;
  style: number;
}

export interface ShapeStop {
  /** 0 to 1 across the shape. */
  pos: number;
  color: string;
  alpha: number;
}

export interface ShapeFill {
  /** A solid colour, or a gradient's first stop. */
  color: string;
  /** Excel's `<a:alpha>`; 1 is opaque. */
  alpha: number;
  /** The stops of a gradient, in order. Empty for a solid fill. */
  stops: ShapeStop[];
  /** A linear gradient's direction in degrees, clockwise from pointing right. */
  angle: number;
}

/** An arrowhead at one end of a line. */
export interface ShapeArrow {
  /** `triangle`, `stealth`, `diamond`, `oval` or `arrow`. */
  kind: string;
  width: "sm" | "med" | "lg";
  length: "sm" | "med" | "lg";
}

export interface ShapeStroke {
  color: string;
  alpha: number;
  /** Already in pixels. */
  width: number;
  /** The file's own preset name, e.g. `lgDash`; the renderer maps it to a dash array. */
  dash: string | null;
  head: ShapeArrow | null;
  tail: ShapeArrow | null;
}

export interface ShapeRun {
  text: string;
  bold: boolean;
  italic: boolean;
}

export interface ShapeText {
  /** Paragraphs of runs; kept apart so the renderer can lay out lines. */
  paras: ShapeRun[][];
  sizePt: number;
  color: string;
  align: "left" | "center" | "right";
  valign: "top" | "middle" | "bottom";
  /** Excel's `wrap="none"` lets text spill past the shape, like an unstyled cell. */
  wrap: boolean;
  /** `lIns`, `tIns`, `rIns`, `bIns` in pixels. */
  insets: [number, number, number, number];
}

/**
 * A shape, picture or connector the sheet floats above its cells.
 *
 * Each corner is a *cell plus pixels into that cell* rather than an absolute
 * offset, so the renderer can resolve it against the grid's own column widths —
 * which is what keeps shapes in place once a filter hides rows.
 */
export interface Shape {
  /** The preset geometry name, `picture`, or `frame` for a chart we do not draw. */
  kind: string;
  fromCol: number;
  fromX: number;
  fromRow: number;
  fromY: number;
  toCol: number;
  toX: number;
  toRow: number;
  toY: number;
  /** Degrees clockwise. */
  rot: number;
  flipH: boolean;
  flipV: boolean;
  fill: ShapeFill | null;
  line: ShapeStroke | null;
  text: ShapeText | null;
  /** Pictures only: the media part inside the workbook, e.g. `xl/media/image3.png`. */
  media: string | null;
  /** Preset adjustments (`adj1`…), as fractions of the shape. Callouts use these. */
  adjust: number[];
}

export interface SheetLayout {
  /**
   * The sheet's real extent. calamine's range covers only cells that hold a
   * value, so it drops columns kept alive by styling or a merge.
   */
  rows: number;
  cols: number;
  /**
   * Column widths in the file's own unit — a multiple of the workbook font's
   * maximum digit width — not in pixels. Only the renderer can measure that
   * font, so the conversion happens there; see `charsToPx`.
   */
  defaultColWidthChars: number;
  defaultRowHeightPx: number;
  /** The workbook's default font, so text renders at the file's own scale. */
  fontFamily: string | null;
  fontSizePx: number;
  /** Excel's View > Gridlines setting. */
  showGridLines: boolean;
  /** Column index -> width in characters, for columns the file sized explicitly. */
  colWidthsChars: Record<string, number>;
  /** Row index -> height in pixels, for rows the file sized explicitly. */
  rowHeightsPx: Record<string, number>;
  merges: MergeRect[];
  autoFilter: AutoFilter | null;
  /** `_xlnm.Print_Area`; Excel shades everything outside it. */
  printArea: MergeRect | null;
  /**
   * The distinct styles in use. Cells index into this — repeating a whole
   * style object per cell made one sheet's payload 4.6 MB.
   */
  styles: CellStyle[];
  /**
   * Per row, runs of consecutive columns sharing one style, as
   * `[fromCol, toCol, styleIndex]`. One entry per cell cost 110k entries on a
   * single sheet; runs collapse a filled row into one.
   */
  cellRuns: Record<string, [number, number, number][]>;
  /** Rows whose unstyled cells inherit a default style. */
  rowStyles: Record<string, number>;
  /** Column runs whose unstyled cells inherit a default style. */
  colStyles: ColumnStyle[];
  /**
   * The shapes and pictures the sheet floats above its cells, in the order Excel
   * draws them — later ones sit on top.
   */
  drawings: Shape[];
  /**
   * Strings split into runs with colours of their own, keyed by the text they
   * spell out — which is exactly what a cell holds, so a cell's runs are found
   * without the sheet recording anything per cell. Empty for a workbook that
   * never colours part of a string.
   */
  richText: Record<string, RichRun[]>;
}

export interface ValueCount {
  value: string;
  count: number;
}

export interface FilterCriterion {
  col: number;
  values: string[];
}

export function openWorkbook(path: string): Promise<WorkbookInfo> {
  return invoke<WorkbookInfo>("open_workbook", { path });
}

export function getCells(
  path: string,
  sheet: string,
  row: number,
  col: number,
  rows: number,
  cols: number,
): Promise<Cell[][]> {
  return invoke<Cell[][]>("get_cells", { path, sheet, row, col, rows, cols });
}

/**
 * Excel's current region around a cell: the block that can be reached from it
 * without crossing a wholly blank row or column.
 */
export function currentRegion(
  path: string,
  sheet: string,
  row: number,
  col: number,
): Promise<Block> {
  return invoke<Block>("current_region", { path, sheet, row, col });
}

/** One cell on the sheet, in zero-based coordinates. */
export interface CellPos {
  row: number;
  col: number;
}

/**
 * Where Ctrl+arrow lands: the far end of the run of filled cells in that
 * direction, or the nearest filled cell when the one it starts from is blank.
 * `null` is nothing ahead at all — the grid answers that with the sheet's own
 * edge, which the backend has no view of.
 */
export function edgeCell(
  path: string,
  sheet: string,
  row: number,
  col: number,
  dir: "up" | "down" | "left" | "right",
): Promise<CellPos | null> {
  return invoke<CellPos | null>("edge_cell", { path, sheet, row, col, dir });
}

export function searchWorkbook(
  path: string,
  query: string,
  sheet: string | null,
  options: SearchOptions,
): Promise<SearchHit[]> {
  return invoke<SearchHit[]>("search_workbook", {
    path,
    query,
    sheet,
    caseSensitive: options.caseSensitive,
    wholeCell: options.wholeCell,
  });
}

export function sheetLayout(path: string, sheet: string): Promise<SheetLayout> {
  return invoke<SheetLayout>("sheet_layout", { path, sheet });
}

export function columnValues(
  path: string,
  sheet: string,
  col: number,
  firstRow: number,
  lastRow: number,
  query: string,
): Promise<ValueCount[]> {
  return invoke<ValueCount[]>("column_values", {
    path,
    sheet,
    col,
    firstRow,
    lastRow,
    query: query.trim() || null,
  });
}

/** Sheet row indices that satisfy every criterion, in ascending order. */
export function filteredRows(
  path: string,
  sheet: string,
  firstRow: number,
  lastRow: number,
  criteria: FilterCriterion[],
): Promise<number[]> {
  return invoke<number[]>("filter_rows", {
    path,
    sheet,
    firstRow,
    lastRow,
    criteria,
  });
}

/** Display text along one axis, for copying a whole row or column. */
export function axisValues(
  path: string,
  sheet: string,
  axis: { row: number } | { col: number },
  from: number,
  to: number,
): Promise<string[]> {
  return invoke<string[]>("axis_values", {
    path,
    sheet,
    row: "row" in axis ? axis.row : null,
    col: "col" in axis ? axis.col : null,
    from,
    to,
  });
}

export function closeWorkbook(path: string): Promise<void> {
  return invoke<void>("close_workbook", { path });
}

/**
 * The embedded images a sheet refers to, as base64, keyed by part.
 *
 * Asked for in one call per sheet: a design document's sheet can float fifty
 * screenshots, and one call per picture meant fifty archive reads and fifty
 * round trips. Parts the workbook does not hold come back missing.
 */
export function readMedia(
  path: string,
  parts: string[],
): Promise<Record<string, string>> {
  return invoke<Record<string, string>>("read_media", { path, parts });
}
