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
   * Whether the file hides the sheet. Excel draws no tab for one of these and
   * cannot put the cursor on one either, so it is left out of the strip, the
   * sheet list's reach and the search.
   */
  hidden: boolean;
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
  /**
   * Excel's one diagonal rule. The style lives in `<diagonal>`; the two flags
   * say which corners it joins, and both may be set — that is the crossing pair.
   */
  diagonal?: BorderEdge | null;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
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
  /**
   * Degrees turned counter-clockwise, from `<alignment textRotation>`; absent
   * for the cell that draws flat, which is nearly every cell.
   */
  rotation?: number;
  /** `textRotation="255"`: characters one below the next, each upright. */
  stacked?: boolean;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  /** `single` or `double`; null when the font carries no underline. */
  underline: "single" | "double" | null;
  /** Set when the cell's font draws raised or lowered (sup/subscript). */
  vertAlign?: "superscript" | "subscript" | null;
  /** Set only when the cell's font differs from the workbook default. */
  sizePx: number | null;
  /** Likewise, and it is what an `indent` is measured in. */
  family: string | null;
  color: string | null;
  /** The cell's background, from a solid fill. */
  fill: string | null;
  /** A hatch pattern in place of a flat background; a cell has one or the other. */
  pattern?: CellPattern | null;
  border: CellBorder | null;
}

/**
 * One of Excel's 18 hatch patterns. `kind` is the file's own `patternType`
 * name; the 8×8 masks live in the renderer.
 */
export interface CellPattern {
  kind: string;
  /** The hatch colour (`<fgColor>`). */
  fg: string;
  /** The ground under it (`<bgColor>`). */
  bg: string;
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
 * One run of a cell's rich text.
 *
 * A cell shows the concatenation of its runs, so the text says nothing new; what
 * a run adds is formatting for part of it, which is how a document makes one word
 * inside a sentence red, or bold, or a size larger.
 */
export interface RichRun {
  text: string;
  /** `#rrggbb`, or null when the run takes the cell's own colour. */
  color: string | null;
  /** The run's own `<rPr>`, or absent when it has none — see `RichFont`. */
  font?: RichFont | null;
}

/**
 * What a run's `<rPr>` says, where it differs from the cell's font.
 *
 * The two halves inherit differently, and the parser settles both from Excel's
 * own files: `bold`, `italic`, `strike` and `underline` are what the run *is*, so
 * an absent one means off rather than "the cell's"; `family` and `sizePx` are
 * overrides that the cell's font fills in when the rPr omits them.
 */
export interface RichFont {
  family?: string | null;
  sizePx?: number | null;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  /** `single`, `double`, or `none` for an rPr that spells the absence out. */
  underline?: "single" | "double" | "none" | null;
  /** The run draws raised or lowered at two thirds of its size. */
  vertAlign?: "superscript" | "subscript" | null;
}

/**
 * One cell's runs. Keyed by cell rather than by text: two cells can spell the
 * same text with different runs — a breadcrumb whose current page is red — and
 * the text alone cannot tell them apart.
 */
export interface RichCell {
  row: number;
  col: number;
  runs: RichRun[];
}

/** One cell's data bar: where it runs across the cell, and what it is drawn in. */
export interface DataBarCell {
  row: number;
  col: number;
  color: string;
  /** Where the bar starts and ends, as shares of the cell's width. */
  from: number;
  to: number;
  /** Whether the bar lies left of the axis; its dark end stays on the axis. */
  negative: boolean;
  /** The axis line's share of the cell, when Excel draws one. */
  axis?: number | null;
  axisColor?: string | null;
}

/** One cell's conditional icon: the family and the level drawn. */
export interface IconCell {
  row: number;
  col: number;
  family: string;
  level: number;
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
  /**
   * Columns filtered by a condition (`<customFilters>`, e.g. "> 20") rather than
   * by a value list. The rows they hide are the file's own; this only tells the
   * filter button to draw Excel's funnel.
   */
  custom: number[];
}

export interface DataValidation {
  /** The cells the rule governs; `sqref` may name several blocks. */
  ranges: MergeRect[];
  /** `list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom`, `none`. */
  kind: string;
  /** The rule as written: a quoted list (`"a,b"`), a range, or an expression. */
  formula1?: string | null;
  formula2?: string | null;
  allowBlank: boolean;
  /** The file's `showDropDown` is inverted: true hides the in-cell dropdown. */
  hideDropdown: boolean;
  showInput: boolean;
  showError: boolean;
  promptTitle?: string | null;
  prompt?: string | null;
  errorTitle?: string | null;
  error?: string | null;
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
  /** A hatch pattern (`<a:pattFill>`), when that is what filled the shape. */
  pattern?: ShapePattern | null;
  /** A radial gradient (`<a:gradFill><a:path>`), when it has a focus. */
  radial?: ShapeRadial | null;
}

/**
 * `<a:path path="circle|rect|shape">` with its `fillToRect`. Measured off
 * Excel: `circle` is a circle through the box's corners, `rect`/`shape` reach
 * the last stop at its edges.
 */
export interface ShapeRadial {
  /** `circle`, `rect` or `shape`. */
  kind: string;
  /** `fillToRect` insets as fractions — left, top, right, bottom. */
  focus: [number, number, number, number];
}

/** One of DrawingML's 54 preset patterns; the 8×8 masks live in the renderer. */
export interface ShapePattern {
  /** The `prst` value, e.g. `ltUpDiag`. */
  kind: string;
  /** The hatch colour, or null for Excel's automatic black. */
  fg?: string | null;
  /** The ground under it, or null for Excel's automatic white. */
  bg?: string | null;
}

/** An arrowhead at one end of a line. */
export interface ShapeArrow {
  /** `triangle`, `stealth`, `diamond`, `oval` or `arrow`. */
  kind: string;
  width: "sm" | "med" | "lg";
  length: "sm" | "med" | "lg";
}

/** One `pt` of a custom geometry: a number or a guide name, as written. */
export type CustomPoint = [string, string];

export interface ShapeCustomCommand {
  kind: "moveTo" | "lnTo" | "cubicBezTo" | "quadBezTo" | "arcTo" | "close";
  points: CustomPoint[];
  /** `arcTo`'s `wR hR stAng swAng`. */
  arc?: [string, string, string, string] | null;
}

export interface ShapeCustomPath {
  /** The path's own coordinate space, scaled onto the shape's box. */
  w: number;
  h: number;
  /** `none` means no fill; absent means `norm`. */
  fill?: string | null;
  /** Absent means stroked. */
  stroke?: boolean | null;
  commands: ShapeCustomCommand[];
}

/** A shape that carries its own geometry rather than naming a preset. */
export interface ShapeCustom {
  guides: { name: string; formula: string }[];
  paths: ShapeCustomPath[];
}

export interface ShapeStroke {  color: string;
  alpha: number;
  /** Already in pixels. */
  width: number;
  /** The file's own preset name, e.g. `lgDash`; the renderer maps it to a dash array. */
  dash: string | null;
  head: ShapeArrow | null;
  tail: ShapeArrow | null;
  /** `miter` when the file names none; `round` and `bevel` are its other two. */
  join?: "miter" | "round" | "bevel" | null;
  /** A miter join's limit as a multiple of the line width; Excel's default is 8. */
  miterLimit?: number | null;
}

export interface ShapeRun {
  text: string;
  bold: boolean;
  italic: boolean;
  /**
   * The face the run's own `<a:latin typeface>` names, theme slots already
   * resolved. Absent when it names none and inherits the shape's.
   */
  family?: string | null;
  /** The face's classification, for choosing a fallback when it is missing. */
  panose?: string | null;
  charset?: number | null;
  pitch?: number | null;
  /** The east asian and complex-script faces named beside the latin one. */
  ea?: string | null;
  cs?: string | null;
}

export interface ShapeText {
  /** Paragraphs of runs; kept apart so the renderer can lay out lines. */
  paras: ShapeRun[][];
  sizePt: number;
  color: string;
  align: "left" | "center" | "right";
  /** One alignment per paragraph, when they are not all `align`. */
  aligns?: ("left" | "center" | "right")[];
  valign: "top" | "middle" | "bottom";
  /** Excel's `wrap="none"` lets text spill past the shape, like an unstyled cell. */
  wrap: boolean;
  /** `lIns`, `tIns`, `rIns`, `bIns` in pixels. */
  insets: [number, number, number, number];
  /** The face the shape's own `defRPr` names, which its runs inherit. */
  family?: string | null;
  /** That face's classification, for runs that name none of their own. */
  panose?: string | null;
  charset?: number | null;
  pitch?: number | null;
  /** The east asian and complex-script faces named beside the latin one. */
  ea?: string | null;
  cs?: string | null;
  /** One flag per paragraph: right-to-left base direction. Absent when none is. */
  rtl?: boolean[];
}

/**
 * A shape, picture or connector the sheet floats above its cells.
 *
 * Each corner is a *cell plus pixels into that cell* rather than an absolute
 * offset, so the renderer can resolve it against the grid's own column widths —
 * which is what keeps shapes in place once a filter hides rows.
 */
export interface ShapeGroup {
  fromCol: number;
  fromX: number;
  fromRow: number;
  fromY: number;
  toCol: number;
  toX: number;
  toRow: number;
  toY: number;
  /** The shape's box as a fraction of the group's box. */
  fx: number;
  fy: number;
  fw: number;
  fh: number;
}

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
  /** Pictures only: the part of the image cut away, as [left, top, right, bottom] fractions. */
  crop?: [number, number, number, number];
  /**
   * Set when the shape sits inside a group. The group's box is given in cells —
   * so it follows the grid — and the shape's own box is a fraction of it, which
   * keeps a group exact without the parser's approximate column widths.
   */
  group?: ShapeGroup;
  /** Preset adjustments (`adj1`…), as fractions of the shape. Callouts use these. */
  adjust: number[];
  /** Set when the file carries its own geometry instead of naming a preset. */
  custom?: ShapeCustom | null;
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
  /** Rows and columns the sheet view freezes at the top and left. */
  frozenRows: number;
  frozenCols: number;
  /** Column index -> width in characters, for columns the file sized explicitly. */
  colWidthsChars: Record<string, number>;
  /** Row index -> height in pixels, for rows the file sized explicitly. */
  rowHeightsPx: Record<string, number>;
  merges: MergeRect[];
  /** Data-validation rules: what a cell accepts, and what Excel offers for it. */
  validations: DataValidation[];
  /**
   * Rows and columns the file hides. Excel draws none of them and skips the
   * number, so the gap is the only sign it gives.
   */
  hiddenRows: number[];
  hiddenCols: number[];
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
   * Cells split into runs of their own, each with its own colour or font. Sparse
   * by nature: only cells that really format part of their text appear, and two
   * cells spelling the same text may carry different runs.
   */
  richCells: RichCell[];
  /** Cells a data-bar rule draws in, with the share of the cell each fills. */
  dataBars: DataBarCell[];
  /** Cells an icon-set rule marks, with the level each takes. */
  icons: IconCell[];
  /** Cells carrying a note; Excel marks each with a red corner triangle. */
  notes?: { row: number; col: number }[];
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
