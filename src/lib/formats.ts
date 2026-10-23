/**
 * The formatting the reader applies from the ribbon.
 *
 * None of it is written back to the workbook: this is a viewer, so a changed
 * font or alignment is a lens laid over the file, and it lives only as long as
 * the tab holding it — the same rule the review marks follow.
 */
import type { AutoFilter, CellStyle } from "@/lib/excel";
import { areasOf, type Area, type Selection } from "@/lib/selection";

/** The fields the ribbon can lay over a cell's own style. */
export type CellFormat = Partial<
  Pick<
    CellStyle,
    | "family"
    | "sizePx"
    | "bold"
    | "italic"
    | "underline"
    | "strike"
    | "horizontal"
    | "vertical"
    | "wrap"
  >
>;

/** One cell's overlay, keyed off the sheet and the cell. */
export type Formats = Map<string, CellFormat>;

/** What the ribbon shows for the active cell. */
export interface ActiveCellStyle {
  style: CellStyle | null;
  /** The family actually drawn: the cell's own, or the workbook's default. */
  family: string | null;
  /** The size actually drawn, in pixels. */
  sizePx: number;
}

/**
 * The reader's filter decision for one sheet: an area they added from the
 * ribbon, or a clearing of the one the file carries. `area` is ignored when
 * `off` is true.
 */
export interface FilterPane {
  area: AutoFilter | null;
  off: boolean;
}

/**
 * What a cell with no style of its own looks like before the overlay arrives.
 * The paint pass reads every field optionally, so this only has to be the
 * style the way the file would have written an empty one.
 */
export const PLAIN_STYLE: CellStyle = {
  horizontal: null,
  vertical: null,
  indent: 0,
  wrap: false,
  shrink: false,
  bold: false,
  italic: false,
  strike: false,
  underline: null,
  sizePx: null,
  family: null,
  color: null,
  fill: null,
  border: null,
};

export function formatKey(sheet: string, row: number, col: number): string {  // A NUL between the sheet and the cell, as the marks key theirs: a sheet name
  // may hold a colon, and the pair has to be unambiguous without one.
  return `${sheet}\u0000${row}:${col}`;
}

/** The cells an area covers, stopping at the data as the marks do. */
export function areaRect(
  area: Area,
  bounds: { rows: number; cols: number },
): { row: number; col: number; rows: number; cols: number } {
  const clampRows = (rows: number, row: number) =>
    Math.max(1, Math.min(rows, Math.max(0, bounds.rows - row)));
  const clampCols = (cols: number, col: number) =>
    Math.max(1, Math.min(cols, Math.max(0, bounds.cols - col)));
  switch (area.kind) {
    case "cell":
      return { row: area.row, col: area.col, rows: 1, cols: 1 };
    case "row":
      return {
        row: area.row,
        col: 0,
        rows: Math.min(area.rows, Math.max(1, bounds.rows)),
        cols: bounds.cols,
      };
    case "column":
      return {
        row: 0,
        col: area.col,
        rows: bounds.rows,
        cols: Math.min(area.cols, Math.max(1, bounds.cols)),
      };
    case "range":
      return {
        row: area.row,
        col: area.col,
        rows: clampRows(area.rows, area.row),
        cols: clampCols(area.cols, area.col),
      };
  }
}

/**
 * The overlays after laying `patch` over everything the selection covers, or
 * after taking the overlay off it when `patch` is null. Returns a new map, so
 * React sees a new value and repaints.
 */
export function formatSelection(
  formats: Formats,
  sheet: string,
  selection: Selection,
  bounds: { rows: number; cols: number },
  patch: CellFormat | null,
): Formats {
  const next = new Map(formats);
  for (const area of areasOf(selection)) {
    const { row, col, rows, cols } = areaRect(area, bounds);
    for (let r = row; r < row + rows; r += 1) {
      for (let c = col; c < col + cols; c += 1) {
        const key = formatKey(sheet, r, c);
        if (patch === null) next.delete(key);
        else next.set(key, { ...next.get(key), ...patch });
      }
    }
  }
  return next;
}

/** A cell's own style with the reader's overlay on top of it. */
export function withFormat(
  base: CellStyle | null,
  patch: CellFormat | undefined,
): CellStyle | null {
  if (!patch) return base;
  return { ...PLAIN_STYLE, ...base, ...patch };
}
