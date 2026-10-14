/**
 * The review marks: cells the reader has shaded to point at something.
 *
 * They are the app's own, not the file's — nothing here is written back, and a
 * mark lives only as long as the tab holding it. The key carries the sheet name
 * because one workbook's sheets share the map, and the workbook owns the map
 * because closing its tab is what forgets the marks, the same way it forgets
 * the cursors.
 */
import { areasOf, type Area, type Selection } from "@/lib/selection";

/** One cell's mark: a fill colour, keyed off the sheet and the cell. */
export type Marks = Map<string, string>;

/**
 * Excel's own Standard Colours row, so the palette is the one a reader already
 * knows rather than a set invented here.
 */
export const MARK_STANDARD = [
  "#C00000",
  "#FF0000",
  "#FFC000",
  "#FFFF00",
  "#92D050",
  "#00B050",
  "#00B0F0",
  "#0070C0",
  "#002060",
  "#7030A0",
];

/**
 * The same idea in light tints. Excel's standard row is for the file's own
 * formatting, where the fill is the point; a review mark sits *under* text that
 * still has to be read, so these are the ones to reach for when the strong
 * colours swallow the cell.
 */
export const MARK_SOFT = [
  "#FFC7CE",
  "#FCE4D6",
  "#FFEB9C",
  "#FFF2CC",
  "#E2EFDA",
  "#C6EFCE",
  "#DDEBF7",
  "#BDD7EE",
  "#E4DFEC",
  "#E7E6E6",
];

/** What the paint bucket starts out holding, and fills with on its own. */
export const DEFAULT_MARK = "#FFEB9C";

const STORAGE_KEY = "sheetlens.markColor";

/** Whether a stored value is one of the palette's, rather than anything at all. */
function isMarkColor(value: unknown): value is string {
  return (
    typeof value === "string" &&
    (MARK_STANDARD.includes(value) || MARK_SOFT.includes(value))
  );
}

/**
 * The colour the bucket last held.
 *
 * Excel keeps its own fill colour between sessions — the button still shows the
 * colour it was last given when it is opened the next day — so this is a
 * setting rather than part of a workbook. The marks themselves are not: they
 * belong to the tab that made them and are gone with it.
 */
export function initialMarkColor(): string {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (isMarkColor(stored)) return stored;
  } catch {
    // A webview that refuses storage still gets the default.
  }
  return DEFAULT_MARK;
}

export function storeMarkColor(fill: string): void {
  try {
    localStorage.setItem(STORAGE_KEY, fill);
  } catch {
    // As above: the colour changes, the choice just does not survive the window.
  }
}

export function markKey(sheet: string, row: number, col: number): string {
  // A NUL between the sheet and the cell: a sheet name may hold a colon, and
  // the pair has to be unambiguous without one.
  return `${sheet}\u0000${row}:${col}`;
}

/** The box one area covers, with a run of rows or columns reaching the data. */
function rectOf(
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
    // A whole row or column runs to the sheet's edge in Excel; here it stops at
    // the data, since marking a million empty cells is not what a fill means on
    // a sheet whose extent this app can see.
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
 * The marks after shading everything the selection covers, or after lifting the
 * shading off it when `fill` is null. Returns a new map, so React sees a new
 * value and repaints.
 *
 * Every area of a Ctrl+clicked selection is shaded: the bucket is pointed at
 * what is selected, and what is selected is all of it.
 */
export function markSelection(
  marks: Marks,
  sheet: string,
  selection: Selection,
  bounds: { rows: number; cols: number },
  fill: string | null,
): Marks {
  const next = new Map(marks);
  for (const area of areasOf(selection)) {
    const { row, col, rows, cols } = rectOf(area, bounds);
    for (let r = row; r < row + rows; r += 1) {
      for (let c = col; c < col + cols; c += 1) {
        const key = markKey(sheet, r, c);
        if (fill === null) next.delete(key);
        else next.set(key, fill);
      }
    }
  }
  return next;
}
