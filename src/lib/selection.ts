import { cellRef, columnLabel } from "@/lib/column";

/**
 * What the grid currently has selected. Excel treats a whole row or column as
 * a first-class selection with its own address (`3:3`, `B:B`), so the model
 * keeps the shape rather than collapsing everything to a single cell. A block
 * of cells is the same idea: `A1:C5`, always stored as a normalised rectangle
 * with its top-left corner.
 */
export type Selection =
  | { kind: "cell"; row: number; col: number }
  | { kind: "row"; row: number }
  | { kind: "column"; col: number }
  | { kind: "range"; row: number; col: number; rows: number; cols: number };

export const ORIGIN: Selection = { kind: "cell", row: 0, col: 0 };

export function isCell(
  selection: Selection,
): selection is { kind: "cell"; row: number; col: number } {
  return selection.kind === "cell";
}

/** The rectangle two corners describe, in reading order. */
export function rangeBetween(
  from: { row: number; col: number },
  to: { row: number; col: number },
): Selection {
  return {
    kind: "range",
    row: Math.min(from.row, to.row),
    col: Math.min(from.col, to.col),
    rows: Math.abs(from.row - to.row) + 1,
    cols: Math.abs(from.col - to.col) + 1,
  };
}

/** Whether two spans on one axis share a cell. */
function overlap(a: number, aLen: number, b: number, bLen: number): boolean {
  return a < b + bLen && b < a + aLen;
}

/**
 * Whether a selection reaches any cell of a block.
 *
 * A merged region is one cell to Excel, so it lights up whole as soon as the
 * selection touches part of it. Comparing the selection's own corner against
 * the block — which is all a single cell needs — leaves every merged region
 * inside a dragged block unlit.
 */
export function selectionIntersects(
  selection: Selection,
  box: { row: number; col: number; rows: number; cols: number },
): boolean {
  switch (selection.kind) {
    case "row":
      return overlap(selection.row, 1, box.row, box.rows);
    case "column":
      return overlap(selection.col, 1, box.col, box.cols);
    case "cell":
      return (
        selection.row >= box.row &&
        selection.row < box.row + box.rows &&
        selection.col >= box.col &&
        selection.col < box.col + box.cols
      );
    case "range":
      return (
        overlap(selection.row, selection.rows, box.row, box.rows) &&
        overlap(selection.col, selection.cols, box.col, box.cols)
      );
  }
}

/** The address shown in the name box. */
export function selectionRef(selection: Selection): string {
  switch (selection.kind) {
    case "cell":
      return cellRef(selection.row, selection.col);
    case "row":
      return `${selection.row + 1}:${selection.row + 1}`;
    case "column":
      return `${columnLabel(selection.col)}:${columnLabel(selection.col)}`;
    case "range": {
      const first = cellRef(selection.row, selection.col);
      if (selection.rows === 1 && selection.cols === 1) return first;
      const last = cellRef(
        selection.row + selection.rows - 1,
        selection.col + selection.cols - 1,
      );
      return `${first}:${last}`;
    }
  }
}

export function selectionCovers(
  selection: Selection,
  row: number,
  col: number,
): boolean {
  switch (selection.kind) {
    case "cell":
      return selection.row === row && selection.col === col;
    case "row":
      return selection.row === row;
    case "column":
      return selection.col === col;
    case "range":
      return (
        row >= selection.row &&
        row < selection.row + selection.rows &&
        col >= selection.col &&
        col < selection.col + selection.cols
      );
  }
}

/** The cell a keyboard move should start from, for any selection shape. */
export function selectionAnchor(selection: Selection): {
  row: number;
  col: number;
} {
  switch (selection.kind) {
    case "cell":
      return { row: selection.row, col: selection.col };
    case "row":
      return { row: selection.row, col: 0 };
    case "column":
      return { row: 0, col: selection.col };
    case "range":
      // Arrow keys collapse a block back to the cell it started at, as Excel
      // does rather than moving the whole rectangle.
      return { row: selection.row, col: selection.col };
  }
}
