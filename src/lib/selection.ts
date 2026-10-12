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
  | { kind: "row"; row: number; rows: number; anchor: number; active: number }
  | { kind: "column"; col: number; cols: number; anchor: number; active: number }
  | {
      kind: "range";
      row: number;
      col: number;
      rows: number;
      cols: number;
      /**
       * The corner the block is pinned at, and the cell walking away from it.
       * Excel keeps both apart and a normalised rectangle cannot say which of
       * its corners is which: Shift+Right followed by Shift+Left has to come
       * back rather than grow again, and a plain arrow leaves from the active
       * cell rather than from the rectangle's top-left.
       */
      anchor: { row: number; col: number };
      active: { row: number; col: number };
    };

/**
 * A block starting at A1 with its anchor sitting on its own top-left corner.
 *
 * That is what Ctrl+A produces, and it is the one block whose anchor is not a
 * corner the active cell walked to: Excel leaves the active cell on A1 however
 * large the block it has just selected, so `rangeBetween`, which puts the active
 * cell at the far end, cannot express it.
 */
export function rangeOver(rows: number, cols: number): Selection {
  return {
    kind: "range",
    row: 0,
    col: 0,
    rows,
    cols,
    anchor: { row: 0, col: 0 },
    active: { row: 0, col: 0 },
  };
}

/** A run of whole rows, by its two ends — the same shape as a cell block. */
export function rowsBetween(from: number, to: number): Selection {
  return {
    kind: "row",
    row: Math.min(from, to),
    rows: Math.abs(from - to) + 1,
    anchor: from,
    active: to,
  };
}

/** A run of whole columns, by its two ends. */
export function colsBetween(from: number, to: number): Selection {
  return {
    kind: "column",
    col: Math.min(from, to),
    cols: Math.abs(from - to) + 1,
    anchor: from,
    active: to,
  };
}

export const ORIGIN: Selection = { kind: "cell", row: 0, col: 0 };

export function isCell(
  selection: Selection,
): selection is { kind: "cell"; row: number; col: number } {
  return selection.kind === "cell";
}

/** The rectangle two corners describe, in reading order — and the two corners. */
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
    anchor: { row: from.row, col: from.col },
    active: { row: to.row, col: to.col },
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
      return overlap(selection.row, selection.rows, box.row, box.rows);
    case "column":
      return overlap(selection.col, selection.cols, box.col, box.cols);
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
      return `${selection.row + 1}:${selection.row + selection.rows}`;
    case "column":
      return `${columnLabel(selection.col)}:${columnLabel(selection.col + selection.cols - 1)}`;
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
      return row >= selection.row && row < selection.row + selection.rows;
    case "column":
      return col >= selection.col && col < selection.col + selection.cols;
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
      return { row: selection.active, col: 0 };
    case "column":
      return { row: 0, col: selection.active };
    case "range":
      // Excel moves from a block's active cell, not from its rectangle. So an
      // arrow after a drag leaves from the cell the pointer ended on, and a
      // Shift+arrow extension is measured from the end that was walking.
      return selection.active;
  }
}
