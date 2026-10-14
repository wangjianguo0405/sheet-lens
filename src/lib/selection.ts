import { cellRef, columnLabel } from "@/lib/column";

/**
 * One area of the sheet's selection: a single cell, a run of whole rows or
 * columns, or a block of cells.
 *
 * Excel treats a whole row or column as a first-class selection with its own
 * address (`3:3`, `B:B`), so the model keeps the shape rather than collapsing
 * everything to a single cell. A block of cells is the same idea: `A1:C5`,
 * always stored as a normalised rectangle with its top-left corner.
 *
 * A shape that was walked carries both of the ends it was walked between, and
 * the two are not interchangeable. `anchor` is the end the selection is pinned
 * at — Excel's active cell, drawn clear inside the fill, and the cell the arrows
 * leave from — while `edge` is the end that moved. A normalised rectangle cannot
 * say which of its corners is which, and Excel needs to know: Shift+click
 * redraws the block around the anchor rather than around the cell clicked last,
 * which is also what lets Shift+Right followed by Shift+Left come back rather
 * than grow again.
 */
export type Area =
  | { kind: "cell"; row: number; col: number }
  | { kind: "row"; row: number; rows: number; anchor: number; edge: number }
  | { kind: "column"; col: number; cols: number; anchor: number; edge: number }
  | {
      kind: "range";
      row: number;
      col: number;
      rows: number;
      cols: number;
      anchor: { row: number; col: number };
      edge: { row: number; col: number };
    };

/**
 * What the grid has selected: one area, or several after Ctrl+clicking.
 *
 * Several areas are still one selection, as they are in Excel rather than a set
 * of independent ones. `active` says which of them the keyboard is on, and it is
 * the one drawn with its active cell ringed — so the area the name box addresses
 * and the area the arrows leave from are the same area.
 */
export type Selection = Area | { kind: "areas"; areas: Area[]; active: number };

/** The areas a selection is made of, whether one shape or several. */
export function areasOf(selection: Selection): Area[] {
  return selection.kind === "areas" ? selection.areas : [selection];
}

/**
 * The area the selection is on: what the name box addresses, what the keyboard
 * leaves from, and the one drawn with its active cell ringed.
 */
export function activeArea(selection: Selection): Area {
  return selection.kind === "areas" ? selection.areas[selection.active] : selection;
}

/**
 * A block with its anchor and its edge both on its own top-left corner.
 *
 * That is what Ctrl+A produces, and it is the one block whose edge was not
 * walked to: Excel leaves the active cell at the block's corner however large
 * the region it has just selected, so `rangeBetween`, which always puts the edge
 * at the far end, cannot express it.
 */
export function rangeOver(
  row: number,
  col: number,
  rows: number,
  cols: number,
): Area {
  return {
    kind: "range",
    row,
    col,
    rows,
    cols,
    anchor: { row, col },
    edge: { row, col },
  };
}

/** A run of whole rows, by its two ends — the same shape as a cell block. */
export function rowsBetween(from: number, to: number): Area {
  return {
    kind: "row",
    row: Math.min(from, to),
    rows: Math.abs(from - to) + 1,
    anchor: from,
    edge: to,
  };
}

/** A run of whole columns, by its two ends. */
export function colsBetween(from: number, to: number): Area {
  return {
    kind: "column",
    col: Math.min(from, to),
    cols: Math.abs(from - to) + 1,
    anchor: from,
    edge: to,
  };
}

export const ORIGIN: Selection = { kind: "cell", row: 0, col: 0 };

/**
 * The rectangle two corners describe, in reading order — and the two corners.
 *
 * `from` is the corner the block is pinned at, so it becomes the anchor: the
 * block is drawn around it and grows away from it.
 */
export function rangeBetween(
  from: { row: number; col: number },
  to: { row: number; col: number },
): Area {
  return {
    kind: "range",
    row: Math.min(from.row, to.row),
    col: Math.min(from.col, to.col),
    rows: Math.abs(from.row - to.row) + 1,
    cols: Math.abs(from.col - to.col) + 1,
    anchor: { row: from.row, col: from.col },
    edge: { row: to.row, col: to.col },
  };
}

/** Whether two spans on one axis share a cell. */
function overlap(a: number, aLen: number, b: number, bLen: number): boolean {
  return a < b + bLen && b < a + aLen;
}

/** Whether one area reaches a cell. */
function areaCovers(area: Area, row: number, col: number): boolean {
  switch (area.kind) {
    case "cell":
      return area.row === row && area.col === col;
    case "row":
      return row >= area.row && row < area.row + area.rows;
    case "column":
      return col >= area.col && col < area.col + area.cols;
    case "range":
      return (
        row >= area.row &&
        row < area.row + area.rows &&
        col >= area.col &&
        col < area.col + area.cols
      );
  }
}

/**
 * Whether one area reaches any cell of a block.
 *
 * A merged region is one cell to Excel, so it lights up whole as soon as the
 * selection touches part of it. Comparing the selection's own corner against the
 * block — which is all a single cell needs — leaves every merged region inside a
 * dragged block unlit.
 */
function areaIntersects(
  area: Area,
  box: { row: number; col: number; rows: number; cols: number },
): boolean {
  switch (area.kind) {
    case "row":
      return overlap(area.row, area.rows, box.row, box.rows);
    case "column":
      return overlap(area.col, area.cols, box.col, box.cols);
    case "cell":
      return (
        area.row >= box.row &&
        area.row < box.row + box.rows &&
        area.col >= box.col &&
        area.col < box.col + box.cols
      );
    case "range":
      return (
        overlap(area.row, area.rows, box.row, box.rows) &&
        overlap(area.col, area.cols, box.col, box.cols)
      );
  }
}

/**
 * Whether the selection reaches a cell, anywhere in it.
 *
 * The one-area case is answered without a loop: this is asked once per cell of
 * every painted block, and a selection is one area until Ctrl+click makes it
 * otherwise.
 */
export function selectionCovers(selection: Selection, row: number, col: number): boolean {
  if (selection.kind !== "areas") return areaCovers(selection, row, col);
  return selection.areas.some((area) => areaCovers(area, row, col));
}

/** Whether the selection reaches any cell of a block, anywhere in it. */
export function selectionIntersects(
  selection: Selection,
  box: { row: number; col: number; rows: number; cols: number },
): boolean {
  if (selection.kind !== "areas") return areaIntersects(selection, box);
  return selection.areas.some((area) => areaIntersects(area, box));
}

/**
 * How far a selection reaches, for the status bar's readout.
 *
 * A whole row or column runs to the edge of the sheet, which the grid does not
 * know, so only the axis the user actually walked is reported. A lone cell
 * spans nothing worth saying. Several areas have no one span between them, so
 * they are counted instead.
 */
export type SelectionSize =
  | { kind: "rows"; rows: number }
  | { kind: "cols"; cols: number }
  | { kind: "block"; rows: number; cols: number }
  | { kind: "areas"; areas: number };

export function selectionSize(selection: Selection): SelectionSize | null {
  if (selection.kind === "areas") return { kind: "areas", areas: selection.areas.length };
  switch (selection.kind) {
    case "cell":
      return null;
    case "row":
      return { kind: "rows", rows: selection.rows };
    case "column":
      return { kind: "cols", cols: selection.cols };
    case "range":
      if (selection.rows === 1 && selection.cols === 1) return null;
      return { kind: "block", rows: selection.rows, cols: selection.cols };
  }
}

/** The address shown in the name box. */
export function selectionRef(selection: Selection): string {
  if (selection.kind === "areas") return selection.areas.map(areaRef).join(",");
  return areaRef(selection);
}

function areaRef(area: Area): string {
  switch (area.kind) {
    case "cell":
      return cellRef(area.row, area.col);
    case "row":
      return `${area.row + 1}:${area.row + area.rows}`;
    case "column":
      return `${columnLabel(area.col)}:${columnLabel(area.col + area.cols - 1)}`;
    case "range": {
      const first = cellRef(area.row, area.col);
      if (area.rows === 1 && area.cols === 1) return first;
      const last = cellRef(
        area.row + area.rows - 1,
        area.col + area.cols - 1,
      );
      return `${first}:${last}`;
    }
  }
}

/**
 * The cell a keyboard move starts from, whichever shape the selection is.
 *
 * It is the shape's anchor rather than its edge: Shift keeps walking the edge,
 * but a plain arrow has to leave from the cell the fill is drawn clear at, which
 * is the anchor. Excel leaves that cell where it was through a whole run of
 * Shift+arrows and Shift+clicks, so the cursor does not wander as the block
 * grows.
 */
export function selectionAnchor(selection: Selection): {
  row: number;
  col: number;
} {
  const area = activeArea(selection);
  switch (area.kind) {
    case "cell":
      return { row: area.row, col: area.col };
    case "row":
      return { row: area.anchor, col: 0 };
    case "column":
      return { row: 0, col: area.anchor };
    case "range":
      return area.anchor;
  }
}

/**
 * Ctrl+click: what was clicked joins the selection, and what it is already part
 * of comes off.
 *
 * The whole area comes off, whatever part of it was pointed at, and only an area
 * of the same kind as the click counts as "already part of it": a cell clicked
 * inside a run of whole rows is not that run, and Excel does not take the run
 * off for it. Letting the last area go would leave nothing selected, so a click
 * that would have leaves the clicked cell standing instead.
 */
export function toggleArea(
  selection: Selection,
  area: Area,
  hit: { kind: "cell" | "row" | "column"; row: number; col: number },
): Selection {
  const kind: Area["kind"] | null =
    hit.kind === "row" ? "row" : hit.kind === "column" ? "column" : null;
  const inTheWay = (one: Area) =>
    (kind === null ? one.kind === "cell" || one.kind === "range" : one.kind === kind) &&
    areaCovers(one, hit.row, hit.col);

  const areas = areasOf(selection);
  const at = areas.findIndex(inTheWay);
  if (at < 0) return withArea(selection, area);

  const left = areas.filter((_, index) => index !== at);
  if (left.length === 0) return area;
  const was = selection.kind === "areas" ? selection.active : 0;
  return {
    kind: "areas",
    areas: left,
    active: was > at ? was - 1 : Math.min(was, left.length - 1),
  };
}

/**
 * The selection with `area` added, as the active one.
 *
 * Ctrl+clicking somewhere the selection does not reach, and the block a
 * Ctrl+drag pulls out, both arrive here: Excel adds to the selection rather than
 * replacing it, and the area just added is the one it then works from.
 */
export function withArea(selection: Selection, area: Area): Selection {
  const areas = [...areasOf(selection), area];
  return { kind: "areas", areas, active: areas.length - 1 };
}
