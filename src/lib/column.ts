export function columnLabel(index: number): string {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function columnIndex(label: string): number {
  let n = 0;
  for (const ch of label.toUpperCase()) {
    const v = ch.charCodeAt(0) - 64;
    if (v < 1 || v > 26) return -1;
    n = n * 26 + v;
  }
  return n - 1;
}

export function cellRef(row: number, col: number): string {
  return `${columnLabel(col)}${row + 1}`;
}

/** Accepts `G408`, `$G$408` and `Sheet1!G408` (the sheet part is ignored). */
export function parseRef(input: string): { row: number; col: number } | null {
  const target = input.includes("!") ? input.slice(input.lastIndexOf("!") + 1) : input;
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(target.trim());
  if (!match) return null;
  const col = columnIndex(match[1]);
  if (col < 0) return null;
  const row = Number(match[2]) - 1;
  if (row < 0) return null;
  return { row, col };
}

/** A reference as it was typed: a cell, and the sheet it named if it named one. */
export interface Target {
  /** The sheet as written, or null for "the one already showing". */
  sheet: string | null;
  row: number;
  col: number;
}

/**
 * Reads a typed reference — `G408`, `Sheet1!G408`, `'My Sheet'!G408` — for the
 * name box and for Go To, which take the same thing.
 *
 * The sheet name comes back as it was written and no further: matching it
 * against the workbook is the caller's, since only the caller knows whether a
 * name that is not there is an error to report or nothing to do.
 */
export function parseTarget(input: string): Target | null {
  const trimmed = input.trim();
  if (!trimmed) return null;
  const bang = trimmed.lastIndexOf("!");
  const sheet = bang >= 0 ? trimmed.slice(0, bang).replace(/^'|'$/g, "").trim() : "";
  const cell = parseRef(bang >= 0 ? trimmed.slice(bang + 1) : trimmed);
  if (!cell) return null;
  return { sheet: sheet || null, row: cell.row, col: cell.col };
}
