/**
 * The values a list data validation offers.
 *
 * The rule carries them either inline (`"a,b,c"`) or as a reference to the
 * cells holding them (`$R$47:$R$50`, optionally with a sheet name). A viewer
 * cannot offer Excel's choice — nothing here writes a value back — but the
 * list itself is what the file says the cell may hold, and that is worth
 * showing.
 */

import { getCells } from "@/lib/excel";

/** A cap on the cells read for one list, so a whole-column reference is safe. */
const MAX_CELLS = 10_000;

export async function validationListItems(
  path: string,
  sheet: string,
  formula: string | null | undefined,
): Promise<string[]> {
  if (!formula) return [];
  let text = formula.trim();
  if (text.startsWith("=")) text = text.slice(1).trim();

  if (text.startsWith("\"")) {
    const inner = text.slice(1, text.endsWith("\"") ? -1 : undefined).replace(/""/g, "\"");
    return inner
      .split(",")
      .map((item) => item.trim())
      .filter((item) => item.length > 0);
  }

  const [target, range] = splitSheet(text, sheet);
  const match = /^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/.exec(range.trim());
  if (!match) return [];
  const fromCol = columnIndex(match[1]);
  const fromRow = Number(match[2]) - 1;
  const toCol = match[3] ? columnIndex(match[3]) : fromCol;
  const toRow = match[4] ? Number(match[4]) - 1 : fromRow;
  if (fromCol < 0 || toCol < 0 || fromRow < 0 || toRow < 0) return [];
  const rows = toRow - fromRow + 1;
  const cols = toCol - fromCol + 1;
  if (rows <= 0 || cols <= 0 || rows * cols > MAX_CELLS) return [];

  const cells = await getCells(path, target, fromRow, fromCol, rows, cols);
  return cells
    .flat()
    .map((cell) => cell.text.trim())
    .filter((text) => text.length > 0);
}

/** `Sheet2!A1:B2` names its sheet; a bare `A1:B2` means the one on screen. */
function splitSheet(text: string, current: string): [string, string] {
  const match = /^(?:'((?:[^']|'')+)'|([^'!]+))!/.exec(text);
  if (!match) return [current, text];
  return [(match[1] ?? match[2]).replace(/''/g, "'"), text.slice(match[0].length)];
}

function columnIndex(letters: string): number {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * 26 + (letter.charCodeAt(0) - 64);
  }
  return value - 1;
}
