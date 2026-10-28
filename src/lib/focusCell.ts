/**
 * Focus cell: the row and the column of the cell the cursor is on, washed so
 * the cursor can be found at a glance.
 *
 * It is the app's own aid rather than the file's, and it exists for the one
 * case where the selection ring alone is not enough: a hit jumped to from the
 * search panel, or from Go To, lands somewhere the eye was not already looking.
 *
 * It belongs to the reader, not to a workbook, so it is remembered across
 * restarts — and it starts on, since a sheet that has to be switched on before
 * a search hit can be found is a sheet that fails at the thing this is for.
 */
const STORAGE_KEY = "sheetlens.focusCell";

export const DEFAULT_FOCUS_CELL = true;

export function initialFocusCell(): boolean {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored === "on") return true;
    if (stored === "off") return false;
  } catch {
    // A webview that refuses storage still gets the default.
  }
  return DEFAULT_FOCUS_CELL;
}

export function storeFocusCell(on: boolean): void {
  try {
    localStorage.setItem(STORAGE_KEY, on ? "on" : "off");
  } catch {
    // As above: the switch works, the choice just does not survive the window.
  }
}
