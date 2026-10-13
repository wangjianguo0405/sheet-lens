/**
 * The two places the sheets can be picked from — the list panel on the left and
 * the tab strip under the grid. They are independent, so this is a pair of
 * switches rather than a choice between them: both, either or neither may be
 * showing. Which is why it is remembered across restarts.
 */
export interface SheetPanels {
  list: boolean;
  tabs: boolean;
}

const STORAGE_KEY = "sheetlens.sheetPanels";

export const DEFAULT_SHEET_PANELS: SheetPanels = { list: true, tabs: true };

/** What the user last left showing, if they ever changed it. */
export function initialSheetPanels(): SheetPanels {
  try {
    const stored = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "null");
    if (typeof stored?.list === "boolean" && typeof stored?.tabs === "boolean") {
      return { list: stored.list, tabs: stored.tabs };
    }
  } catch {
    // A webview that refuses storage still gets the defaults.
  }
  return DEFAULT_SHEET_PANELS;
}

export function storeSheetPanels(panels: SheetPanels): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(panels));
  } catch {
    // As above: the panels change, the choice just does not survive the window.
  }
}
