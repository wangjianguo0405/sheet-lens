/**
 * Which of the start page's three columns are shown. The tree's flag is the
 * app's own `treeShown` — the same panel sits beside a workbook — so only the
 * two lists are kept here.
 */
const KEY = "sheetlens.welcomeColumns";

export interface WelcomeColumns {
  folders: boolean;
  files: boolean;
}

export function readWelcomeColumns(): WelcomeColumns {
  try {
    const stored: unknown = JSON.parse(localStorage.getItem(KEY) ?? "null");
    if (stored !== null && typeof stored === "object" && !Array.isArray(stored)) {
      const value = stored as Record<string, unknown>;
      return { folders: value.folders !== false, files: value.files !== false };
    }
  } catch {
    // A webview that refuses storage simply shows every column.
  }
  return { folders: true, files: true };
}

export function storeWelcomeColumns(columns: WelcomeColumns): WelcomeColumns {
  try {
    localStorage.setItem(KEY, JSON.stringify(columns));
  } catch {
    // The columns still change; the choice just does not survive.
  }
  return columns;
}
