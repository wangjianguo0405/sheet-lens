/**
 * The workbooks opened lately, newest first, the way a File menu lists them.
 *
 * This outlives the tabs on purpose. Where each sheet of a workbook was left
 * scrolled belongs to the tab and goes with it, but a workbook closed an hour
 * ago is often exactly the one being come back to, so the path is kept.
 */
const STORAGE_KEY = "sheetlens.recentFiles";

/** How many of them the File menu shows inline, before "More". */
export const RECENT_SHOWN = 10;

/** How many are kept at all, the overflow being what "More" holds. */
const KEEP = 50;

/** Newest first, and never the same path twice. */
export function readRecentFiles(): string[] {
  try {
    const stored: unknown = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "null");
    if (Array.isArray(stored)) {
      return stored.filter((path): path is string => typeof path === "string");
    }
  } catch {
    // A webview that refuses storage simply has nothing to offer.
  }
  return [];
}

/** Puts `path` at the front, and hands back the list the menu should show now. */
export function rememberFile(path: string): string[] {
  const next = [path, ...readRecentFiles().filter((other) => other !== path)].slice(
    0,
    KEEP,
  );
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  } catch {
    // As above: the workbook still opens, the list just does not grow.
  }
  return next;
}

/** The last segment of a path, which is all a one-line menu entry shows. */
export function fileName(path: string): string {
  return path.split(/[\\/]/).pop() || path;
}
