/**
 * The workbooks opened lately, newest first, the way a File menu lists them.
 *
 * This outlives the tabs on purpose. Where each sheet of a workbook was left
 * scrolled belongs to the tab and goes with it, but a workbook closed an hour
 * ago is often exactly the one being come back to, so the path is kept.
 */
const STORAGE_KEY = "sheetlens.recentFiles";

/**
 * The ones that would not open — moved, deleted, or on a share that is not
 * mounted. They stay in the list rather than being dropped, because a file that
 * has gone is usually a file that is coming back, and the menu only greys the
 * entry to say so.
 */
const MISSING_KEY = "sheetlens.missingFiles";

/** How many of them the File menu shows inline, before "More". */
export const RECENT_SHOWN = 10;

/** How many are kept at all, the overflow being what "More" holds. */
const KEEP = 50;

/** A list of paths under `key`, with anything that is not one dropped. */
function readList(key: string): string[] {
  try {
    const stored: unknown = JSON.parse(localStorage.getItem(key) ?? "null");
    if (Array.isArray(stored)) {
      return stored.filter((path): path is string => typeof path === "string");
    }
  } catch {
    // A webview that refuses storage simply has nothing to offer.
  }
  return [];
}

function writeList(key: string, paths: string[]): void {
  try {
    localStorage.setItem(key, JSON.stringify(paths));
  } catch {
    // As above: the workbook still opens, the list just does not survive.
  }
}

/** Newest first, and never the same path twice. */
export function readRecentFiles(): string[] {
  return readList(STORAGE_KEY);
}

/** The paths the menu should grey out. */
export function readMissingFiles(): string[] {
  return readList(MISSING_KEY);
}

/**
 * Puts `path` at the front, and hands back the list the menu should show now.
 * It opened, so whatever the list said about it before is no longer true.
 */
export function rememberFile(path: string): string[] {
  const next = [path, ...readRecentFiles().filter((other) => other !== path)].slice(
    0,
    KEEP,
  );
  writeList(STORAGE_KEY, next);
  writeList(
    MISSING_KEY,
    readMissingFiles().filter((other) => other !== path),
  );
  return next;
}

/** Records that `path` would not open, and hands back the paths to grey out. */
export function markFileMissing(path: string): string[] {
  const missing = readMissingFiles();
  if (!missing.includes(path)) writeList(MISSING_KEY, [...missing, path]);
  return [...missing, path];
}

/** Drops `path` from the list, and anything remembered about it with it. */
export function forgetFile(path: string): string[] {
  const next = readRecentFiles().filter((other) => other !== path);
  writeList(STORAGE_KEY, next);
  writeList(
    MISSING_KEY,
    readMissingFiles().filter((other) => other !== path),
  );
  return next;
}

/** Empties the list. The workbooks stay open; only the offer of them goes. */
export function forgetAllFiles(): void {
  writeList(STORAGE_KEY, []);
  writeList(MISSING_KEY, []);
}

/** The last segment of a path, which is all a one-line menu entry shows. */
export function fileName(path: string): string {
  return path.split(/[\\/]/).pop() || path;
}
