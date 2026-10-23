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

/**
 * The workbooks fixed to the start page by hand.
 *
 * Their own list rather than a flag on the history: a workbook is fixed because
 * it is wanted whatever was opened last, so it must not age out with the fifty
 * the history keeps — and emptying the history is not the same as unfixing
 * everything, since one is a privacy sweep and the other a choice.
 */
const PINNED_KEY = "sheetlens.pinnedFiles";

/**
 * The folders fixed to the start page's own column.
 *
 * Only their own paths are kept: the files in them are whatever is there when
 * the column is looked at, and this app is a viewer, so a folder is useful
 * exactly because it is a place to open something new from.
 */
const PINNED_DIRS_KEY = "sheetlens.pinnedDirs";

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

/**
 * The workbooks fixed to the start page, most recently fixed first — the same
 * order as the history, so that a workbook just fixed is the one at the top.
 */
export function readPinnedFiles(): string[] {
  return readList(PINNED_KEY);
}

/** Fixes `path` there, and hands back the list the page should show now. */
export function pinFile(path: string): string[] {
  const next = [
    path,
    ...readPinnedFiles().filter((other) => other !== path),
  ];
  writeList(PINNED_KEY, next);
  return next;
}

/**
 * Unfixes it, and hands back what is left. The workbook itself is not touched:
 * it goes back to being an ordinary entry of the history, or of nothing at all
 * if it has already aged out of that.
 */
export function unpinFile(path: string): string[] {
  const next = readPinnedFiles().filter((other) => other !== path);
  writeList(PINNED_KEY, next);
  return next;
}

/**
 * What makes two spellings of a folder the same folder. Windows hands back the
 * same path with either separator and any case, so both are folded there;
 * anywhere else the string is the folder.
 */
export function folderKey(path: string): string {
  const trimmed = path.replace(/[\\/]+$/, "");
  const windows =
    typeof navigator !== "undefined" && /Windows/i.test(navigator.userAgent);
  return windows ? trimmed.toLowerCase() : trimmed;
}

/** The folders fixed to the start page, most recently fixed first. */
export function readPinnedDirs(): string[] {
  return readList(PINNED_DIRS_KEY);
}

/** Fixes `path` there, and hands back the list the page should show now. */
export function pinDir(path: string): string[] {
  const key = folderKey(path);
  const next = [
    path,
    ...readPinnedDirs().filter((other) => folderKey(other) !== key),
  ];
  writeList(PINNED_DIRS_KEY, next);
  return next;
}

/** Unfixes it, and hands back what is left. */
export function unpinDir(path: string): string[] {
  const key = folderKey(path);
  const next = readPinnedDirs().filter((other) => folderKey(other) !== key);
  writeList(PINNED_DIRS_KEY, next);
  return next;
}

/**
 * The folders the given workbooks are in, newest first, without repeats.
 * A workbook with no folder in its path contributes none.
 */
export function foldersOf(paths: string[]): string[] {
  const seen = new Set<string>();
  const folders: string[] = [];
  for (const path of paths) {
    const folder = folderName(path);
    if (!folder) continue;
    const key = folderKey(folder);
    if (seen.has(key)) continue;
    seen.add(key);
    folders.push(folder);
  }
  return folders;
}

/** The last segment of a path, which is all a one-line menu entry shows. */
export function fileName(path: string): string {
  return path.split(/[\\/]/).pop() || path;
}

/**
 * Everything before that last segment — the folder a workbook is in, for the
 * places that have the room to say where it is as well as what it is called.
 * Empty for a path with no folder in it, which a row shows as nothing at all.
 */
export function folderName(path: string): string {
  const cut = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  return cut < 0 ? "" : path.slice(0, cut);
}
