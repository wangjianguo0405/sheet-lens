/**
 * The reader's own folder tree: which folders it starts from, which of their
 * folders are open, and whether the tree shows beside the grid.
 *
 * Roots are folders rather than workbooks. The documents this reader is built
 * for sit under folders of the same name in every project branch, so the
 * thing worth remembering is where to start — a tree of roots reaches all of
 * them, where a list of leaf folders repeats itself.
 */
import { invoke } from "@tauri-apps/api/core";

export interface DirEntry {
  name: string;
  path: string;
  dir: boolean;
}

/** One level of a folder — its folders and the workbooks in it. */
export function listDirectory(path: string): Promise<DirEntry[]> {
  return invoke<DirEntry[]>("list_directory", { path });
}

/** One workbook a folder search found. */
export interface FoundFile {
  name: string;
  path: string;
  /** The root the search reached it under. */
  root: string;
  /** The folder it sits in, relative to that root. */
  folder: string;
}

/**
 * Finds workbooks by name under the given roots, walking every folder under
 * them — the tree itself lists a level only when it is opened, so it cannot
 * find a workbook the reader has not already walked down to.
 */
export function searchTree(roots: string[], query: string): Promise<FoundFile[]> {
  return invoke<FoundFile[]>("search_tree", { roots, query });
}

const ROOTS_KEY = "sheetlens.treeRoots";
const OPEN_KEY = "sheetlens.treeOpen";
const SHOWN_KEY = "sheetlens.treeShown";
const AUTO_LOCATE_KEY = "sheetlens.treeAutoLocate";

function readList(key: string): string[] {
  try {
    const stored: unknown = JSON.parse(localStorage.getItem(key) ?? "null");
    if (Array.isArray(stored)) {
      return stored.filter((value): value is string => typeof value === "string");
    }
  } catch {
    // A webview that refuses storage simply starts with nothing.
  }
  return [];
}

function writeList(key: string, values: string[]): string[] {
  try {
    localStorage.setItem(key, JSON.stringify(values));
  } catch {
    // As above: the list still changes, the choice just does not survive.
  }
  return values;
}

/** The folders the tree starts from, in the order they were added. */
export function readTreeRoots(): string[] {
  return readList(ROOTS_KEY);
}

export function storeTreeRoots(paths: string[]): string[] {
  return writeList(ROOTS_KEY, paths);
}

/** The folders left open, so the tree comes back as it was left. */
export function readTreeOpen(): string[] {
  return readList(OPEN_KEY);
}

export function storeTreeOpen(paths: string[]): string[] {
  return writeList(OPEN_KEY, paths);
}

/**
 * Whether the tree's own sidebar shows beside the grid. It is there unless the
 * reader has put it away: opening a workbook should not take away the list it
 * was opened from.
 */
export function readTreeShown(): boolean {
  try {
    return localStorage.getItem(SHOWN_KEY) !== "0";
  } catch {
    return true;
  }
}

export function storeTreeShown(shown: boolean): void {
  try {
    localStorage.setItem(SHOWN_KEY, shown ? "1" : "0");
  } catch {
    // As above.
  }
}

/**
 * Whether opening a workbook brings its row into sight by itself. It does not
 * unless the reader has asked for it: the header's locate button is the
 * default way to ask, the menu makes it happen every time.
 */
export function readTreeAutoLocate(): boolean {
  try {
    return localStorage.getItem(AUTO_LOCATE_KEY) === "1";
  } catch {
    return false;
  }
}

export function storeTreeAutoLocate(auto: boolean): void {
  try {
    localStorage.setItem(AUTO_LOCATE_KEY, auto ? "1" : "0");
  } catch {
    // As above.
  }
}
