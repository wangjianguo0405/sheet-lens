/**
 * The sheets fixed to the top of the left-hand sheet list, one list per
 * workbook.
 *
 * By workbook rather than by sheet name alone: the same name in two files is
 * two different sheets, and documents of one family tend to repeat their sheet
 * names. Only the names are kept — fixing a sheet moves nothing in the file,
 * and nothing is written back to it.
 */
import { folderKey } from "@/lib/recent";

const STORAGE_KEY = "sheetlens.pinnedSheets";

type Pins = Record<string, string[]>;

function readAll(): Pins {
  try {
    const stored: unknown = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "null");
    if (stored && typeof stored === "object" && !Array.isArray(stored)) {
      const pins: Pins = {};
      for (const [key, names] of Object.entries(stored as Record<string, unknown>)) {
        if (!Array.isArray(names)) continue;
        pins[key] = names.filter((name): name is string => typeof name === "string");
      }
      return pins;
    }
  } catch {
    // A webview that refuses storage simply has nothing to offer.
  }
  return {};
}

function writeAll(pins: Pins): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(pins));
  } catch {
    // As above: the list still reorders, the choice just does not survive.
  }
}

/** The sheets fixed in `path`, in the order they were fixed. */
export function readPinnedSheets(path: string): string[] {
  return readAll()[folderKey(path)] ?? [];
}

/** Writes the fixed order for `path` as given, and hands it back. */
export function setPinnedSheets(path: string, names: string[]): string[] {
  const all = readAll();
  const key = folderKey(path);
  if (names.length > 0) writeAll({ ...all, [key]: names });
  else {
    delete all[key];
    writeAll(all);
  }
  return names;
}

/** Fixes `name` in `path`, and hands back the names the list should show now. */
export function pinSheet(path: string, name: string): string[] {
  return setPinnedSheets(path, [
    name,
    ...readPinnedSheets(path).filter((other) => other !== name),
  ]);
}

/** Unfixes it, and hands back what is left. The sheet itself is untouched. */
export function unpinSheet(path: string, name: string): string[] {
  return setPinnedSheets(
    path,
    readPinnedSheets(path).filter((other) => other !== name),
  );
}
