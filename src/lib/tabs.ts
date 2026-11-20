import type { SheetInfo, WorkbookInfo } from "@/lib/excel";
import type { FilterPane, Formats, TextColors } from "@/lib/formats";
import type { SessionOutline } from "@/lib/outline";
import type { Selection } from "@/lib/selection";
import type { ViewMode } from "@/lib/viewMode";

/**
 * One open workbook: what it holds, and everything the reader has laid over it
 * for as long as the tab is open. The file itself is never written back to.
 */
export interface Tab {
  info: WorkbookInfo;
  sheet: string;
  /** Cursor per sheet, so switching away and back returns to where you were. */
  cursors: Record<string, Selection>;
  /**
   * The review marks, keyed by sheet and cell. They belong to the tab rather
   * than to the app, so closing it — like closing the file — forgets them; they
   * are never written back to the workbook.
   */
  marks: Map<string, string>;
  /**
   * The font and alignment the reader laid over the file from the ribbon,
   * keyed by sheet and cell. Like the marks they are the app's own, and the
   * tab closing is what forgets them.
   */
  formats: Formats;
  /**
   * Text the reader coloured by hand, as character ranges of a cell's value.
   * Kept beside the formats for the same reason: it is the app's own lens on
   * the file, and closing the tab is what forgets it.
   */
  textColors: TextColors;
  /**
   * The frozen bands the reader set or cleared from the ribbon, by sheet name.
   * An entry beats the file's own `<pane>`; zero rows and zero columns means
   * the file's freeze was cleared.
   */
  frozen: Record<string, { rows: number; cols: number }>;
  /**
   * The reader's filter decision by sheet: an area added from the ribbon, or a
   * clearing of the file's own. Like everything else the ribbon changes, it
   * lasts as long as the tab.
   */
  filterPane: Record<string, FilterPane>;
  /**
   * The outline groups the reader added or took away from the ribbon, by sheet.
   * The file's own levels are untouched; this is laid over them while the tab
   * is open, and forgotten with it.
   */
  sessionOutline: Record<string, SessionOutline>;
  /**
   * Where the reader split the sheet from the ribbon, by sheet name. Excel's
   * split is per window rather than per sheet; keeping it per sheet is the
   * same kind of decision the freeze and the filters already make.
   */
  splitAt: Record<string, { row: number; col: number }>;
  /**
   * The colours the reader gave the tabs, by sheet name. Excel writes a tab
   * colour into the file; this app has nowhere to write it, so the colour is its
   * own the way a mark is — it lives as long as the tab holding it, and a tab
   * left out of the map is drawn in whatever colour the workbook set.
   */
  tabColors: Map<string, string>;
  /**
   * Which sheets the reader has hidden or brought back, over the file's own
   * answer — the map holds a decision for the sheets they have touched and
   * nothing for the rest, which keeps the file's `state="hidden"` in charge of
   * what nobody has overruled. Like the tab colours, they last as long as the
   * tab: Excel writes the flag into the workbook, and nothing here does.
   */
  sheetHidden: Map<string, boolean>;
  /**
   * The view every sheet of this workbook opens in — the page view when the
   * file sets a print area anywhere, since a workbook that is printed is read
   * with an eye on the page. Only a default: each sheet is then switched on its
   * own, and a switch lasts as long as the tab.
   */
  defaultView: ViewMode;
  /** The sheets switched away from that default, by name. */
  views: Record<string, ViewMode>;
  /**
   * How many times this workbook has been read from disk: 0 when it was opened,
   * and one more each time it is reloaded. The path cannot say this — it is the
   * same file — so this is what tells the grid its layout and its blocks of
   * cells come from the workbook as it used to be.
   */
  revision: number;
}

/**
 * Whether one sheet is off the tab strip: what the file says, unless the reader
 * has already answered for that sheet.
 */
export function isHidden(tab: Tab, sheet: SheetInfo): boolean {
  return tab.sheetHidden.get(sheet.name) ?? sheet.hidden;
}

/** The sheets a tab's strip draws, in the workbook's own order. */
export function shownSheets(tab: Tab): SheetInfo[] {
  return tab.info.sheets.filter((sheet) => !isHidden(tab, sheet));
}
