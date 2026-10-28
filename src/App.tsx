import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { open as openFileDialog } from "@tauri-apps/plugin-dialog";
import { EllipsisVertical } from "lucide-react";
import { useTranslation } from "react-i18next";

import { CanvasGrid, type SheetViews } from "@/components/CanvasGrid";
import { FileTabs } from "@/components/FileTabs";
import { Hint } from "@/components/Hint";
import { FormulaBar } from "@/components/FormulaBar";
import { GoToDialog } from "@/components/GoToDialog";
import { Ribbon } from "@/components/Ribbon";
import { SearchPanel } from "@/components/SearchPanel";
import { SheetList } from "@/components/SheetList";
import { NoteList } from "@/components/NoteList";
import { SheetScrollBar } from "@/components/SheetScrollBar";
import { SheetTabs } from "@/components/SheetTabs";
import { StatusBar } from "@/components/StatusBar";
import { Toaster } from "@/components/ui/sonner";
import { Toolbar } from "@/components/Toolbar";
import { UpdateDialog } from "@/components/UpdateDialog";
import { UnhideDialog, type UnhideAnchor } from "@/components/UnhideDialog";
import { Welcome } from "@/components/Welcome";
import { TooltipProvider } from "@/components/ui/tooltip";
import {
  ORIGIN,
  activeArea,
  areasOf,
  savedCursor,
  selectionAnchor,
  selectionRef,
  selectionSize,
  type Selection,
} from "@/lib/selection";
import { toggleShapeDebug } from "@/lib/debugPanels";
import { initialFocusCell, storeFocusCell } from "@/lib/focusCell";
import { initialMarkColor, markSelection, storeMarkColor } from "@/lib/marks";
import type { Target } from "@/lib/column";
import {
  initialSheetPanels,
  storeSheetPanels,
  type SheetPanels,
} from "@/lib/sheetPanels";
import {
  folderKey,
  folderName,
  forgetAllFiles,
  forgetFile,
  markFileMissing,
  readMissingFiles,
  readPinnedFiles,
  readRecentFiles,
  rememberFile,
  unpinDir,
  unpinFile,
} from "@/lib/recent";
import { checkForUpdate, updateStatus, type UpdateStatus } from "@/lib/update";
import { DEFAULT_ZOOM } from "@/lib/zoom";
import { initialView, type ViewMode } from "@/lib/viewMode";
import {
  closeWorkbook,
  currentRegion,
  openWorkbook,
  type AutoFilter,
  type Cell,
  type SearchHit,
  type SheetInfo,
  type WorkbookInfo,
} from "@/lib/excel";
import {
  areaRect,
  formatSelection,
  type ActiveCellStyle,
  type CellFormat,
  type FilterPane,
  type Formats,
} from "@/lib/formats";
import type { SessionOutline } from "@/lib/outline";

interface Tab {
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

const EXCEL_EXTENSIONS = ["xlsx", "xlsm", "xls", "xlsb", "ods"];

/**
 * Whether one sheet is off the tab strip: what the file says, unless the reader
 * has already answered for that sheet.
 */
function isHidden(tab: Tab, sheet: SheetInfo): boolean {
  return tab.sheetHidden.get(sheet.name) ?? sheet.hidden;
}

/** The sheets a tab's strip draws, in the workbook's own order. */
function shownSheets(tab: Tab): SheetInfo[] {
  return tab.info.sheets.filter((sheet) => !isHidden(tab, sheet));
}

const DEFAULT_SHEET_LIST_W = 240;
const MIN_SHEET_LIST_W = 160;
const MAX_SHEET_LIST_W = 480;

const DEFAULT_SEARCH_W = 288;
const MIN_SEARCH_W = 200;
const MAX_SEARCH_W = 640;

/**
 * The sheet's scrollbar takes a third of the row until it is dragged, and there
 * is no configured ceiling on it afterwards: what it may take is whatever the
 * tab strip beside it can spare, measured during the drag. A fixed one was wrong
 * twice over — on a narrow window the strip's own floor decides first, and the
 * default share can start out *wider* than the number, so the first touch of the
 * handle would snap the bar shorter instead of moving it.
 */
const MIN_SHEET_BAR_W = 120;

/**
 * The three things a drag can size: the sheet list, the search panel, and the
 * sheet's scrollbar at the end of the tab row.
 */
type DraggablePanel = "sheetList" | "search" | "sheetBar";

const PANEL_LIMITS: Record<DraggablePanel, [number, number]> = {
  sheetList: [MIN_SHEET_LIST_W, MAX_SHEET_LIST_W],
  search: [MIN_SEARCH_W, MAX_SEARCH_W],
  sheetBar: [MIN_SHEET_BAR_W, Number.POSITIVE_INFINITY],
};

function looksLikeWorkbook(path: string): boolean {
  const lower = path.toLowerCase();
  return EXCEL_EXTENSIONS.some((ext) => lower.endsWith(`.${ext}`));
}

export default function App() {
  const { t, i18n } = useTranslation();
  const [tabs, setTabs] = useState<Tab[]>([]);
  const [activePath, setActivePath] = useState<string | null>(null);
  /**
   * The tab strip's own multi-selection: Ctrl+click groups sheets and
   * Shift+click takes a run of them, as Excel's tabs do. The displayed sheet is
   * always one of them; what the group is for here is acting on several sheets
   * at once — Hide.
   */
  const [sheetGroup, setSheetGroup] = useState<{ names: string[]; anchor: string } | null>(null);
  const [searchOpen, setSearchOpen] = useState(false);
  /**
   * Bumped by Ctrl+F. The panel focuses its input when it opens; this is what
   * brings the keyboard back to it when it is already open, where pressing the
   * shortcut used to put the whole panel away.
   */
  const [searchFocus, setSearchFocus] = useState(0);
  const [gotoOpen, setGotoOpen] = useState(false);
  /** The sheet dialog Excel opens from a tab's own menu — see `SheetTabs`. */
  const [unhideOpen, setUnhideOpen] = useState(false);
  /**
   * Where the menu that opens it stood, so the dialog can take its place rather
   * than the centred spot the shared dialog would put it in.
   */
  const [unhideAt, setUnhideAt] = useState<UnhideAnchor | null>(null);
  const [detail, setDetail] = useState<Cell | null>(null);
  /** The active cell's style as the ribbon shows it, overlays included. */
  const [activeStyle, setActiveStyle] = useState<ActiveCellStyle | null>(null);
  /** The frozen bands in force, and whether a filter is, as the grid reports them. */
  const [frozenBands, setFrozenBands] = useState({ rows: 0, cols: 0 });
  const [filterOn, setFilterOn] = useState(false);
  /** What the grid actually shows, which can exceed calamine's range. */
  const [extent, setExtent] = useState<{ rows: number; cols: number } | null>(
    null,
  );
  /**
   * How many rows the active sheet's filter is hiding, or null when there is no
   * filter. The grid owns the filter, so it reports this up for the status bar
   * to say — the same footing the extent is on.
   */
  const [filtered, setFiltered] = useState<{
    shown: number;
    total: number;
  } | null>(null);
  const [error, setError] = useState<string | null>(null);
  /** Which of the two sheet pickers are showing; they are independent. */
  const [panels, setPanels] = useState<SheetPanels>(initialSheetPanels);
  /** Whether the cursor's row and column are washed, kept between sessions. */
  const [focusCell, setFocusCell] = useState(initialFocusCell);
  /** The colour the toolbar's paint bucket is holding, kept between sessions. */
  const [markColor, setMarkColor] = useState(initialMarkColor);
  /** The workbook paths the File menu offers, newest first. */
  const [recentFiles, setRecentFiles] = useState<string[]>(readRecentFiles);
  /** Those of them that would not open, which the menu greys out. */
  const [missingFiles, setMissingFiles] = useState<string[]>(readMissingFiles);
  /**
   * The grid's scroller. The sheet's scrollbar is drawn in the tab row, so it
   * has to be given the one thing it is a view of.
   */
  const [viewport, setViewport] = useState<HTMLDivElement | null>(null);
  /**
   * The tab row's free area, when it is drawing a split's horizontal bars. The
   * grid portals them there so they sit level with the tabs, as Excel draws
   * them; with the tabs hidden it draws them itself and this stays null.
   */
  const [splitBarHost, setSplitBarHost] = useState<HTMLDivElement | null>(null);
  /** Where the split line lies, in pixels from the grid's left edge. */
  const [splitLineX, setSplitLineX] = useState<number | null>(null);
  /** View scale as a percentage, the way Excel names it. */
  const [zoom, setZoom] = useState(DEFAULT_ZOOM);
  const [sheetListWidth, setSheetListWidth] = useState(DEFAULT_SHEET_LIST_W);
  /** Which panel the left sidebar shows: the sheet list, or the sheet's notes. */
  const [sidebarTab, setSidebarTab] = useState<"sheets" | "notes">("sheets");
  /** The note the sidebar is pointing at, to be shown in the sheet. */
  const [notePreview, setNotePreview] = useState<{ row: number; col: number } | null>(
    null,
  );
  const handleNotePreview = useCallback(
    (note: { row: number; col: number } | null) => setNotePreview(note),
    [],
  );
  const [searchWidth, setSearchWidth] = useState(DEFAULT_SEARCH_W);
  const [sheetBarWidth, setSheetBarWidth] = useState<number | null>(null);
  /** The row's own box, so a drag can start from the length it has now. */
  const sheetBarRef = useRef<HTMLDivElement>(null);
  /**
   * The same box, for a split's own horizontal bar to be drawn in: a split
   * across rows only has one column band, so its bar lives where the sheet's
   * ordinary one does and the splitter keeps working.
   */
  const setSheetBarHost = useCallback((element: HTMLDivElement | null) => {
    sheetBarRef.current = element;
    setSplitBarHost(element);
  }, []);
  /** The search panel's box, so Escape can tell whose key it was. */
  const searchPanelRef = useRef<HTMLDivElement>(null);
  /**
   * Which side panel a drag is sizing, so one set of listeners serves all of
   * them — they differ only in which way the pointer's delta counts, how far
   * they may be dragged, and where the result is stored.
   */
  const [draggingPanel, setDraggingPanel] = useState<DraggablePanel | null>(null);
  /** Where the drag started, so the width follows the pointer's delta. */
  const panelDrag = useRef<{ x: number; width: number } | null>(null);
  /** Only a portable build reports `portable`, and only then is it offered. */
  const [update, setUpdate] = useState<UpdateStatus | null>(null);
  const [updateAvailable, setUpdateAvailable] = useState(false);
  const [updateOpen, setUpdateOpen] = useState(false);
  /** The start page shown as a tab, created by the "+" and closed by its own ×. */
  const [homeOpen, setHomeOpen] = useState(false);
  /** The workbook to put back when that tab is closed. */
  const lastFileRef = useRef<string | null>(null);
  /**
   * Bumped when the cursor is placed somewhere the grid has to show, rather
   * than somewhere the user is walking to. Switching sheets is not that: the
   * grid puts the view back where each sheet was left, which is a position the
   * cursor does not get to overrule.
   */
  const [reveal, setReveal] = useState(0);
  /**
   * Where each sheet of each open workbook was left scrolled. It is the tabs
   * that own this — a clue about a workbook nobody has open is not worth
   * keeping — so `closeTab` drops the whole entry, the way the tab's cursors go
   * with it.
   */
  const views = useRef<SheetViews>(new Map());

  const active = useMemo(
    () => tabs.find((tab) => tab.info.path === activePath) ?? null,
    [tabs, activePath],
  );

  const activeSheet = useMemo(() => {
    if (!active) return null;
    return active.info.sheets.find((sheet) => sheet.name === active.sheet) ?? null;
  }, [active]);

  /** The tabs the strip draws, and the sheets Unhide offers, in sheet order. */
  const shown = useMemo(() => (active ? shownSheets(active) : []), [active]);
  const hiddenNames = useMemo(
    () => (active ? active.info.sheets.filter((s) => isHidden(active, s)).map((s) => s.name) : []),
    [active],
  );

  // A group belongs to the file it was made in.
  useEffect(() => {
    setSheetGroup(null);
  }, [activePath]);

  /**
   * The group as the strip draws it: only sheets still on the strip, and always
   * the one on screen. A sheet that gets hidden or reordered out of the group
   * simply drops out rather than leaving a stale name to act on.
   */
  const sheetSelection = useMemo(() => {
    const visible = new Set(shown.map((sheet) => sheet.name));
    const names = new Set((sheetGroup?.names ?? []).filter((name) => visible.has(name)));
    if (active && visible.has(active.sheet)) names.add(active.sheet);
    return names;
  }, [sheetGroup, shown, active]);

  const selection = active?.cursors[active.sheet] ?? ORIGIN;

  /** What the sheet on screen is drawn in: its own switch, or the default. */
  const view: ViewMode = active
    ? (active.views[active.sheet] ?? active.defaultView)
    : "normal";

  /** Whether the sheet on screen is split, which brings its own scrollbars. */
  const splitOn = Boolean(active?.splitAt[active.sheet]);

  const openPaths = useCallback(async (paths: string[]) => {
    setError(null);
    for (const path of paths) {
      try {
        const info = await openWorkbook(path);
        // Excel opens on the tab the file was saved with, not on the first
        // one; a hidden sheet cannot be current, so it is skipped.
        const opening =
          info.sheets.find((sheet) => sheet.selected && !sheet.hidden) ??
          info.sheets.find((sheet) => !sheet.hidden) ??
          info.sheets[0];
        const cursor = savedCursor(opening?.selection);
        setTabs((prev) =>
          prev.some((tab) => tab.info.path === info.path)
            ? prev
            : [
                ...prev,
                {
                  info,
                  // A workbook opens on the sheet it was saved on. One it
                  // hides has no tab to be current, which is the sheet Excel
                  // would be on too.
                  sheet: opening?.name ?? "",
                  cursors: cursor && opening ? { [opening.name]: cursor } : {},
                  marks: new Map(),
                  formats: new Map(),
                  frozen: {},
                  filterPane: {},
                  sessionOutline: {},
                  splitAt: {},
                  tabColors: new Map(),
                  sheetHidden: new Map(),
                  defaultView: initialView(info, opening),
                  views: {},
                  revision: 0,
                },
              ],
        );
        setActivePath(info.path);
        lastFileRef.current = info.path;
        setRecentFiles(rememberFile(info.path));
        setMissingFiles(readMissingFiles());
        // Opening a workbook closes the start page's tab: it was somewhere to
        // start from, not somewhere to come back to.
        setHomeOpen(false);
      } catch (reason: unknown) {
        setError(`${path} — ${String(reason)}`);
        // A drag can fail on a file that is not in the menu; marking it is
        // harmless, since only the ones that are get greyed.
        setMissingFiles(markFileMissing(path));
      }
    }
  }, []);

  /** The picker's own flow, with an optional folder to start it in. */
  const pickWorkbooks = useCallback(
    async (defaultPath?: string) => {
      try {
        const picked = await openFileDialog({
          ...(defaultPath ? { defaultPath } : {}),
          multiple: true,
          filters: [{ name: "Excel", extensions: EXCEL_EXTENSIONS }],
        });
        if (!picked) return;
        await openPaths(Array.isArray(picked) ? picked : [picked]);
      } catch (reason: unknown) {
        setError(String(reason));
      }
    },
    [openPaths],
  );

  const handleOpen = useCallback(() => pickWorkbooks(), [pickWorkbooks]);

  /**
   * A folder off the start page's own column: the same picker, already inside
   * it. The folder button is not a second way to open a workbook — it is the
   * same way, starting one step in.
   */
  const handleOpenInDirectory = useCallback(
    (directory: string) => pickWorkbooks(directory),
    [pickWorkbooks],
  );

  /** A path off the File menu's list, which opens exactly as the picker would. */
  const handleOpenRecent = useCallback(
    (path: string) => {
      void openPaths([path]);
    },
    [openPaths],
  );

  /** Takes one entry out of the File menu's list. Any tab it opened stays. */
  const handleForgetRecent = useCallback((path: string) => {
    setRecentFiles(forgetFile(path));
    setMissingFiles(readMissingFiles());
  }, []);

  const handleClearRecents = useCallback(() => {
    forgetAllFiles();
    setRecentFiles([]);
    setMissingFiles([]);
  }, []);

  /**
   * Takes a folder off the start page with everything it held: its pin, and
   * every workbook in it out of both file lists. The page lists a folder
   * because its workbooks come from there, so dropping the folder alone would
   * bring it back with the next workbook looked at.
   */
  const handleForgetFolder = useCallback((folder: string) => {
    const key = folderKey(folder);
    const inside = (path: string) => folderKey(folderName(path)) === key;
    for (const path of readRecentFiles()) if (inside(path)) forgetFile(path);
    for (const path of readPinnedFiles()) if (inside(path)) unpinFile(path);
    unpinDir(folder);
    setRecentFiles(readRecentFiles());
    setMissingFiles(readMissingFiles());
  }, []);

  /**
   * The "+" beside the file tabs. It opens the start page as a tab of its own —
   * and, pressed while the start page is already showing, opens the picker,
   * because that is what the start page is for.
   */
  const handleNewTab = useCallback(() => {
    if (homeOpen && activePath === null) {
      void handleOpen();
      return;
    }
    if (activePath !== null) lastFileRef.current = activePath;
    setHomeOpen(true);
    setActivePath(null);
  }, [homeOpen, activePath, handleOpen]);

  /** Its ×, which puts back whichever workbook was showing before it. */
  const closeHome = useCallback(() => {
    setHomeOpen(false);
    if (activePath !== null) return;
    const remembered = tabs.find((tab) => tab.info.path === lastFileRef.current);
    setActivePath(remembered?.info.path ?? tabs.at(-1)?.info.path ?? null);
  }, [activePath, tabs]);

  /** Remember where we were, so the start page has somewhere to go back to. */
  const selectTab = useCallback((path: string) => {
    lastFileRef.current = path;
    setActivePath(path);
  }, []);

  /**
   * A tab dragged along the strip. It moves in this list and nowhere else —
   * the order is the reader's, not the workbook's — and the tab showing stays
   * the tab showing, since the cursors travel with their own entry.
   */
  const handleReorderTabs = useCallback((from: number, to: number) => {
    setTabs((prev) => {
      if (from === to) return prev;
      if (from < 0 || from >= prev.length || to < 0 || to >= prev.length) return prev;
      const next = [...prev];
      const [moved] = next.splice(from, 1);
      next.splice(to, 0, moved);
      return next;
    });
  }, []);

  /**
   * A sheet tab dragged along the strip. The order is the reader's for as long
   * as the workbook is open — nothing is written back — and the sheet list
   * beside the grid and every Ctrl+PageUp/PageDown step read this same list, so
   * all three move together.
   *
   * The hidden sheets are not on the strip and so are not dragged past: each
   * keeps the place it held, and the sheets that could be seen are re-laid
   * around them.
   */
  const handleReorderSheets = useCallback(
    (from: number, to: number) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const order = shownSheets(tab);
          const [moved] = order.splice(from, 1);
          if (!moved) return tab;
          order.splice(to, 0, moved);
          let next = 0;
          const sheets = tab.info.sheets.map((sheet) =>
            isHidden(tab, sheet) ? sheet : order[next++],
          );
          return { ...tab, info: { ...tab.info, sheets } };
        }),
      );
    },
    [active],
  );

  const closeTab = useCallback(
    (path: string) => {
      void closeWorkbook(path).catch(() => {});
      views.current.delete(path);
      const remaining = tabs.filter((tab) => tab.info.path !== path);
      setTabs(remaining);
      if (activePath === path) setActivePath(remaining.at(-1)?.info.path ?? null);
    },
    [tabs, activePath],
  );

  /**
   * Closes several tabs at once: the strip's menu takes away the others, the
   * ones to its right, or all of them. `activate` is the tab the menu was
   * opened on — it survives the sweep and is the one left showing, the way a
   * browser leaves the tab you cleared the rest out from.
   */
  const closeTabs = useCallback(
    (paths: string[], activate?: string) => {
      if (paths.length === 0) return;
      const gone = new Set(paths);
      for (const path of paths) {
        void closeWorkbook(path).catch(() => {});
        views.current.delete(path);
      }
      const remaining = tabs.filter((tab) => !gone.has(tab.info.path));
      setTabs(remaining);
      if (activate !== undefined && !gone.has(activate)) setActivePath(activate);
      else if (activePath !== null && gone.has(activePath)) {
        setActivePath(remaining.at(-1)?.info.path ?? null);
      }
    },
    [tabs, activePath],
  );

  /**
   * Reads a workbook again from disk, for the file that changed underneath us —
   * Excel had it open, or something else wrote it.
   *
   * Only what the tab *remembered about the file* is thrown away: the sheets
   * that no longer exist, and a cursor or a current sheet that points at one.
   * What the reader put there themselves — the marks, the tab colours, the
   * sheets they hid, the widths they dragged, where each sheet was left scrolled
   * — is theirs, and survives the file being read again.
   */
  const reloadTab = useCallback(async (path: string) => {
    try {
      const info = await openWorkbook(path);
      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== path) return tab;
          const names = new Set(info.sheets.map((sheet) => sheet.name));
          const cursors = Object.fromEntries(
            Object.entries(tab.cursors).filter(([name]) => names.has(name)),
          );
          // The sheet showing is kept unless the file has taken it away or now
          // hides it — either way there is no tab for it to be current on.
          const here = info.sheets.find((sheet) => sheet.name === tab.sheet);
          const sheet =
            here && !isHidden(tab, here)
              ? here.name
              : (info.sheets.find((one) => !isHidden(tab, one))?.name ??
                info.sheets[0]?.name ??
                "");
          return {
            ...tab,
            info,
            cursors,
            sheet,
            // Which view the workbook opens in is read from the file, so a file
            // that has changed its print areas changes its answer too.
            defaultView: initialView(info, here),
            revision: tab.revision + 1,
          };
        }),
      );
      setError(null);
    } catch (reason: unknown) {
      setError(`${path} — ${String(reason)}`);
    }
  }, []);

  /**
   * File ▸ Close, which is the tab's × for whichever workbook is showing. It is
   * nothing at all while the start page is, which is why the entry is offered
   * disabled rather than left to do something surprising.
   */
  const handleCloseWorkbook = useCallback(() => {
    if (activePath !== null) closeTab(activePath);
  }, [activePath, closeTab]);

  const updateActive = useCallback(
    (patch: Partial<Omit<Tab, "info">>) => {
      setTabs((prev) =>
        prev.map((tab) => (tab.info.path === activePath ? { ...tab, ...patch } : tab)),
      );
    },
    [activePath],
  );

  /** Remembers where the cursor was, keyed by sheet. */
  const putCursor = useCallback(
    (sheetName: string, next: Selection) => {
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === activePath
            ? { ...tab, cursors: { ...tab.cursors, [sheetName]: next } }
            : tab,
        ),
      );
    },
    [activePath],
  );

  const handleSelect = useCallback(
    (next: Selection) => {
      if (!active) return;
      putCursor(active.sheet, next);
    },
    [active, putCursor],
  );

  /**
   * Shades the current selection in a review colour, or lifts the shading off
   * it when `fill` is null. The marks belong to the tab, so they last only as
   * long as the workbook is open and the file is never touched.
   */
  const applyMark = useCallback(
    (fill: string | null) => {
      if (!active || !activeSheet) return;
      // The bucket then holds what was just used, Excel's own behaviour: the
      // next press of the bucket is the same colour without opening anything.
      if (fill !== null) setMarkColor(fill);
      // The extent the grid is drawing, not `activeSheet`'s, which is calamine's
      // range: that spans only cells holding a value, so a column kept alive by
      // a border or a merge falls outside it and the fill stopped short of what
      // is on screen. The grid reports its own extent; before the first report
      // the parser's range is all there is.
      const bounds = extent ?? { rows: activeSheet.rows, cols: activeSheet.cols };
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === active.info.path
            ? {
                ...tab,
                marks: markSelection(
                  tab.marks,
                  tab.sheet,
                  tab.cursors[tab.sheet] ?? ORIGIN,
                  bounds,
                  fill,
                ),
              }
            : tab,
        ),
      );
    },
    [active, activeSheet, extent],
  );

  /**
   * Lays a font or alignment choice over the selection, or takes the overlay
   * off it when `patch` is null. It belongs to the tab and never reaches the
   * file, exactly as a mark does.
   */
  const applyFormat = useCallback(
    (patch: CellFormat | null) => {
      if (!active || !activeSheet) return;
      const bounds = extent ?? { rows: activeSheet.rows, cols: activeSheet.cols };
      const cursor = active.cursors[active.sheet] ?? ORIGIN;
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === active.info.path
            ? {
                ...tab,
                formats: formatSelection(tab.formats, active.sheet, cursor, bounds, patch),
              }
            : tab,
        ),
      );
    },
    [active, activeSheet, extent],
  );

  /**
   * Freezes the sheet above and to the left of the active cell, or clears
   * whatever freeze is in force — the file's included, which is why clearing
   * is a zero-sized override rather than a return to null. Nothing reaches the
   * workbook; the tab is what remembers it.
   */
  const applyFreeze = useCallback(
    (next: { rows: number; cols: number }) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === active.info.path
            ? { ...tab, frozen: { ...tab.frozen, [tab.sheet]: next } }
            : tab,
        ),
      );
    },
    [active],
  );

  /**
   * Splits the sheet at the active cell, or closes the split. Like the freeze
   * this is a view decision kept on the tab and never written to the file.
   */
  const applySplit = useCallback(
    (next: { row: number; col: number } | null) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const splitAt = { ...tab.splitAt };
          if (next) splitAt[tab.sheet] = next;
          else delete splitAt[tab.sheet];
          return { ...tab, splitAt };
        }),
      );
    },
    [active],
  );

  const toggleSplit = useCallback(() => {
    if (!active) return;
    if (active.splitAt[active.sheet]) {
      applySplit(null);
      return;
    }
    const at = selectionAnchor(active.cursors[active.sheet] ?? ORIGIN);
    // The lines cross at the active cell's top-left corner. A freeze in the
    // file is released first — it is not where the split goes.
    //
    // A cell in the first row or the first column splits only the other axis,
    // as Excel does: there is nothing above row 1 or left of column A to put in
    // a pane (measured: Split at A5 is SplitRow=4/SplitColumn=0, two panes). At
    // A1 itself neither axis has anywhere to go and Excel halves the window
    // instead, which the grid turns the zero/zero into.
    applySplit({ row: at.row, col: at.col });
    applyFreeze({ rows: 0, cols: 0 });
    setFrozenBands({ rows: 0, cols: 0 });
  }, [active, applySplit, applyFreeze]);

  const toggleFreeze = useCallback(() => {
    if (!active) return;
    if (frozenBands.rows > 0 || frozenBands.cols > 0) {
      applyFreeze({ rows: 0, cols: 0 });
      return;
    }
    const at = selectionAnchor(active.cursors[active.sheet] ?? ORIGIN);
    applyFreeze({ rows: at.row, cols: at.col });
    // A frozen pane and a split are two answers to the same question, and
    // Excel lets only one of them stand.
    applySplit(null);
  }, [active, frozenBands, applyFreeze, applySplit]);

  /**
   * Adds a filter over the current region, or takes away whichever filter is
   * in force. Like the freeze, this is a decision kept on the tab.
   */
  const setFilterDecision = useCallback(
    (pane: FilterPane) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === active.info.path
            ? { ...tab, filterPane: { ...tab.filterPane, [tab.sheet]: pane } }
            : tab,
        ),
      );
    },
    [active],
  );

  const toggleFilter = useCallback(async () => {
    if (!active) return;
    if (filterOn) {
      setFilterDecision({ area: null, off: true });
      setFilterOn(false);
      return;
    }
    const at = selectionAnchor(active.cursors[active.sheet] ?? ORIGIN);
    try {
      const block = await currentRegion(
        active.info.path,
        active.sheet,
        at.row,
        at.col,
      );
      const area: AutoFilter = {
        row: block.row,
        col: block.col,
        rows: block.rows,
        cols: block.cols,
        applied: {},
        custom: [],
      };
      setFilterDecision({ area, off: false });
      setFilterOn(true);
    } catch (reason: unknown) {
      setError(String(reason));
    }
  }, [active, filterOn, setFilterDecision]);

  /**
   * Groups or ungroups the rows or columns the selection covers. Excel asks
   * which axis from the Group dialog; the ribbon's little menu asks instead.
   * Nothing is written back — the grid lays these over the file's own levels.
   */
  const applyOutline = useCallback(
    (action: "rows" | "cols" | "ungroup") => {
      if (!active || !activeSheet) return;
      const bounds = extent ?? { rows: activeSheet.rows, cols: activeSheet.cols };
      const cursor = active.cursors[active.sheet] ?? ORIGIN;
      const rects = areasOf(cursor).map((area) => areaRect(area, bounds));
      if (rects.length === 0) return;
      const rowFrom = Math.min(...rects.map((rect) => rect.row));
      const rowTo = Math.max(...rects.map((rect) => rect.row + rect.rows - 1));
      const colFrom = Math.min(...rects.map((rect) => rect.col));
      const colTo = Math.max(...rects.map((rect) => rect.col + rect.cols - 1));

      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const current: SessionOutline = tab.sessionOutline[tab.sheet] ?? {
            added: { rows: [], cols: [] },
            removed: { rows: [], cols: [] },
          };
          const next: SessionOutline = {
            added: {
              rows: [...current.added.rows],
              cols: [...current.added.cols],
            },
            removed: {
              rows: [...current.removed.rows],
              cols: [...current.removed.cols],
            },
          };
          if (action === "rows") {
            next.added.rows.push([rowFrom, rowTo]);
          } else if (action === "cols") {
            next.added.cols.push([colFrom, colTo]);
          } else {
            // Ungroup follows the shape of what is selected, as Excel does: a
            // run of whole rows is a row group, a run of columns a column one,
            // and anything else is taken as rows — the axis a cell selection
            // groups by default.
            const axis = activeArea(cursor).kind === "column" ? "cols" : "rows";
            const [from, to] = axis === "rows" ? [rowFrom, rowTo] : [colFrom, colTo];
            next.removed[axis].push([from, to]);
            next.added[axis] = next.added[axis].filter(
              ([start, end]) => end < from || start > to,
            );
          }
          return {
            ...tab,
            sessionOutline: { ...tab.sessionOutline, [tab.sheet]: next },
          };
        }),
      );
    },
    [active, activeSheet, extent],
  );

  /**
   * Colours a sheet's tab, or takes the reader's colour off it when `color` is
   * null — which hands the tab back to the colour the workbook gives it, not to
   * no colour at all. Like a mark, it belongs to the tab and the file is never
   * touched.
   */
  const applyTabColor = useCallback(
    (sheet: string, color: string | null) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const tabColors = new Map(tab.tabColors);
          if (color === null) tabColors.delete(sheet);
          else tabColors.set(sheet, color);
          return { ...tab, tabColors };
        }),
      );
    },
    [active],
  );

  /**
   * The strip's own clicks, with Excel's modifiers: a plain click selects one
   * sheet and shows it; Ctrl+click adds or removes one from the group; Shift
   * +click takes the run between the anchor and the clicked sheet. The sheet
   * that was clicked is the one the window shows, as in Excel.
   */
  const handleSheetClick = useCallback(
    (name: string, mods: { ctrl: boolean; shift: boolean }) => {
      if (!active) return;
      const order = shownSheets(active).map((sheet) => sheet.name);
      const at = order.indexOf(name);
      if (at < 0) return;

      if (mods.shift) {
        const anchor =
          sheetGroup?.anchor && order.includes(sheetGroup.anchor)
            ? sheetGroup.anchor
            : active.sheet;
        const from = order.indexOf(anchor);
        if (from < 0) return;
        const [lo, hi] = from <= at ? [from, at] : [at, from];
        setSheetGroup({ names: order.slice(lo, hi + 1), anchor });
        updateActive({ sheet: name });
        return;
      }

      if (mods.ctrl) {
        const current = (sheetGroup?.names ?? [active.sheet]).filter((one) =>
          order.includes(one),
        );
        const names = new Set(current);
        if (names.has(name)) {
          // One sheet always stays selected; Excel does not let the group empty.
          if (names.size === 1) return;
          names.delete(name);
        } else {
          names.add(name);
        }
        const kept = order.filter((one) => names.has(one));
        const anchor =
          sheetGroup?.anchor && kept.includes(sheetGroup.anchor) ? sheetGroup.anchor : name;
        setSheetGroup({ names: kept, anchor });
        // Adding a sheet shows it; removing the one on screen moves to one that
        // is still selected.
        updateActive({ sheet: names.has(name) ? name : (kept[kept.length - 1] ?? active.sheet) });
        return;
      }

      setSheetGroup(null);
      updateActive({ sheet: name });
    },
    [active, sheetGroup, updateActive],
  );

  /**
   * Hides the sheets on the strip that the menu asked for — the whole group
   * when the right-clicked tab is part of it, that tab alone otherwise, which
   * is how Excel reads the gesture. Hiding everything still shown is refused:
   * it would leave the workbook with nothing to display.
   */
  const applySheetsHidden = useCallback(
    (names: string[]) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const visible = shownSheets(tab);
          const hiding = new Set(names.filter((name) => visible.some((one) => one.name === name)));
          if (hiding.size === 0 || hiding.size >= visible.length) return tab;

          const sheetHidden = new Map(tab.sheetHidden);
          for (const name of hiding) sheetHidden.set(name, true);
          let sheet = tab.sheet;
          if (hiding.has(sheet)) {
            const stillShown = new Set(
              shownSheets({ ...tab, sheetHidden }).map((one) => one.name),
            );
            const at = tab.info.sheets.findIndex((one) => one.name === sheet);
            const next =
              tab.info.sheets.slice(at + 1).find((one) => stillShown.has(one.name)) ??
              [...tab.info.sheets.slice(0, at)]
                .reverse()
                .find((one) => stillShown.has(one.name));
            sheet = next?.name ?? sheet;
          }
          return { ...tab, sheetHidden, sheet };
        }),
      );
      setSheetGroup(null);
    },
    [active],
  );

  const handleHideSheet = useCallback(
    (name: string) => {
      const group = sheetGroup?.names;
      applySheetsHidden(group && group.includes(name) ? group : [name]);
    },
    [sheetGroup, applySheetsHidden],
  );

  /**
   * Hides a sheet, or brings one back, for as long as the tab is open — the
   * file's own flag is never touched, and the sheet comes back as the workbook
   * has it when the workbook is opened again.
   *
   * Hiding the sheet on screen moves to the tab that takes its place: the next
   * one still shown, or the one before it when it was the last. That is where
   * Excel's own selection lands, measured by hiding a middle tab and an end one
   * and reading `ActiveSheet` back.
   */
  const applySheetHidden = useCallback(
    (sheet: string, hidden: boolean) => {
      if (!active) return;
      setTabs((prev) =>
        prev.map((tab) => {
          if (tab.info.path !== active.info.path) return tab;
          const sheetHidden = new Map(tab.sheetHidden);
          sheetHidden.set(sheet, hidden);
          if (!hidden || tab.sheet !== sheet) return { ...tab, sheetHidden };

          const shownHere = new Set(shownSheets({ ...tab, sheetHidden }).map((one) => one.name));
          const at = tab.info.sheets.findIndex((one) => one.name === sheet);
          const next =
            tab.info.sheets.slice(at + 1).find((one) => shownHere.has(one.name)) ??
            [...tab.info.sheets.slice(0, at)].reverse().find((one) => shownHere.has(one.name));
          return { ...tab, sheetHidden, sheet: next?.name ?? tab.sheet };
        }),
      );
    },
    [active],
  );

  /**
   * Bringing sheets back makes the last one picked the current one, as Excel's
   * Unhide does — you asked to see it, so it is what the window shows.
   */
  const handleUnhideSheets = useCallback(
    (names: string[]) => {
      for (const name of names) applySheetHidden(name, false);
      setSheetGroup(null);
      if (names.length) updateActive({ sheet: names[names.length - 1] });
    },
    [applySheetHidden, updateActive],
  );

  const handleSheetSelect = useCallback(
    (name: string) => {
      setSheetGroup(null);
      updateActive({ sheet: name });
    },
    [updateActive],
  );

  /**
   * Remembers which view a sheet was switched to. Keyed by sheet, like the
   * cursors: switching away and back returns the sheet to the view it was left
   * in, rather than to the one the workbook opened with.
   */
  const putView = useCallback(
    (sheetName: string, next: ViewMode) => {
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === activePath
            ? { ...tab, views: { ...tab.views, [sheetName]: next } }
            : tab,
        ),
      );
    },
    [activePath],
  );

  /** Which of Excel's two views the sheet on screen is being read in. */
  const handleView = useCallback(
    (next: ViewMode) => {
      if (!active) return;
      putView(active.sheet, next);
    },
    [active, putView],
  );

  /**
   * Each switch flips only its own half. Taking the whole pair from the render
   * this handler belongs to would let two presses before the next render
   * overwrite one another — and the second press would restore what the first
   * one turned off.
   */
  const togglePanel = useCallback((which: keyof SheetPanels) => {
    setPanels((prev) => ({ ...prev, [which]: !prev[which] }));
  }, []);

  useEffect(() => {
    storeSheetPanels(panels);
  }, [panels]);

  useEffect(() => {
    storeFocusCell(focusCell);
  }, [focusCell]);

  useEffect(() => {
    storeMarkColor(markColor);
  }, [markColor]);

  /**
   * Ctrl+PageUp/PageDown steps along the workbook's sheets in tab order, and
   * stops at either end rather than wrapping — which is what Excel does. It
   * steps over the hidden ones: they are not on the strip, and Excel's own
   * keys walk what is on it.
   */
  const handleSheetStep = useCallback(
    (delta: number) => {
      if (!active) return;
      const names = shownSheets(active).map((sheet) => sheet.name);
      const at = names.indexOf(active.sheet);
      if (at < 0) return;
      const next = at + delta;
      if (next < 0 || next >= names.length) return;
      updateActive({ sheet: names[next] });
    },
    [active, updateActive],
  );

  const handleJump = useCallback(
    (hit: SearchHit) => {
      updateActive({ sheet: hit.sheet });
      putCursor(hit.sheet, { kind: "cell", row: hit.row, col: hit.col });
      setReveal((n) => n + 1);
    },
    [updateActive, putCursor],
  );

  const handleGo = useCallback(
    (sheet: string, row: number, col: number) => {
      updateActive({ sheet });
      putCursor(sheet, { kind: "cell", row, col });
      setReveal((n) => n + 1);
    },
    [updateActive, putCursor],
  );

  /**
   * A reference typed into the name box goes where Go To goes. A sheet name that
   * is not in the workbook is not an error here — there is no dialog to report
   * it in — so the selection stays where it was and the box goes back to saying
   * where that is.
   */
  const handleNameBoxGo = useCallback(
    (target: Target) => {
      if (!active) return;
      const sheet = target.sheet
        ? active.info.sheets.find((sheet) => sheet.name === target.sheet)?.name
        : active.sheet;
      if (!sheet) return;
      handleGo(sheet, target.row, target.col);
      // Back to the grid, or the arrow keys would be typing into the name box.
      viewport?.focus({ preventScroll: true });
    },
    [active, handleGo, viewport],
  );

  const handleExtent = useCallback((rows: number, cols: number) => {
    setExtent({ rows, cols });
  }, []);

  /** What the sheet's filter is hiding, for the status bar to say. */
  const handleFiltered = useCallback(
    (next: { shown: number; total: number } | null) => setFiltered(next),
    [],
  );

  useEffect(() => {
    let unlisten: (() => void) | undefined;
    let disposed = false;

    getCurrentWebview()
      .onDragDropEvent((event) => {
        if (event.payload.type !== "drop") return;
        const paths = event.payload.paths.filter(looksLikeWorkbook);
        if (paths.length) void openPaths(paths);
      })
      .then((fn) => {
        if (disposed) fn();
        else unlisten = fn;
      })
      .catch(() => {});

    return () => {
      disposed = true;
      unlisten?.();
    };
  }, [openPaths]);

  // A silent check at startup, so a waiting release earns a dot on the toolbar.
  // Nothing is shown when there is none, when the build cannot replace itself,
  // or when the check fails — being offline is not worth interrupting anyone.
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const status = await updateStatus();
        if (cancelled) return;
        setUpdate(status);
        if (!status.portable) return;
        const found = await checkForUpdate(i18n.language);
        if (!cancelled) setUpdateAvailable(found !== null);
      } catch {
        /* nothing to report */
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  /**
   * The shortcuts' own handlers, read through a ref: they are rebuilt with the
   * selection they act on, and the window listener should not have to be taken
   * off and put back on every time the cursor moves.
   */
  const keysRef = useRef({ tabs, activeStyle, applyFormat, toggleFilter });
  keysRef.current = { tabs, activeStyle, applyFormat, toggleFilter };

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const { tabs, activeStyle, applyFormat, toggleFilter } = keysRef.current;
      if (!event.ctrlKey && !event.metaKey) {
        // Excel's F5 opens Go To. Without this the webview would reload the
        // page, which for a viewer means losing every open workbook's view.
        if (event.key === "F5") {
          event.preventDefault();
          if (activePath) setGotoOpen(true);
          return;
        }
        if (event.key === "Escape") {
          // Escape puts the panel away only from the panel's own row — the
          // search field is where it is asked for. Everywhere else the key is
          // the sheet's or a field's: it takes the marquee off a copy, closes
          // a cell's text, and none of that is meant to take the panel with
          // it. Ctrl+F is the shortcut that closes it from anywhere.
          const target = event.target;
          if (target instanceof Node && searchPanelRef.current?.contains(target)) {
            setSearchOpen(false);
          }
        }
        return;
      }
      const key = event.key.toLowerCase();
      // Ctrl+Tab walks the open workbooks, the way it walks Excel's windows;
      // the sheet's own Ctrl+PageUp/Down belong to the grid.
      if (key === "tab") {
        event.preventDefault();
        if (tabs.length < 2) return;
        const at = tabs.findIndex((tab) => tab.info.path === activePath);
        const forward = !event.shiftKey;
        const next =
          at < 0
            ? (forward ? tabs[0] : tabs[tabs.length - 1])
            : tabs[(at + (forward ? 1 : -1) + tabs.length) % tabs.length];
        if (next) selectTab(next.info.path);
        return;
      }
      // Shift makes a different shortcut rather than the same one louder.
      // Without this, Ctrl+Shift+G answered as Ctrl+G and opened the ordinary
      // Go To, and every other letter did the same thing. Only Ctrl+Shift+D has
      // a meaning of its own: the hidden readout, dev builds only, like the
      // panel itself.
      if (event.shiftKey) {
        if (key === "d" && import.meta.env.DEV) {
          event.preventDefault();
          toggleShapeDebug();
        } else if (key === "l") {
          // Excel's filter switch. It is the ribbon's own button: the key
          // belongs to the sheet, not the cursor.
          if (!activePath) return;
          event.preventDefault();
          void toggleFilter();
        }
        return;
      }
      switch (key) {
        case "o":
          event.preventDefault();
          void handleOpen();
          break;
        case "f":
          if (!activePath) return;
          event.preventDefault();
          setSearchOpen(true);
          setSearchFocus((n) => n + 1);
          break;
        case "g":
          if (!activePath) return;
          event.preventDefault();
          setGotoOpen(true);
          break;
        case "w":
          if (!activePath) return;
          event.preventDefault();
          closeTab(activePath);
          break;
        // Excel's typing keys, applied to what is selected. The formats are
        // the session's own — the ribbon writes the same ones. A field being
        // typed into keeps them: that text is not the sheet's.
        case "b":
        case "i":
        case "u": {
          if (!activePath) return;
          const target = event.target as HTMLElement | null;
          if (target?.closest("input, textarea, [contenteditable]")) return;
          event.preventDefault();
          const cell = activeStyle?.style ?? null;
          if (key === "b") applyFormat({ bold: !(cell?.bold ?? false) });
          else if (key === "i") applyFormat({ italic: !(cell?.italic ?? false) });
          else applyFormat({ underline: cell?.underline ? null : "single" });
          break;
        }
        default:
          break;
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [handleOpen, activePath, closeTab, selectTab]);

  const startPanelDrag = useCallback(
    (which: "sheetList" | "search", event: React.PointerEvent<HTMLDivElement>) => {
      panelDrag.current = {
        x: event.clientX,
        width: which === "sheetList" ? sheetListWidth : searchWidth,
      };
      setDraggingPanel(which);
    },
    [sheetListWidth, searchWidth],
  );

  /**
   * The sheet's scrollbar is a share of the row until it is dragged, so the drag
   * starts from whatever that measures now rather than from a number the row has
   * since outgrown.
   */
  const startSheetBarDrag = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const el = sheetBarRef.current;
    if (!el) return;
    panelDrag.current = {
      x: event.clientX,
      width: Math.round(el.getBoundingClientRect().width),
    };
    setDraggingPanel("sheetBar");
  }, []);

  /**
   * The widest the sheet's scrollbar can be before the tab strip beside it would
   * have to give up more room than it has — which is the only ceiling there is,
   * and so has to be measured rather than written down: the dots come and go,
   * and the floor the tabs keep is the stylesheet's business, not this file's.
   *
   * Asking for more than this is not merely pointless, it is what made the drag
   * go dead: the layout clamps while the number goes on climbing, so the handle
   * stops following the pointer and the first stretch of a drag back does
   * nothing at all.
   */
  const sheetBarRoom = useCallback(() => {
    const bar = sheetBarRef.current;
    const strip = bar
      ?.closest("[data-sheet-strip]")
      ?.querySelector("[data-sheet]")?.parentElement;
    // Neither can be missing while a drag is running — the handle that starts it
    // is in the same row — so the fallback is only here to keep the type honest.
    if (!bar || !strip) return Number.POSITIVE_INFINITY;
    const floor = parseFloat(getComputedStyle(strip).minWidth) || 0;
    const spare = strip.getBoundingClientRect().width - floor;
    return bar.getBoundingClientRect().width + Math.max(0, spare);
  }, []);

  /**
   * The listeners live on the window, not the handle: a drag leaves the four
   * pixels of it immediately, and pointer capture would keep the cursor from
   * reading as a resize once it is outside.
   */
  useEffect(() => {
    if (!draggingPanel) return;
    const [min, max] = PANEL_LIMITS[draggingPanel];
    const resize = {
      sheetList: setSheetListWidth,
      search: setSearchWidth,
      sheetBar: setSheetBarWidth,
    }[draggingPanel];
    // A panel on the left grows as the pointer moves right; the two on the
    // right grow as it moves left.
    const direction = draggingPanel === "sheetList" ? 1 : -1;

    const onMove = (event: PointerEvent) => {
      const drag = panelDrag.current;
      if (!drag) return;
      const delta = (event.clientX - drag.x) * direction;
      // The scrollbar's ceiling is whichever comes first: the width the row has
      // to give it, or the one configured for it.
      const ceiling =
        draggingPanel === "sheetBar" ? Math.min(max, sheetBarRoom()) : max;
      resize(Math.min(ceiling, Math.max(min, drag.width + delta)));
    };
    const onUp = () => setDraggingPanel(null);

    const { cursor, userSelect } = document.body.style;
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
    return () => {
      document.body.style.cursor = cursor;
      document.body.style.userSelect = userSelect;
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
    };
  }, [draggingPanel, sheetBarRoom]);

  return (
    // Nothing here is a tooltip you can hover into — they are one line of text
    // beside the control. Turning that off also lets the pointer walk straight
    // from one small icon to its neighbour: with it on, Radix holds the old
    // hint open and treats the pointer as still on its way to it, so the next
    // icon's hint never appears.
    <TooltipProvider disableHoverableContent>
      {/* Dragging anywhere — a block of cells, a gutter, the sheet list — must
          not smear a text selection across the chrome, so nothing is
          selectable by default and the few read-only strings opt back in. */}
      <div className="flex h-screen w-screen flex-col overflow-hidden bg-background text-foreground select-none">
        <Toolbar
          onOpen={handleOpen}
          hasWorkbook={active !== null}
          recentFiles={recentFiles}
          missingFiles={missingFiles}
          onOpenRecent={handleOpenRecent}
          onForgetRecent={handleForgetRecent}
          onClearRecents={handleClearRecents}
          onCloseWorkbook={handleCloseWorkbook}
          canUpdate={update?.portable ?? false}
          updateAvailable={updateAvailable}
          onShowUpdate={() => setUpdateOpen(true)}
        />

        <FileTabs
          tabs={tabs.map((tab) => ({
            path: tab.info.path,
            fileName: tab.info.fileName,
          }))}
          active={activePath}
          onSelect={selectTab}
          onClose={closeTab}
          onCloseMany={closeTabs}
          onReload={(path) => void reloadTab(path)}
          onReorder={handleReorderTabs}
          homeOpen={homeOpen}
          atHome={homeOpen && activePath === null}
          onHome={() => setActivePath(null)}
          onCloseHome={closeHome}
          onNew={handleNewTab}
        />

        {/* Under the file tabs, and only while one is showing: the start page
            has no sheet for any of this to act on. */}
        {active && activeSheet && (
          <Ribbon
            sidebarOpen={panels.list}
            onToggleSidebar={() => togglePanel("list")}
            sheetTabsOpen={panels.tabs}
            onToggleSheetTabs={() => togglePanel("tabs")}
            onMark={applyMark}
            markColor={markColor}
            activeStyle={activeStyle}
            onFormat={applyFormat}
            frozen={frozenBands.rows > 0 || frozenBands.cols > 0}
            onToggleFreeze={toggleFreeze}
            splitOn={splitOn}
            onToggleSplit={toggleSplit}
            filterOn={filterOn}
            onToggleFilter={() => void toggleFilter()}
            onGroup={applyOutline}
            focusCell={focusCell}
            onToggleFocusCell={() => setFocusCell((on) => !on)}
          />
        )}

        {error && active && (
          <div className="flex shrink-0 items-center justify-between gap-2 border-b border-border bg-destructive/10 px-3 py-1 text-[11px] text-destructive">
            <span className="truncate" title={error}>
              {error}
            </span>
            <button
              type="button"
              className="shrink-0 underline"
              onClick={() => setError(null)}
            >
              {t("app.dismiss")}
            </button>
          </div>
        )}

        <div className="flex min-h-0 flex-1">
          {active && activeSheet ? (
            <>
              {panels.list && (
                <>
                  <div
                    className="flex min-h-0 shrink-0 flex-col"
                    style={{ width: sheetListWidth }}
                  >
                    {/* Two panels share the sidebar; the grid is the same
                        either way, so the switch costs nothing. */}
                    <div className="flex shrink-0 border-b border-border bg-sidebar">
                      {(["sheets", "notes"] as const).map((tab) => (
                        <button
                          key={tab}
                          type="button"
                          onClick={() => setSidebarTab(tab)}
                          className={
                            sidebarTab === tab
                              ? "flex-1 border-b-2 border-[#107C41] px-2 py-1 text-[11px] font-medium text-foreground"
                              : "flex-1 border-b-2 border-transparent px-2 py-1 text-[11px] text-muted-foreground hover:text-foreground"
                          }
                        >
                          {tab === "sheets" ? t("sheets.list") : t("notes.title")}
                        </button>
                      ))}
                    </div>
                    <div className="min-h-0 flex-1">
                      {sidebarTab === "sheets" ? (
                        <SheetList
                          path={active.info.path}
                          sheets={active.info.sheets}
                          hidden={new Set(hiddenNames)}
                          active={active.sheet}
                          onSelect={handleSheetSelect}
                        />
                      ) : (
                        <NoteList
                          path={active.info.path}
                          sheet={active.sheet}
                          onGo={(row, col) => handleGo(active.sheet, row, col)}
                          onPreview={handleNotePreview}
                        />
                      )}
                    </div>
                  </div>
                  <Hint label={t("app.dragToResize")}>
                    <div
                      role="separator"
                      aria-orientation="vertical"
                      aria-label={t("app.resizeSheetList")}
                      onPointerDown={(event) => startPanelDrag("sheetList", event)}
                      className="w-1 shrink-0 cursor-col-resize transition-colors hover:bg-[#107C41]/40"
                    />
                  </Hint>
                </>
              )}

              {/* The strip is at the foot of the sheet area, as Excel's is, so
                  it stops where the search panel begins rather than running
                  under it. */}
              <div className="flex min-w-0 flex-1 flex-col">
                <div className="min-h-0 flex-1">
                  <CanvasGrid
                    path={active.info.path}
                    sheet={activeSheet}
                    revision={active.revision}
                    views={views.current}
                    selection={selection}
                    marks={active.marks}
                    formats={active.formats}
                    onActiveStyle={setActiveStyle}
                    frozen={active.frozen[active.sheet] ?? null}
                    onFrozenChange={setFrozenBands}
                    filterPane={active.filterPane[active.sheet] ?? null}
                    onFilterState={setFilterOn}
                    sessionOutline={active.sessionOutline[active.sheet] ?? null}
                    split={active.splitAt[active.sheet] ?? null}
                    onSplitChange={applySplit}
                    onMark={applyMark}
                    markColor={markColor}
                    onSelect={handleSelect}
                    onDetail={setDetail}
                    onExtent={handleExtent}
                    onFiltered={handleFiltered}
                    onSheetStep={handleSheetStep}
                    reveal={reveal}
                    previewNote={notePreview}
                    zoom={zoom}
                    onZoom={setZoom}
                    hideHScroll={panels.tabs || splitOn}
                    tabsShown={panels.tabs}
                    splitBarHost={panels.tabs && splitOn ? splitBarHost : null}
                    onSplitLine={setSplitLineX}
                    pageBreak={view === "pageBreak"}
                    focusCell={focusCell}
                    onViewport={setViewport}
                    bar={
                      <FormulaBar
                        reference={selectionRef(selection)}
                        value={detail?.text ?? ""}
                        onGo={handleNameBoxGo}
                      />
                    }
                  />
                </div>
                {panels.tabs && (
                  <SheetTabs
                    sheets={shown}
                    active={active.sheet}
                    onSelect={handleSheetSelect}
                    onTabClick={handleSheetClick}
                    selected={sheetSelection}
                    colors={active.tabColors}
                    onColor={applyTabColor}
                    onHide={handleHideSheet}
                    onUnhide={(at) => {
                      setUnhideAt(at);
                      setUnhideOpen(true);
                    }}
                    onReorder={handleReorderSheets}
                    hiddenCount={hiddenNames.length}
                    splitX={splitOn ? splitLineX : null}
                    trailing={
                      // A split with two column bands draws its own pair here,
                      // segmented at the line. With one band — a split across
                      // rows only — the row stays exactly as it is: splitter,
                      // share and all. Only what drives the bar changes, from
                      // the view to the band, so choosing a row split does not
                      // move the scrollbar.
                      splitOn && splitLineX != null ? (
                        <div
                          ref={setSplitBarHost}
                          className="relative min-w-0 flex-1 self-stretch"
                        />
                      ) : (
                      <>
                        {/* Excel's splitter, at the left end of the sheet's
                            scrollbar: drag it and the scrollbar is longer or
                            shorter, the tabs taking what is left. Excel draws
                            it as a column of dots, so it can be found at all. */}
                        <Hint label={t("app.dragToResize")}>
                          <div
                            role="separator"
                            aria-orientation="vertical"
                            aria-label={t("app.resizeSheetScrollBar")}
                            onPointerDown={startSheetBarDrag}
                            className="flex w-4 shrink-0 cursor-col-resize items-center justify-center text-muted-foreground/60 transition-colors hover:bg-[#107C41]/40 hover:text-foreground"
                          >
                            <EllipsisVertical className="size-3" />
                          </div>
                        </Hint>
                        {/* A share of the row until it is dragged, so the length
                            is the row's rather than a number chosen here. The
                            bar is centred in the row; only the engine knows how
                            tall it is, so `SheetScrollBar` measures that. */}
                        <div
                          ref={setSheetBarHost}
                          className={`relative flex min-w-24 shrink items-center ${sheetBarWidth === null ? "w-1/3" : ""}`}
                          style={sheetBarWidth === null ? undefined : { width: sheetBarWidth }}
                        >
                          {splitOn ? null : <SheetScrollBar viewport={viewport} />}
                        </div>
                      </>
                      )
                    }
                  />
                )}
              </div>

              {searchOpen && (
                <>
                  {/* Mirrors the sheet list's handle, on the other side. */}
                  <Hint label={t("app.dragToResize")}>
                    <div
                      role="separator"
                      aria-orientation="vertical"
                      aria-label={t("app.resizeSearch")}
                      onPointerDown={(event) => startPanelDrag("search", event)}
                      className="w-1 shrink-0 cursor-col-resize transition-colors hover:bg-[#107C41]/40"
                    />
                  </Hint>
                  <div ref={searchPanelRef} className="shrink-0" style={{ width: searchWidth }}>
                    <SearchPanel
                      path={active.info.path}
                      sheet={active.sheet}
                      focusSignal={searchFocus}
                      onJump={handleJump}
                      onClose={() => setSearchOpen(false)}
                    />
                  </div>
                </>
              )}
            </>
          ) : active ? (
            <p className="m-auto text-xs text-muted-foreground">
              {t("app.noSheets")}
            </p>
          ) : (
            <Welcome
              onOpen={handleOpen}
              onOpenDirectory={handleOpenInDirectory}
              onForgetFolder={handleForgetFolder}
              recentFiles={recentFiles}
              missingFiles={missingFiles}
              onOpenRecent={handleOpenRecent}
              error={error}
            />
          )}
        </div>

        <StatusBar
          fileName={active?.info.fileName ?? null}
          sheetCount={active?.info.sheets.length ?? null}
          sheet={active?.sheet ?? null}
          size={
            extent
              ? t("app.extent", { rows: extent.rows, cols: extent.cols })
              : null
          }
          span={activeSheet ? selectionSize(selection) : null}
          filtered={active ? filtered : null}
          view={view}
          onView={handleView}
          zoom={zoom}
          onZoom={setZoom}
        />

        <GoToDialog
          open={gotoOpen}
          onOpenChange={setGotoOpen}
          sheets={shown.map((sheet) => sheet.name)}
          currentSheet={active?.sheet ?? ""}
          onGo={handleGo}
        />

        <UnhideDialog
          open={unhideOpen}
          onOpenChange={setUnhideOpen}
          sheets={hiddenNames}
          at={unhideAt}
          onUnhide={handleUnhideSheets}
        />

        <UpdateDialog
          open={updateOpen}
          onOpenChange={setUpdateOpen}
          status={update}
          onAvailableChange={setUpdateAvailable}
        />

        <Toaster />
      </div>
    </TooltipProvider>
  );
}
