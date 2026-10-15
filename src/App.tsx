import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { open as openFileDialog } from "@tauri-apps/plugin-dialog";
import { EllipsisVertical } from "lucide-react";
import { useTranslation } from "react-i18next";

import { CanvasGrid, type SheetViews } from "@/components/CanvasGrid";
import { FileTabs } from "@/components/FileTabs";
import { FormulaBar } from "@/components/FormulaBar";
import { GoToDialog } from "@/components/GoToDialog";
import { SearchPanel } from "@/components/SearchPanel";
import { SheetList } from "@/components/SheetList";
import { SheetScrollBar } from "@/components/SheetScrollBar";
import { SheetTabs } from "@/components/SheetTabs";
import { StatusBar } from "@/components/StatusBar";
import { Toolbar } from "@/components/Toolbar";
import { UpdateDialog } from "@/components/UpdateDialog";
import { Welcome } from "@/components/Welcome";
import { TooltipProvider } from "@/components/ui/tooltip";
import {
  ORIGIN,
  selectionRef,
  selectionSize,
  type Selection,
} from "@/lib/selection";
import { toggleShapeDebug } from "@/lib/debugPanels";
import { initialMarkColor, markSelection, storeMarkColor } from "@/lib/marks";
import type { Target } from "@/lib/column";
import {
  initialSheetPanels,
  storeSheetPanels,
  type SheetPanels,
} from "@/lib/sheetPanels";
import {
  forgetAllFiles,
  forgetFile,
  markFileMissing,
  readMissingFiles,
  readRecentFiles,
  rememberFile,
} from "@/lib/recent";
import { checkForUpdate, updateStatus, type UpdateStatus } from "@/lib/update";
import { DEFAULT_ZOOM } from "@/lib/zoom";
import { initialView, type ViewMode } from "@/lib/viewMode";
import {
  closeWorkbook,
  openWorkbook,
  type Cell,
  type SearchHit,
  type WorkbookInfo,
} from "@/lib/excel";

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
   * The colours the reader gave the tabs, by sheet name. Excel writes a tab
   * colour into the file; this app has nowhere to write it, so the colour is its
   * own the way a mark is — it lives as long as the tab holding it, and a tab
   * left out of the map is drawn in whatever colour the workbook set.
   */
  tabColors: Map<string, string>;
  /**
   * The view every sheet of this workbook opens in — the page view when the
   * file sets a print area anywhere, since a workbook that is printed is read
   * with an eye on the page. Only a default: each sheet is then switched on its
   * own, and a switch lasts as long as the tab.
   */
  defaultView: ViewMode;
  /** The sheets switched away from that default, by name. */
  views: Record<string, ViewMode>;
}

const EXCEL_EXTENSIONS = ["xlsx", "xlsm", "xls", "xlsb", "ods"];

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
  const [searchOpen, setSearchOpen] = useState(false);
  const [gotoOpen, setGotoOpen] = useState(false);
  const [detail, setDetail] = useState<Cell | null>(null);
  /** What the grid actually shows, which can exceed calamine's range. */
  const [extent, setExtent] = useState<{ rows: number; cols: number } | null>(
    null,
  );
  const [error, setError] = useState<string | null>(null);
  /** Which of the two sheet pickers are showing; they are independent. */
  const [panels, setPanels] = useState<SheetPanels>(initialSheetPanels);
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
  /** View scale as a percentage, the way Excel names it. */
  const [zoom, setZoom] = useState(DEFAULT_ZOOM);
  const [sheetListWidth, setSheetListWidth] = useState(DEFAULT_SHEET_LIST_W);
  const [searchWidth, setSearchWidth] = useState(DEFAULT_SEARCH_W);
  const [sheetBarWidth, setSheetBarWidth] = useState<number | null>(null);
  /** The row's own box, so a drag can start from the length it has now. */
  const sheetBarRef = useRef<HTMLDivElement>(null);
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

  const selection = active?.cursors[active.sheet] ?? ORIGIN;

  /** What the sheet on screen is drawn in: its own switch, or the default. */
  const view: ViewMode = active
    ? (active.views[active.sheet] ?? active.defaultView)
    : "normal";

  const openPaths = useCallback(async (paths: string[]) => {
    setError(null);
    for (const path of paths) {
      try {
        const info = await openWorkbook(path);
        setTabs((prev) =>
          prev.some((tab) => tab.info.path === info.path)
            ? prev
            : [
                ...prev,
                {
                  info,
                  sheet: info.sheets[0]?.name ?? "",
                  cursors: {},
                  marks: new Map(),
                  tabColors: new Map(),
                  defaultView: initialView(info),
                  views: {},
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

  const handleOpen = useCallback(async () => {
    try {
      const picked = await openFileDialog({
        multiple: true,
        filters: [{ name: "Excel", extensions: EXCEL_EXTENSIONS }],
      });
      if (!picked) return;
      await openPaths(Array.isArray(picked) ? picked : [picked]);
    } catch (reason: unknown) {
      setError(String(reason));
    }
  }, [openPaths]);

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

  const handleSheetSelect = useCallback(
    (name: string) => updateActive({ sheet: name }),
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
    storeMarkColor(markColor);
  }, [markColor]);

  /**
   * Ctrl+PageUp/PageDown steps along the workbook's sheets in tab order, and
   * stops at either end rather than wrapping — which is what Excel does.
   */
  const handleSheetStep = useCallback(
    (delta: number) => {
      if (!active) return;
      const names = active.info.sheets.map((sheet) => sheet.name);
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

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!event.ctrlKey && !event.metaKey) {
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
      // Shift makes a different shortcut rather than the same one louder.
      // Without this, Ctrl+Shift+G answered as Ctrl+G and opened the ordinary
      // Go To, and every other letter did the same thing. Only Ctrl+Shift+D has
      // a meaning of its own: the hidden readout, dev builds only, like the
      // panel itself.
      if (event.shiftKey) {
        if (key === "d" && import.meta.env.DEV) {
          event.preventDefault();
          toggleShapeDebug();
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
          setSearchOpen((open) => !open);
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
        default:
          break;
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [handleOpen, activePath, closeTab]);

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
    <TooltipProvider>
      {/* Dragging anywhere — a block of cells, a gutter, the sheet list — must
          not smear a text selection across the chrome, so nothing is
          selectable by default and the few read-only strings opt back in. */}
      <div className="flex h-screen w-screen flex-col overflow-hidden bg-background text-foreground select-none">
        <Toolbar
          onOpen={handleOpen}
          onToggleSidebar={() => togglePanel("list")}
          onToggleSheetTabs={() => togglePanel("tabs")}
          sidebarOpen={panels.list}
          sheetTabsOpen={panels.tabs}
          hasWorkbook={active !== null}
          onMark={applyMark}
          markColor={markColor}
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
          onReorder={handleReorderTabs}
          homeOpen={homeOpen}
          atHome={homeOpen && activePath === null}
          onHome={() => setActivePath(null)}
          onCloseHome={closeHome}
          onNew={handleNewTab}
        />

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
                  <div className="shrink-0" style={{ width: sheetListWidth }}>
                    <SheetList
                      sheets={active.info.sheets}
                      active={active.sheet}
                      onSelect={handleSheetSelect}
                    />
                  </div>
                  <div
                    role="separator"
                    aria-orientation="vertical"
                    aria-label={t("app.resizeSheetList")}
                    title={t("app.dragToResize")}
                    onPointerDown={(event) => startPanelDrag("sheetList", event)}
                    className="w-1 shrink-0 cursor-col-resize transition-colors hover:bg-[#107C41]/40"
                  />
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
                    views={views.current}
                    selection={selection}
                    marks={active.marks}
                    onMark={applyMark}
                    markColor={markColor}
                    onSelect={handleSelect}
                    onDetail={setDetail}
                    onExtent={handleExtent}
                    onSheetStep={handleSheetStep}
                    reveal={reveal}
                    zoom={zoom}
                    onZoom={setZoom}
                    hideHScroll={panels.tabs}
                    pageBreak={view === "pageBreak"}
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
                    sheets={active.info.sheets}
                    active={active.sheet}
                    onSelect={handleSheetSelect}
                    colors={active.tabColors}
                    onColor={applyTabColor}
                    trailing={
                      <>
                        {/* Excel's splitter, at the left end of the sheet's
                            scrollbar: drag it and the scrollbar is longer or
                            shorter, the tabs taking what is left. Excel draws
                            it as a column of dots, so it can be found at all. */}
                        <div
                          role="separator"
                          aria-orientation="vertical"
                          aria-label={t("app.resizeSheetScrollBar")}
                          title={t("app.dragToResize")}
                          onPointerDown={startSheetBarDrag}
                          className="flex w-4 shrink-0 cursor-col-resize items-center justify-center text-muted-foreground/60 transition-colors hover:bg-[#107C41]/40 hover:text-foreground"
                        >
                          <EllipsisVertical className="size-3" />
                        </div>
                        {/* A share of the row until it is dragged, so the length
                            is the row's rather than a number chosen here. The
                            bar is centred in the row; only the engine knows how
                            tall it is, so `SheetScrollBar` measures that. */}
                        <div
                          ref={sheetBarRef}
                          className={`flex min-w-24 shrink items-center ${sheetBarWidth === null ? "w-1/3" : ""}`}
                          style={sheetBarWidth === null ? undefined : { width: sheetBarWidth }}
                        >
                          <SheetScrollBar viewport={viewport} />
                        </div>
                      </>
                    }
                  />
                )}
              </div>

              {searchOpen && (
                <>
                  {/* Mirrors the sheet list's handle, on the other side. */}
                  <div
                    role="separator"
                    aria-orientation="vertical"
                    aria-label={t("app.resizeSearch")}
                    title={t("app.dragToResize")}
                    onPointerDown={(event) => startPanelDrag("search", event)}
                    className="w-1 shrink-0 cursor-col-resize transition-colors hover:bg-[#107C41]/40"
                  />
                  <div ref={searchPanelRef} className="shrink-0" style={{ width: searchWidth }}>
                    <SearchPanel
                      path={active.info.path}
                      sheet={active.sheet}
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
              recentFiles={recentFiles}
              missingFiles={missingFiles}
              onOpenRecent={handleOpenRecent}
              error={error}
            />
          )}
        </div>

        <StatusBar
          fileName={active?.info.fileName ?? null}
          sheet={active?.sheet ?? null}
          size={
            extent
              ? t("app.extent", { rows: extent.rows, cols: extent.cols })
              : null
          }
          span={activeSheet ? selectionSize(selection) : null}
          view={view}
          onView={handleView}
          zoom={zoom}
          onZoom={setZoom}
        />

        <GoToDialog
          open={gotoOpen}
          onOpenChange={setGotoOpen}
          sheets={active?.info.sheets.map((sheet) => sheet.name) ?? []}
          currentSheet={active?.sheet ?? ""}
          onGo={handleGo}
        />

        <UpdateDialog
          open={updateOpen}
          onOpenChange={setUpdateOpen}
          status={update}
          onAvailableChange={setUpdateAvailable}
        />
      </div>
    </TooltipProvider>
  );
}
