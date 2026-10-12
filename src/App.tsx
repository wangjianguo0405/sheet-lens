import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { open as openFileDialog } from "@tauri-apps/plugin-dialog";
import { useTranslation } from "react-i18next";

import { CanvasGrid, type SheetViews } from "@/components/CanvasGrid";
import { FileTabs } from "@/components/FileTabs";
import { FormulaBar } from "@/components/FormulaBar";
import { GoToDialog } from "@/components/GoToDialog";
import { SearchPanel } from "@/components/SearchPanel";
import { SheetList } from "@/components/SheetList";
import { StatusBar } from "@/components/StatusBar";
import { Toolbar } from "@/components/Toolbar";
import { UpdateDialog } from "@/components/UpdateDialog";
import { Welcome } from "@/components/Welcome";
import { TooltipProvider } from "@/components/ui/tooltip";
import { ORIGIN, selectionRef, type Selection } from "@/lib/selection";
import { toggleShapeDebug } from "@/lib/debugPanels";
import { checkForUpdate, updateStatus, type UpdateStatus } from "@/lib/update";
import { DEFAULT_ZOOM } from "@/lib/zoom";
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
}

const EXCEL_EXTENSIONS = ["xlsx", "xlsm", "xls", "xlsb", "ods"];

const DEFAULT_SHEET_LIST_W = 240;
const MIN_SHEET_LIST_W = 160;
const MAX_SHEET_LIST_W = 480;

const DEFAULT_SEARCH_W = 288;
const MIN_SEARCH_W = 200;
const MAX_SEARCH_W = 640;

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
  const [sheetListOpen, setSheetListOpen] = useState(true);
  /** View scale as a percentage, the way Excel names it. */
  const [zoom, setZoom] = useState(DEFAULT_ZOOM);
  const [sheetListWidth, setSheetListWidth] = useState(DEFAULT_SHEET_LIST_W);
  const [searchWidth, setSearchWidth] = useState(DEFAULT_SEARCH_W);
  /**
   * Which side panel a drag is sizing, so one set of listeners serves both —
   * the two differ only in which way the pointer's delta counts.
   */
  const [draggingPanel, setDraggingPanel] = useState<"sheetList" | "search" | null>(null);
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
                },
              ],
        );
        setActivePath(info.path);
        lastFileRef.current = info.path;
        // Opening a workbook closes the start page's tab: it was somewhere to
        // start from, not somewhere to come back to.
        setHomeOpen(false);
      } catch (reason: unknown) {
        setError(`${path} — ${String(reason)}`);
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

  const handleSheetSelect = useCallback(
    (name: string) => updateActive({ sheet: name }),
    [updateActive],
  );

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
        if (event.key === "Escape") setSearchOpen(false);
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
   * The listeners live on the window, not the handle: a drag leaves the four
   * pixels of it immediately, and pointer capture would keep the cursor from
   * reading as a resize once it is outside.
   */
  useEffect(() => {
    if (!draggingPanel) return;
    const isLeft = draggingPanel === "sheetList";
    const clamp = (next: number) =>
      isLeft
        ? Math.min(MAX_SHEET_LIST_W, Math.max(MIN_SHEET_LIST_W, next))
        : Math.min(MAX_SEARCH_W, Math.max(MIN_SEARCH_W, next));

    const onMove = (event: PointerEvent) => {
      const drag = panelDrag.current;
      if (!drag) return;
      // The left panel grows as the pointer moves right, the right one as it
      // moves left.
      const delta = (event.clientX - drag.x) * (isLeft ? 1 : -1);
      const next = clamp(drag.width + delta);
      if (isLeft) setSheetListWidth(next);
      else setSearchWidth(next);
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
  }, [draggingPanel]);

  return (
    <TooltipProvider>
      {/* Dragging anywhere — a block of cells, a gutter, the sheet list — must
          not smear a text selection across the chrome, so nothing is
          selectable by default and the few read-only strings opt back in. */}
      <div className="flex h-screen w-screen flex-col overflow-hidden bg-background text-foreground select-none">
        <Toolbar
          onOpen={handleOpen}
          onToggleSearch={() => setSearchOpen((open) => !open)}
          onToggleSidebar={() => setSheetListOpen((open) => !open)}
          onGoTo={() => setGotoOpen(true)}
          searchOpen={searchOpen}
          sidebarOpen={sheetListOpen}
          hasWorkbook={active !== null}
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
              {sheetListOpen && (
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
                    className="w-1 shrink-0 cursor-col-resize transition-colors hover:bg-[#217346]/40"
                  />
                </>
              )}

              <div className="min-w-0 flex-1">
                <CanvasGrid
                  path={active.info.path}
                  sheet={activeSheet}
                  views={views.current}
                  selection={selection}
                  onSelect={handleSelect}
                  onDetail={setDetail}
                  onExtent={handleExtent}
                  onSheetStep={handleSheetStep}
                  reveal={reveal}
                  zoom={zoom}
                  onZoom={setZoom}
                  bar={
                    <FormulaBar
                      reference={selectionRef(selection)}
                      value={detail?.text ?? ""}
                    />
                  }
                />
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
                    className="w-1 shrink-0 cursor-col-resize transition-colors hover:bg-[#217346]/40"
                  />
                  <div className="shrink-0" style={{ width: searchWidth }}>
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
            <Welcome onOpen={handleOpen} error={error} />
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
          cell={activeSheet ? selectionRef(selection) : null}
          kind={detail?.kind ?? null}
          value={detail?.text ?? null}
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
