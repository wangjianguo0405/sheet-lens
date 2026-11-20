import { useCallback, useEffect, useRef, useState } from "react";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { EllipsisVertical } from "lucide-react";
import { useTranslation } from "react-i18next";

import { CanvasGrid, type SheetViews } from "@/components/CanvasGrid";
import { AppDialogs } from "@/components/AppDialogs";
import { FileTabs } from "@/components/FileTabs";
import { FileTree } from "@/components/FileTree";
import { Hint } from "@/components/Hint";
import { FormulaBar } from "@/components/FormulaBar";
import { Ribbon } from "@/components/Ribbon";
import { SearchPanel } from "@/components/SearchPanel";
import { SheetSidebar } from "@/components/SheetSidebar";
import { SheetScrollBar } from "@/components/SheetScrollBar";
import { SheetTabs } from "@/components/SheetTabs";
import { StatusBar } from "@/components/StatusBar";
import { Toaster } from "@/components/ui/sonner";
import { Toolbar } from "@/components/Toolbar";
import { Welcome } from "@/components/Welcome";
import { TooltipProvider } from "@/components/ui/tooltip";
import { ORIGIN, selectionRef, selectionSize } from "@/lib/selection";
import { usePanels } from "@/hooks/usePanels";
import { useSearchPanel } from "@/hooks/useSearchPanel";
import { useGridNavigation } from "@/hooks/useGridNavigation";
import { useUpdateCheck } from "@/hooks/useUpdateCheck";
import { useGridReports } from "@/hooks/useGridReports";
import { useSessionEdits } from "@/hooks/useSessionEdits";
import { DEFAULT_ZOOM } from "@/lib/zoom";
import type { ViewMode } from "@/lib/viewMode";
import { useUndo, type SessionSnapshot, type UndoTabState } from "@/hooks/useUndo";
import { useShortcuts } from "@/hooks/useShortcuts";
import { useSheets } from "@/hooks/useSheets";
import { looksLikeWorkbook, useWorkbooks } from "@/hooks/useWorkbooks";

export default function App() {
  const { t, i18n } = useTranslation();
  const {
    panels,
    togglePanel,
    focusCell,
    setFocusCell,
    treeShown,
    setTreeShown,
    welcomeColumns,
    setWelcomeColumns,
    sheetListWidth,
    treeWidth,
    searchWidth,
    sheetBarWidth,
    sheetBarRef,
    startPanelDrag,
    startSheetBarDrag,
  } = usePanels();
  const { update, updateAvailable, setUpdateAvailable, updateOpen, setUpdateOpen } = useUpdateCheck(
    i18n.language,
  );
  const {
    searchOpen,
    setSearchOpen,
    searchSeed,
    setSearchSeed,
    searchFocus,
    setSearchFocus,
    searchFor,
    searchPanelRef,
  } = useSearchPanel();
  const [gotoOpen, setGotoOpen] = useState(false);
  const [error, setError] = useState<string | null>(null);
  /**
   * Where each sheet of each open workbook was left scrolled. It is the tabs
   * that own this — a clue about a workbook nobody has open is not worth
   * keeping — so `closeTab` drops the whole entry, the way the tab's cursors go
   * with it.
   */
  const views = useRef<SheetViews>(new Map());
  /**
   * How a closed tab tells the undo history to forget it. A ref because the
   * history is built on the workbooks' own `selectTab`, so the two are wired in
   * that order.
   */
  const forgetHistoryRef = useRef<(path: string) => void>(() => {});
  const {
    tabs,
    setTabs,
    activePath,
    setActivePath,
    active,
    activeSheet,
    homeOpen,
    recentFiles,
    missingFiles,
    openPaths,
    handleOpen,
    handleOpenInDirectory,
    handleOpenRecent,
    handleForgetRecent,
    handleClearRecents,
    handleForgetFolder,
    handleNewTab,
    closeHome,
    selectTab,
    handleReorderTabs,
    updateActive,
    closeTab,
    closeTabs,
    reloadTab,
    handleCloseWorkbook,
  } = useWorkbooks({ setError, views, forgetHistoryRef });
  /**
   * The grid puts its "colour the field's selection, else the cell" function
   * here, so the ribbon's button — which is not inside the grid — can call it.
   * Null while no grid is mounted.
   */
  const textColorVia = useRef<((color: string | null) => void) | null>(null);
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
  /** Where a cell's right-click put the mini toolbar, or null when it is away. */
  const [miniToolbar, setMiniToolbar] = useState<{ x: number; y: number } | null>(null);
  const closeMiniToolbar = useCallback(() => setMiniToolbar(null), []);
  /** The note the sidebar is pointing at, to be shown in the sheet. */
  const [notePreview, setNotePreview] = useState<{ row: number; col: number } | null>(null);
  const handleNotePreview = useCallback(
    (note: { row: number; col: number } | null) => setNotePreview(note),
    [],
  );
  /**
   * The same box, for a split's own horizontal bar to be drawn in: a split
   * across rows only has one column band, so its bar lives where the sheet's
   * ordinary one does and the splitter keeps working.
   */
  const setSheetBarHost = useCallback(
    (element: HTMLDivElement | null) => {
      sheetBarRef.current = element;
      setSplitBarHost(element);
    },
    [sheetBarRef],
  );
  /** The search panel's box, so Escape can tell whose key it was. */
  /**
   * Bumped when the cursor is placed somewhere the grid has to show, rather
   * than somewhere the user is walking to. Switching sheets is not that: the
   * grid puts the view back where each sheet was left, which is a position the
   * cursor does not get to overrule.
   */
  const [reveal, setReveal] = useState(0);
  // The toolbar belongs to the cell it was opened on; a sheet or workbook
  // switch leaves it pointing at nothing.
  useEffect(() => setMiniToolbar(null), [active?.info.path, activeSheet]);

  const selection = active?.cursors[active.sheet] ?? ORIGIN;

  /** What the sheet on screen is drawn in: its own switch, or the default. */
  const view: ViewMode = active ? (active.views[active.sheet] ?? active.defaultView) : "normal";

  /** Whether the sheet on screen is split, which brings its own scrollbars. */
  const splitOn = Boolean(active?.splitAt[active.sheet]);

  /**
   * The tabs as they stand, for the undo history: it takes a step's snapshot
   * between renders, when the latest `tabs` a callback closed over may already
   * be a render old.
   */
  const tabsRef = useRef(tabs);
  tabsRef.current = tabs;

  /** What an undo step reads and writes on a tab; the tab's shape stays here. */
  const readUndoState = useCallback((path: string): UndoTabState | null => {
    const tab = tabsRef.current.find((one) => one.info.path === path);
    if (!tab) return null;
    return {
      marks: tab.marks,
      formats: tab.formats,
      textColors: tab.textColors,
      tabColors: tab.tabColors,
      sheetHidden: tab.sheetHidden,
    };
  }, []);

  const restoreUndoState = useCallback(
    (step: SessionSnapshot) => {
      setTabs((prev) =>
        prev.map((tab) =>
          tab.info.path === step.path
            ? {
                ...tab,
                sheet: step.sheet,
                marks: step.marks,
                formats: step.formats,
                textColors: step.textColors,
                tabColors: step.tabColors,
                sheetHidden: step.sheetHidden,
                // The grid re-seeds the sheet's sizes, hidden lines and shape
                // offsets from the view store on a revision bump; this is what
                // draws the snapshot rather than the state it replaced.
                revision: tab.revision + 1,
              }
            : tab,
        ),
      );
    },
    [setTabs],
  );

  const { recordEdit, undo, redo, forgetHistory, canUndo, canRedo } = useUndo({
    views,
    readTab: readUndoState,
    selectTab,
    restoreTab: restoreUndoState,
  });
  // From here on, a closed tab takes its undo steps with it.
  forgetHistoryRef.current = forgetHistory;

  /** A grid-side session edit, recorded against the sheet showing. */
  const recordGridEdit = useCallback(() => {
    if (!active) return;
    recordEdit(active.info.path, active.sheet);
  }, [active, recordEdit]);

  const {
    sheetSelection,
    shown,
    hiddenNames,
    unhideOpen,
    setUnhideOpen,
    unhideAt,
    setUnhideAt,
    handleSheetClick,
    handleHideSheet,
    applySheetHidden,
    handleUnhideSheets,
    handleSheetSelect,
    handleSheetStep,
    handleReorderSheets,
  } = useSheets({ active, activePath, setTabs, updateActive, recordEdit });

  const {
    detail,
    setDetail,
    activeStyle,
    setActiveStyle,
    frozenBands,
    setFrozenBands,
    filterOn,
    setFilterOn,
    extent,
    setExtent,
    filtered,
    setFiltered,
  } = useGridReports();

  const {
    markColor,
    fontColor,
    setFontColor,
    applyMark,
    applyFormat,
    applyTextColor,
    applyTabColor,
    applySplit,
    toggleSplit,
    toggleFreeze,
    toggleFilter,
    applyOutline,
  } = useSessionEdits({
    active,
    activeSheet,
    extent,
    frozenBands,
    setFrozenBands,
    filterOn,
    setFilterOn,
    setError,
    setTabs,
    recordEdit,
  });

  const { handleSelect, handleView, handleJump, handleGo, handleNameBoxGo, handleExtent } =
    useGridNavigation({
      active,
      activePath,
      setTabs,
      updateActive,
      setReveal,
      viewport,
      setExtent,
    });

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

  useShortcuts({
    tabs,
    activePath,
    activeStyle,
    applyFormat,
    toggleFilter,
    handleOpen,
    closeTab,
    selectTab,
    reloadTab,
    undo,
    redo,
    setGotoOpen,
    setSearchOpen,
    setSearchSeed,
    setSearchFocus,
    searchPanelRef,
  });

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
          canUndo={canUndo}
          canRedo={canRedo}
          onUndo={undo}
          onRedo={redo}
          fileTreeOpen={treeShown}
          onToggleFileTree={() => setTreeShown((shown) => !shown)}
          welcomeDirsOpen={welcomeColumns.folders}
          onToggleWelcomeDirs={() =>
            setWelcomeColumns((columns) => ({ ...columns, folders: !columns.folders }))
          }
          welcomeFilesOpen={welcomeColumns.files}
          onToggleWelcomeFiles={() =>
            setWelcomeColumns((columns) => ({ ...columns, files: !columns.files }))
          }
          sidebarOpen={panels.list}
          onToggleSidebar={() => togglePanel("list")}
          sheetTabsOpen={panels.tabs}
          onToggleSheetTabs={() => togglePanel("tabs")}
          canUpdate={update?.portable ?? false}
          updateAvailable={updateAvailable}
          onShowUpdate={() => setUpdateOpen(true)}
        />

        <div className="flex min-h-0 flex-1">
          {/* The reader's own folders, beside the workbook rather than on the
              start page only: switching to another checkout's document is a
              thing done while reading, not before it. The column runs from
              the top bar to the status bar — the file tabs begin to its
              right, so nothing crosses it. */}
          {active && activeSheet && treeShown && (
            <>
              <div
                className="flex min-h-0 shrink-0 flex-col border-r border-border bg-sidebar"
                style={{ width: treeWidth }}
              >
                <FileTree currentPath={active.info.path} onOpenFile={handleOpenRecent} />
              </div>
              <Hint label={t("app.dragToResize")}>
                <div
                  role="separator"
                  aria-orientation="vertical"
                  aria-label={t("app.resizeFileTree")}
                  onPointerDown={(event) => startPanelDrag("fileTree", event)}
                  className="w-1 shrink-0 cursor-col-resize transition-colors hover:bg-[#107C41]/40"
                />
              </Hint>
            </>
          )}

          {/* The tabs and everything under them share the column beside the
              tree. The ribbon crosses the sheet list and the grid as it always
              has; the start page has no workbook for any of it. */}
          <div className="flex min-w-0 flex-1 flex-col">
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

            {active && activeSheet ? (
              <>
                <Ribbon
                  sidebarOpen={panels.list}
                  onToggleSidebar={() => togglePanel("list")}
                  fileTreeOpen={treeShown}
                  onToggleFileTree={() => setTreeShown((shown) => !shown)}
                  sheetTabsOpen={panels.tabs}
                  onToggleSheetTabs={() => togglePanel("tabs")}
                  onMark={applyMark}
                  markColor={markColor}
                  fontColor={fontColor}
                  onFontColor={(color) => {
                    if (color !== null) setFontColor(color);
                    // The grid decides whether this lands on the whole selection
                    // or on the stretch of the open cell's text the reader selected.
                    textColorVia.current?.(color);
                  }}
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
                  miniToolbar={miniToolbar}
                  onCloseMiniToolbar={closeMiniToolbar}
                />

                {error && (
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
                  {panels.list && (
                    <SheetSidebar
                      path={active.info.path}
                      sheets={active.info.sheets}
                      hidden={hiddenNames}
                      active={active.sheet}
                      colors={active.tabColors}
                      width={sheetListWidth}
                      onSelect={handleSheetSelect}
                      onColor={applyTabColor}
                      onHide={(name) => applySheetHidden(name, true)}
                      onUnhide={(at) => {
                        setUnhideAt(at);
                        setUnhideOpen(true);
                      }}
                      onGo={(row, col) => handleGo(active.sheet, row, col)}
                      onPreview={handleNotePreview}
                      onDrag={(event) => startPanelDrag("sheetList", event)}
                    />
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
                        textColors={active.textColors}
                        textColorVia={textColorVia}
                        onTextColor={applyTextColor}
                        onFormatColor={(color) => applyFormat({ color })}
                        onActiveStyle={setActiveStyle}
                        frozen={active.frozen[active.sheet] ?? null}
                        onFrozenChange={setFrozenBands}
                        filterPane={active.filterPane[active.sheet] ?? null}
                        onFilterState={setFilterOn}
                        sessionOutline={active.sessionOutline[active.sheet] ?? null}
                        split={active.splitAt[active.sheet] ?? null}
                        onSplitChange={applySplit}
                        onMiniToolbar={setMiniToolbar}
                        onPreviewNote={handleNotePreview}
                        onSearchValue={searchFor}
                        onSelect={handleSelect}
                        onEdit={recordGridEdit}
                        onDetail={setDetail}
                        onExtent={handleExtent}
                        onFiltered={setFiltered}
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
                                style={
                                  sheetBarWidth === null ? undefined : { width: sheetBarWidth }
                                }
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
                          seed={searchSeed}
                          focusSignal={searchFocus}
                          onJump={handleJump}
                          onClose={() => {
                            setSearchOpen(false);
                            setSearchSeed(null);
                          }}
                        />
                      </div>
                    </>
                  )}
                </div>
              </>
            ) : active ? (
              <div className="flex min-h-0 flex-1 items-center justify-center">
                <p className="text-xs text-muted-foreground">{t("app.noSheets")}</p>
              </div>
            ) : (
              <div className="min-h-0 flex-1">
                <Welcome
                  onOpen={handleOpen}
                  onOpenDirectory={handleOpenInDirectory}
                  onForgetFolder={handleForgetFolder}
                  onForgetFile={handleForgetRecent}
                  recentFiles={recentFiles}
                  missingFiles={missingFiles}
                  onOpenRecent={handleOpenRecent}
                  error={error}
                  treeShown={treeShown}
                  dirsShown={welcomeColumns.folders}
                  filesShown={welcomeColumns.files}
                />
              </div>
            )}
          </div>
        </div>

        <StatusBar
          fileName={active?.info.fileName ?? null}
          sheetCount={active?.info.sheets.length ?? null}
          sheet={active?.sheet ?? null}
          size={extent ? t("app.extent", { rows: extent.rows, cols: extent.cols }) : null}
          span={activeSheet ? selectionSize(selection) : null}
          filtered={active ? filtered : null}
          view={view}
          onView={handleView}
          zoom={zoom}
          onZoom={setZoom}
        />

        <AppDialogs
          gotoOpen={gotoOpen}
          setGotoOpen={setGotoOpen}
          sheetNames={shown.map((sheet) => sheet.name)}
          currentSheet={active?.sheet ?? ""}
          onGo={handleGo}
          unhideOpen={unhideOpen}
          setUnhideOpen={setUnhideOpen}
          hiddenSheets={hiddenNames}
          unhideAt={unhideAt}
          onUnhide={handleUnhideSheets}
          updateOpen={updateOpen}
          setUpdateOpen={setUpdateOpen}
          update={update}
          setUpdateAvailable={setUpdateAvailable}
        />

        <Toaster />
      </div>
    </TooltipProvider>
  );
}
