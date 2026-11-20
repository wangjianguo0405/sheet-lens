import {
  useCallback,
  useMemo,
  useRef,
  useState,
  type Dispatch,
  type RefObject,
  type SetStateAction,
} from "react";
import { open as openFileDialog } from "@tauri-apps/plugin-dialog";

import type { SheetViews } from "@/components/CanvasGrid";
import { closeWorkbook, openWorkbook } from "@/lib/excel";
import { initialView } from "@/lib/viewMode";
import { savedCursor } from "@/lib/selection";
import { isHidden, type Tab } from "@/lib/tabs";
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

const EXCEL_EXTENSIONS = ["xlsx", "xlsm", "xls", "xlsb", "ods"];

export function looksLikeWorkbook(path: string): boolean {
  const lower = path.toLowerCase();
  return EXCEL_EXTENSIONS.some((ext) => lower.endsWith(`.${ext}`));
}

export interface UseWorkbooksOptions {
  setError: Dispatch<SetStateAction<string | null>>;
  /** The per-sheet state a closed tab takes with it. */
  views: RefObject<SheetViews>;
  /**
   * Drops a tab's undo steps when it closes. A ref because the history is built
   * on this hook's own `selectTab`, so the two are wired in that order.
   */
  forgetHistoryRef: RefObject<(path: string) => void>;
}

/**
 * The workbooks the window holds open, as tabs, and the places they come from:
 * the picker, the File menu's recent list, the start page's own tab, the tab
 * strip's order, and closing or reading one again.
 */
export function useWorkbooks({ setError, views, forgetHistoryRef }: UseWorkbooksOptions) {
  const [tabs, setTabs] = useState<Tab[]>([]);
  const [activePath, setActivePath] = useState<string | null>(null);
  /** The start page shown as a tab, created by the "+" and closed by its own ×. */
  const [homeOpen, setHomeOpen] = useState(false);
  /** The workbook to put back when that tab is closed. */
  const lastFileRef = useRef<string | null>(null);
  /** The workbook paths the File menu offers, newest first. */
  const [recentFiles, setRecentFiles] = useState<string[]>(readRecentFiles);
  /** Those of them that would not open, which the menu greys out. */
  const [missingFiles, setMissingFiles] = useState<string[]>(readMissingFiles);

  const active = useMemo(
    () => tabs.find((tab) => tab.info.path === activePath) ?? null,
    [tabs, activePath],
  );

  const activeSheet = useMemo(() => {
    if (!active) return null;
    return active.info.sheets.find((sheet) => sheet.name === active.sheet) ?? null;
  }, [active]);

  const openPaths = useCallback(
    async (paths: string[]) => {
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
                    textColors: new Map(),
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
    },
    [setError],
  );

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
    [openPaths, setError],
  );

  const handleOpen = useCallback(() => {
    void pickWorkbooks();
  }, [pickWorkbooks]);

  /**
   * A folder off the start page's own column: the same picker, already inside
   * it. The folder button is not a second way to open a workbook — it is the
   * same way, starting one step in.
   */
  const handleOpenInDirectory = useCallback(
    (directory: string) => {
      void pickWorkbooks(directory);
    },
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

  const updateActive = useCallback(
    (patch: Partial<Omit<Tab, "info">>) => {
      setTabs((prev) =>
        prev.map((tab) => (tab.info.path === activePath ? { ...tab, ...patch } : tab)),
      );
    },
    [activePath],
  );

  const closeTab = useCallback(
    (path: string) => {
      void closeWorkbook(path).catch(() => {});
      views.current?.delete(path);
      forgetHistoryRef.current(path);
      const remaining = tabs.filter((tab) => tab.info.path !== path);
      setTabs(remaining);
      if (activePath === path) setActivePath(remaining.at(-1)?.info.path ?? null);
    },
    [tabs, activePath, views, forgetHistoryRef],
  );

  /**
   * Closes several tabs at once: the strip's menu takes away the others, the
   * ones on one side of it, or all of them. `activate` is the tab the menu was
   * opened on — it survives the sweep and is the one left showing, the way a
   * browser leaves the tab you cleared the rest out from.
   */
  const closeTabs = useCallback(
    (paths: string[], activate?: string) => {
      if (paths.length === 0) return;
      const gone = new Set(paths);
      for (const path of paths) {
        void closeWorkbook(path).catch(() => {});
        views.current?.delete(path);
        forgetHistoryRef.current(path);
      }
      const remaining = tabs.filter((tab) => !gone.has(tab.info.path));
      setTabs(remaining);
      if (activate !== undefined && !gone.has(activate)) setActivePath(activate);
      else if (activePath !== null && gone.has(activePath)) {
        setActivePath(remaining.at(-1)?.info.path ?? null);
      }
    },
    [tabs, activePath, views, forgetHistoryRef],
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
  const reloadTab = useCallback(
    async (path: string) => {
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
    },
    [setError],
  );

  /**
   * File ▸ Close, which is the tab's × for whichever workbook is showing. It is
   * nothing at all while the start page is, which is why the entry is offered
   * disabled rather than left to do something surprising.
   */
  const handleCloseWorkbook = useCallback(() => {
    if (activePath !== null) closeTab(activePath);
  }, [activePath, closeTab]);

  return {
    tabs,
    setTabs,
    activePath,
    setActivePath,
    active,
    activeSheet,
    homeOpen,
    setHomeOpen,
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
  };
}
