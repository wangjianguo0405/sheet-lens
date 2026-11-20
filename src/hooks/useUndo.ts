import { useCallback, useRef, useState, type RefObject } from "react";

import type { SheetView, SheetViews } from "@/components/CanvasGrid";
import type { Formats, TextColors } from "@/lib/formats";

/** The tab fields a step puts back, as the caller reads them from its own tab. */
export interface UndoTabState {
  marks: Map<string, string>;
  formats: Formats;
  textColors: TextColors;
  tabColors: Map<string, string>;
  sheetHidden: Map<string, boolean>;
}

/** The part of a sheet the grid keeps for the session: sizes, hidden lines, shape offsets. */
type GridSlice = Pick<SheetView, "sizes" | "hidden" | "shapes">;

/** One step's state: the tab's own editing state, plus the grid's slices by sheet. */
export interface SessionSnapshot extends UndoTabState {
  path: string;
  /** The sheet the edit was made on, which undo puts back on screen. */
  sheet: string;
  /** By sheet name, exactly the slices kept in the view store. */
  grid: Map<string, GridSlice>;
}

export interface UseUndoOptions {
  /** The per-sheet state the grid keeps; a step's slices are written back into it. */
  views: RefObject<SheetViews>;
  /** The tab as it stands, or null when it is not open. */
  readTab: (path: string) => UndoTabState | null;
  /** Brings a tab on screen, so the step it is about to restore can be seen. */
  selectTab: (path: string) => void;
  /** Puts a step's tab fields back; the tab's own shape stays the caller's. */
  restoreTab: (step: SessionSnapshot) => void;
}

/**
 * The undo history: a snapshot of the session's own editing state taken just
 * before each edit. A tab keeps its marks, formats, text colours, tab colours
 * and hidden sheets; the grid keeps each sheet's sizes, hidden lines and moved
 * shapes in the view store. What a step puts back is exactly those. View
 * decisions — scroll, zoom, freeze, filters — are outside it, as they are in
 * Excel.
 *
 * A snapshot holds the state's own objects rather than copies: every edit
 * builds new maps instead of changing old ones, so a step costs references.
 */
export function useUndo({ views, readTab, selectTab, restoreTab }: UseUndoOptions) {
  const undoStack = useRef<SessionSnapshot[]>([]);
  const redoStack = useRef<SessionSnapshot[]>([]);
  const [depth, setDepth] = useState({ undo: 0, redo: 0 });
  /** The caller's callbacks as they stand now, so a step never runs a stale one. */
  const latest = useRef({ readTab, selectTab, restoreTab });
  latest.current = { readTab, selectTab, restoreTab };

  const snapshotSession = useCallback(
    (path: string, sheet: string): SessionSnapshot | null => {
      const tab = latest.current.readTab(path);
      if (!tab) return null;
      const grid = new Map<string, GridSlice>();
      for (const [name, view] of views.current?.get(path) ?? []) {
        grid.set(name, { sizes: view.sizes, hidden: view.hidden, shapes: view.shapes });
      }
      return { path, sheet, ...tab, grid };
    },
    [views],
  );

  /** Takes the step an edit is about to need. */
  const recordEdit = useCallback(
    (path: string, sheet: string) => {
      const snapshot = snapshotSession(path, sheet);
      if (!snapshot) return;
      undoStack.current.push(snapshot);
      redoStack.current = [];
      setDepth({ undo: undoStack.current.length, redo: 0 });
    },
    [snapshotSession],
  );

  /** Puts a step back — the tab's own state first, then the grid's slices. */
  const apply = useCallback(
    (step: SessionSnapshot) => {
      latest.current.selectTab(step.path);
      latest.current.restoreTab(step);
      const bySheet = views.current?.get(step.path) ?? new Map<string, SheetView>();
      for (const [name, slice] of step.grid) {
        const view = bySheet.get(name) ?? { top: 0, left: 0 };
        bySheet.set(name, {
          ...view,
          sizes: slice.sizes,
          hidden: slice.hidden,
          shapes: slice.shapes,
        });
      }
      views.current?.set(step.path, bySheet);
    },
    [views],
  );

  const undo = useCallback(() => {
    const step = undoStack.current.pop();
    if (!step) return;
    const now = snapshotSession(step.path, step.sheet);
    if (now) redoStack.current.push(now);
    apply(step);
    setDepth({ undo: undoStack.current.length, redo: redoStack.current.length });
  }, [apply, snapshotSession]);

  const redo = useCallback(() => {
    const step = redoStack.current.pop();
    if (!step) return;
    const now = snapshotSession(step.path, step.sheet);
    if (now) undoStack.current.push(now);
    apply(step);
    setDepth({ undo: undoStack.current.length, redo: redoStack.current.length });
  }, [apply, snapshotSession]);

  /** A closed workbook takes its steps with it. */
  const forgetHistory = useCallback((path: string) => {
    undoStack.current = undoStack.current.filter((step) => step.path !== path);
    redoStack.current = redoStack.current.filter((step) => step.path !== path);
    setDepth({ undo: undoStack.current.length, redo: redoStack.current.length });
  }, []);

  return {
    recordEdit,
    undo,
    redo,
    forgetHistory,
    canUndo: depth.undo > 0,
    canRedo: depth.redo > 0,
  };
}
