import {
  useCallback,
  useEffect,
  useMemo,
  useState,
  type Dispatch,
  type SetStateAction,
} from "react";

import type { UnhideAnchor } from "@/components/UnhideDialog";
import { isHidden, shownSheets, type Tab } from "@/lib/tabs";

export interface UseSheetsOptions {
  active: Tab | null;
  activePath: string | null;
  setTabs: Dispatch<SetStateAction<Tab[]>>;
  updateActive: (patch: Partial<Omit<Tab, "info">>) => void;
  /** Takes the undo step an edit is about to need. */
  recordEdit: (path: string, sheet: string) => void;
}

/**
 * The workbook's sheets as the window acts on them: which are on the strip,
 * which is grouped with which, hiding and bringing back, stepping along, and
 * dragging one tab past another. Nothing here reaches the file — every answer
 * is the reader's, and lasts as long as the tab.
 */
export function useSheets({
  active,
  activePath,
  setTabs,
  updateActive,
  recordEdit,
}: UseSheetsOptions) {
  /**
   * The tab strip's own multi-selection: Ctrl+click groups sheets and
   * Shift+click takes a run of them, as Excel's tabs do. The displayed sheet is
   * always one of them; what the group is for here is acting on several sheets
   * at once — Hide.
   */
  const [sheetGroup, setSheetGroup] = useState<{ names: string[]; anchor: string } | null>(null);
  /** The sheet dialog Excel opens from a tab's own menu — see `SheetTabs`. */
  const [unhideOpen, setUnhideOpen] = useState(false);
  /**
   * Where the menu that opens it stood, so the dialog can take its place rather
   * than the centred spot the shared dialog would put it in.
   */
  const [unhideAt, setUnhideAt] = useState<UnhideAnchor | null>(null);

  // A group belongs to the file it was made in.
  useEffect(() => {
    setSheetGroup(null);
  }, [activePath]);

  /** The tabs the strip draws, and the sheets Unhide offers, in sheet order. */
  const shown = useMemo(() => (active ? shownSheets(active) : []), [active]);
  const hiddenNames = useMemo(
    () => (active ? active.info.sheets.filter((s) => isHidden(active, s)).map((s) => s.name) : []),
    [active],
  );

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
        const current = (sheetGroup?.names ?? [active.sheet]).filter((one) => order.includes(one));
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
      recordEdit(active.info.path, active.sheet);
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
            const stillShown = new Set(shownSheets({ ...tab, sheetHidden }).map((one) => one.name));
            const at = tab.info.sheets.findIndex((one) => one.name === sheet);
            const next =
              tab.info.sheets.slice(at + 1).find((one) => stillShown.has(one.name)) ??
              [...tab.info.sheets.slice(0, at)].reverse().find((one) => stillShown.has(one.name));
            sheet = next?.name ?? sheet;
          }
          return { ...tab, sheetHidden, sheet };
        }),
      );
      setSheetGroup(null);
    },
    [active, setTabs, recordEdit],
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
    (sheet: string, hidden: boolean, record = true) => {
      if (!active) return;
      // A sweep — the Unhide dialog bringing several back at once — takes its
      // one step itself, rather than a step per name that all restore the same.
      if (record) recordEdit(active.info.path, sheet);
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
    [active, setTabs, recordEdit],
  );

  /**
   * Bringing sheets back makes the last one picked the current one, as Excel's
   * Unhide does — you asked to see it, so it is what the window shows.
   */
  const handleUnhideSheets = useCallback(
    (names: string[]) => {
      if (names.length > 0 && active) recordEdit(active.info.path, active.sheet);
      for (const name of names) applySheetHidden(name, false, false);
      setSheetGroup(null);
      if (names.length) updateActive({ sheet: names[names.length - 1] });
    },
    [active, applySheetHidden, recordEdit, updateActive],
  );

  const handleSheetSelect = useCallback(
    (name: string) => {
      setSheetGroup(null);
      updateActive({ sheet: name });
    },
    [updateActive],
  );

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
    [active, setTabs],
  );

  return {
    sheetGroup,
    sheetSelection,
    shown,
    hiddenNames,
    unhideOpen,
    setUnhideOpen,
    unhideAt,
    setUnhideAt,
    handleSheetClick,
    applySheetsHidden,
    handleHideSheet,
    applySheetHidden,
    handleUnhideSheets,
    handleSheetSelect,
    handleSheetStep,
    handleReorderSheets,
  };
}
