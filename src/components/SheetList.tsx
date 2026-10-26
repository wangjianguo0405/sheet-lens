import { useEffect, useMemo, useRef, useState } from "react";
import { Pin, PinOff, Search, X } from "lucide-react";
import { useTranslation } from "react-i18next";

import { Hint } from "@/components/Hint";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { copyText } from "@/lib/clipboard";
import type { SheetInfo } from "@/lib/excel";
import {
  pinSheet,
  readPinnedSheets,
  setPinnedSheets,
  unpinSheet,
} from "@/lib/sheetPins";
import { cn } from "@/lib/utils";

/** A group's heading, the same type the start page puts over its own lists. */
const HEAD =
  "px-2 py-1 text-[10px] font-medium tracking-wide text-muted-foreground uppercase";

interface SheetListProps {
  /** The workbook the list belongs to, which is what the fixes are kept by. */
  path: string;
  sheets: SheetInfo[];
  /** The sheets off the tab strip, which are listed but go nowhere. */
  hidden: Set<string>;
  active: string;
  onSelect: (name: string) => void;
}

export function SheetList({ path, sheets, hidden, active, onSelect }: SheetListProps) {
  const { t } = useTranslation();
  const [filter, setFilter] = useState("");
  /** Fixed by hand, and kept here rather than in the app: this list is its
      only reader, and mounting reads the names back out of storage. */
  const [pins, setPins] = useState<string[]>(() => readPinnedSheets(path));
  const pinned = useMemo(() => new Set(pins), [pins]);

  // Another tab's workbook brings its own list of fixed sheets with it.
  useEffect(() => {
    setPins(readPinnedSheets(path));
  }, [path]);

  const { fixed, rest } = useMemo(() => {
    const needle = filter.trim().toLowerCase();
    const matched = needle
      ? sheets.filter((sheet) => sheet.name.toLowerCase().includes(needle))
      : sheets;
    const byName = new Map(matched.map((sheet) => [sheet.name, sheet]));
    return {
      // Most recently fixed first, the order the start page keeps its own two
      // columns in. A name the workbook no longer has simply drops out.
      fixed: pins.flatMap((name) => {
        const sheet = byName.get(name);
        return sheet ? [sheet] : [];
      }),
      rest: matched.filter((sheet) => !pinned.has(sheet.name)),
    };
  }, [sheets, filter, pins, pinned]);

  const togglePin = (name: string) =>
    setPins(pinned.has(name) ? unpinSheet(path, name) : pinSheet(path, name));

  /** The fixed sheet under the pointer, and the edge it would land on. */
  const [dragging, setDragging] = useState<string | null>(null);
  const [dropAt, setDropAt] = useState<{ name: string; after: boolean } | null>(null);
  /**
   * A press on a fixed row that may yet become a drag: which sheet, where the
   * pointer went down, and whether it has moved far enough to be one.
   *
   * The pointer rather than the platform's own drag, because a window that
   * takes dropped files — which is how a workbook is opened — makes the
   * webview refuse one: the platform's drag turns into a no-drop cursor and
   * never starts. The window follows the pointer, since it leaves the row at
   * once.
   */
  const press = useRef<{ name: string; y: number; moved: boolean } | null>(null);
  const over = useRef<{ name: string; after: boolean } | null>(null);
  /** Set by a drop, so letting go of the drag is not also a click on the sheet. */
  const dropped = useRef(false);

  /**
   * Carries a pulled sheet to where it was let go. Only the fixed run is
   * ordered by hand — the rest follows the workbook, which is not this app's
   * to change — so this moves a name within the fixed list and writes it back.
   */
  const moveFixed = (from: string, to: string, after: boolean) => {
    const order = pins.filter((name) => name !== from);
    order.splice(order.indexOf(to) + (after ? 1 : 0), 0, from);
    setPins(setPinnedSheets(path, order));
  };
  // The window's listeners are bound once, so they reach the newest of this.
  const moveRef = useRef(moveFixed);
  moveRef.current = moveFixed;

  useEffect(() => {
    /** The fixed row the pointer is over, and which half of it. */
    const edgeAt = (clientY: number) => {
      const rows = [...document.querySelectorAll<HTMLElement>("[data-fixed]")];
      for (const row of rows) {
        const box = row.getBoundingClientRect();
        if (clientY < box.top || clientY > box.bottom) continue;
        return {
          name: row.dataset.fixed!,
          after: clientY > box.top + box.height / 2,
        };
      }
      // Past the last of them, the run's own end is the place.
      const last = rows[rows.length - 1];
      if (last && clientY > last.getBoundingClientRect().bottom) {
        return { name: last.dataset.fixed!, after: true };
      }
      return null;
    };

    const onMove = (event: PointerEvent) => {
      const state = press.current;
      if (!state) return;
      if (!state.moved) {
        if (Math.abs(event.clientY - state.y) < 4) return;
        state.moved = true;
        setDragging(state.name);
        document.body.style.cursor = "grabbing";
      }
      over.current = edgeAt(event.clientY);
      setDropAt(over.current);
    };

    const onUp = () => {
      const state = press.current;
      press.current = null;
      document.body.style.cursor = "";
      if (!state?.moved) return;
      const edge = over.current;
      if (edge && edge.name !== state.name) {
        moveRef.current(state.name, edge.name, edge.after);
      }
      over.current = null;
      dropped.current = true;
      setDragging(null);
      setDropAt(null);
    };

    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
      document.body.style.cursor = "";
    };
  }, []);

  const row = (sheet: SheetInfo) => {
    /**
     * A hidden sheet is listed but goes nowhere: it has no tab to select
     * from, and a click that opened one would leave the strip with no
     * current sheet on it. Excel's own sheet list greys them the same
     * way, and says which they are rather than only that they are dim.
     */
    const isHidden = hidden.has(sheet.name);
    const isPinned = pinned.has(sheet.name);
    const pinLabel = t(isPinned ? "sheets.unpin" : "sheets.pin", {
      name: sheet.name,
    });
    const edge = dropAt?.name === sheet.name ? (dropAt.after ? "after" : "before") : null;
    return (
      <div
        key={sheet.name}
        // Only the fixed run is reordered by hand; the rest follows the file.
        data-fixed={isPinned ? sheet.name : undefined}
        onPointerDown={(event) => {
          if (!isPinned) return;
          // The pin is its own control; pulling a sheet is the row's.
          if ((event.target as HTMLElement).closest("[data-sheet-pin]")) return;
          dropped.current = false;
          press.current = { name: sheet.name, y: event.clientY, moved: false };
        }}
        className={cn(
          "group flex items-center rounded-sm transition-colors",
          sheet.name === active ? "bg-accent" : !isHidden && "hover:bg-accent/50",
          isPinned && "cursor-grab",
          dragging === sheet.name && "opacity-40",
          edge === "before" && "shadow-[inset_0_2px_0_0_#107C41]",
          edge === "after" && "shadow-[inset_0_-2px_0_0_#107C41]",
        )}
      >
        <button
          type="button"
          disabled={isHidden}
          onClick={() => {
            // Letting go of a pull is not a click on the sheet under it.
            if (dropped.current) return;
            onSelect(sheet.name);
          }}
          onKeyDown={(event) => {
            // Ctrl+C on a focused sheet copies its name, matching the
            // "copy whatever is selected" behaviour of the grid.
            if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
              event.preventDefault();
              void copyText(sheet.name);
            }
          }}
          className={cn(
            "flex min-w-0 flex-1 items-center justify-between gap-2 rounded-sm px-2 py-1.5 text-left text-xs",
            sheet.name === active
              ? "font-medium text-accent-foreground"
              : isHidden
                ? "text-muted-foreground/60"
                : "text-foreground/80",
          )}
        >
          <span className="truncate" title={sheet.name}>
            {sheet.name}
          </span>
          <span className="shrink-0 text-[10px] tabular-nums text-muted-foreground">
            {isHidden ? t("sheets.hidden") : `${sheet.rows}×${sheet.cols}`}
          </span>
        </button>
        {/* Not offered on a hidden sheet: it has no tab to come back to, so
            there is nothing a fixed place would save. */}
        {!isHidden && (
          <Hint label={pinLabel}>
            <button
              type="button"
              data-sheet-pin
              onClick={() => togglePin(sheet.name)}
              aria-label={pinLabel}
              aria-pressed={isPinned}
              className={cn(
                "shrink-0 rounded-sm p-1 text-muted-foreground transition-opacity hover:bg-accent hover:text-foreground hover:opacity-100",
                // A fixed sheet shows the way back out of being fixed without
                // being asked for; the rest offer it under a pointer already on
                // the row.
                isPinned ? "opacity-60" : "opacity-0 group-hover:opacity-60",
              )}
            >
              {isPinned ? <PinOff className="size-3" /> : <Pin className="size-3" />}
            </button>
          </Hint>
        )}
      </div>
    );
  };

  return (
    <div className="flex h-full min-h-0 flex-col border-r border-border bg-sidebar">
      <div className="flex shrink-0 items-center gap-1.5 border-b border-border px-2 py-1.5">
        <Search className="size-3.5 shrink-0 text-muted-foreground" />
        <Input
          value={filter}
          onChange={(event) => setFilter(event.target.value)}
          placeholder={t("sheets.filter")}
          className="h-6 border-0 bg-transparent px-0 text-xs shadow-none select-text focus-visible:ring-0"
        />
        {filter && (
          <button
            type="button"
            aria-label={t("sheets.clearFilter")}
            onClick={() => setFilter("")}
            className="shrink-0 rounded-sm p-0.5 text-muted-foreground hover:bg-accent"
          >
            <X className="size-3" />
          </button>
        )}
      </div>

      <ScrollArea className="min-h-0 flex-1">
        <div className="flex flex-col gap-px p-1">
          {/* The fixed ones are their own run above the count, as the start
              page's pinned files are above its history. */}
          {fixed.length > 0 && (
            <>
              <h2 className={HEAD}>{t("sheets.pinned")}</h2>
              {fixed.map(row)}
            </>
          )}
          <h2 className={cn(HEAD, fixed.length > 0 && "mt-4")}>
            {t("sheets.count", { count: sheets.length })}
          </h2>
          {rest.map(row)}
          {rest.length === 0 && fixed.length === 0 && (
            <p className="px-2 py-4 text-xs text-muted-foreground">
              {t("sheets.none")}
            </p>
          )}
        </div>
      </ScrollArea>
    </div>
  );
}
