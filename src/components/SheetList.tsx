import { Fragment, useEffect, useMemo, useState } from "react";
import { Pin, PinOff, Search, X } from "lucide-react";
import { useTranslation } from "react-i18next";

import { Hint } from "@/components/Hint";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { copyText } from "@/lib/clipboard";
import type { SheetInfo } from "@/lib/excel";
import { pinSheet, readPinnedSheets, unpinSheet } from "@/lib/sheetPins";
import { cn } from "@/lib/utils";

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

  const visible = useMemo(() => {
    const needle = filter.trim().toLowerCase();
    const matched = needle
      ? sheets.filter((sheet) => sheet.name.toLowerCase().includes(needle))
      : sheets;
    // The fixed ones come first; each group keeps the workbook's own order, so
    // the list does not reshuffle under a reader who fixes a second sheet.
    return [
      ...matched.filter((sheet) => pinned.has(sheet.name)),
      ...matched.filter((sheet) => !pinned.has(sheet.name)),
    ];
  }, [sheets, filter, pinned]);

  const togglePin = (name: string) =>
    setPins(pinned.has(name) ? unpinSheet(path, name) : pinSheet(path, name));

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

      <div className="shrink-0 px-2 py-1 text-[10px] font-medium tracking-wide text-muted-foreground uppercase">
        {t("sheets.count", { count: sheets.length })}
      </div>

      <ScrollArea className="min-h-0 flex-1">
        <div className="flex flex-col gap-px p-1">
          {visible.map((sheet, index) => {
            /**
             * A hidden sheet is listed but goes nowhere: it has no tab to select
             * from, and a click that opened one would leave the strip with no
             * current sheet on it. Excel's own sheet list greys them the same
             * way, and says which they are rather than only that they are dim.
             */
            const isHidden = hidden.has(sheet.name);
            const isPinned = pinned.has(sheet.name);
            // A rule under the fixed ones, so the list reads as two runs: what
            // stays at the top, and the rest in the workbook's own order.
            const endsFixed =
              isPinned &&
              (index + 1 >= visible.length || !pinned.has(visible[index + 1].name));
            const pinLabel = t(isPinned ? "sheets.unpin" : "sheets.pin", {
              name: sheet.name,
            });
            return (
              <Fragment key={sheet.name}>
                <div
                  className={cn(
                    "group flex items-center rounded-sm transition-colors",
                    sheet.name === active
                      ? "bg-accent"
                      : !isHidden && "hover:bg-accent/50",
                  )}
                >
                  <button
                    type="button"
                    disabled={isHidden}
                    onClick={() => onSelect(sheet.name)}
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
                  {/* Not offered on a hidden sheet: it has no tab to come back
                      to, so there is nothing a fixed place would save. */}
                  {!isHidden && (
                    <Hint label={pinLabel}>
                      <button
                        type="button"
                        onClick={() => togglePin(sheet.name)}
                        aria-label={pinLabel}
                        aria-pressed={isPinned}
                        className={cn(
                          "mr-1 shrink-0 rounded-sm p-0.5 text-muted-foreground transition-opacity hover:bg-accent hover:text-foreground",
                          // A fixed sheet shows the way back out of being fixed
                          // without being asked for; the rest offer it under a
                          // pointer already on the row.
                          isPinned
                            ? "opacity-60"
                            : "opacity-0 group-hover:opacity-60 focus-visible:opacity-100",
                        )}
                      >
                        {isPinned ? <PinOff className="size-3" /> : <Pin className="size-3" />}
                      </button>
                    </Hint>
                  )}
                </div>
                {endsFixed && <div className="my-1 border-t border-border" />}
              </Fragment>
            );
          })}
          {visible.length === 0 && (
            <p className="px-2 py-4 text-xs text-muted-foreground">
              {t("sheets.none")}
            </p>
          )}
        </div>
      </ScrollArea>
    </div>
  );
}
