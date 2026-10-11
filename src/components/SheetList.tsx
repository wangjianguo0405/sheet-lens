import { useMemo, useState } from "react";
import { Search, X } from "lucide-react";
import { useTranslation } from "react-i18next";

import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { copyText } from "@/lib/clipboard";
import type { SheetInfo } from "@/lib/excel";
import { cn } from "@/lib/utils";

interface SheetListProps {
  sheets: SheetInfo[];
  active: string;
  onSelect: (name: string) => void;
}

export function SheetList({ sheets, active, onSelect }: SheetListProps) {
  const { t } = useTranslation();
  const [filter, setFilter] = useState("");

  const visible = useMemo(() => {
    const needle = filter.trim().toLowerCase();
    if (!needle) return sheets;
    return sheets.filter((sheet) => sheet.name.toLowerCase().includes(needle));
  }, [sheets, filter]);

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
          {visible.map((sheet) => (
            <button
              key={sheet.name}
              type="button"
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
                "flex w-full items-center justify-between gap-2 rounded-sm px-2 py-1.5 text-left text-xs transition-colors",
                sheet.name === active
                  ? "bg-accent font-medium text-accent-foreground"
                  : "text-foreground/80 hover:bg-accent/50",
              )}
            >
              <span className="truncate" title={sheet.name}>
                {sheet.name}
              </span>
              <span className="shrink-0 text-[10px] tabular-nums text-muted-foreground">
                {sheet.rows}×{sheet.cols}
              </span>
            </button>
          ))}
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
