import { useEffect, useState } from "react";
import { Loader2 } from "lucide-react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Input } from "@/components/ui/input";
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from "@/components/ui/popover";
import { ScrollArea } from "@/components/ui/scroll-area";
import { columnValues, type ValueCount } from "@/lib/excel";

interface FilterButtonProps {
  path: string;
  sheet: string;
  col: number;
  /** Data rows only — the header row is excluded. */
  firstRow: number;
  lastRow: number;
  selected: string[];
  /**
   * The file filters this column with a condition ("&gt; 20"), which the value
   * list cannot show — the funnel is the only sign Excel gives it.
   */
  filtered?: boolean;
  onApply: (values: string[] | null) => void;
  /** The grid's own zoom, so the button scales with the cells it sits in. */
  zoom: number;
}

export function FilterButton({
  path,
  sheet,
  col,
  firstRow,
  lastRow,
  selected,
  filtered,
  onApply,
  zoom,
}: FilterButtonProps) {
  const { t } = useTranslation();
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [options, setOptions] = useState<ValueCount[]>([]);
  const [draft, setDraft] = useState<Set<string>>(() => new Set());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!open) return;
    setDraft(new Set(selected));
    setQuery("");
  }, [open, selected]);

  useEffect(() => {
    if (!open) return;
    let cancelled = false;
    setBusy(true);
    // Typing in the search box re-queries the backend, so a column with more
    // distinct values than the cap still remains reachable.
    const timer = window.setTimeout(() => {
      columnValues(path, sheet, col, firstRow, lastRow, query)
        .then((result) => {
          if (cancelled) return;
          setOptions(result);
          setError(null);
        })
        .catch((reason: unknown) => {
          if (cancelled) return;
          setError(String(reason));
        })
        .finally(() => {
          if (!cancelled) setBusy(false);
        });
    }, 180);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [open, path, sheet, col, firstRow, lastRow, query]);

  const toggle = (value: string) =>
    setDraft((prev) => {
      const next = new Set(prev);
      if (next.has(value)) next.delete(value);
      else next.add(value);
      return next;
    });

  const isFiltered = selected.length > 0 || filtered === true;

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <button
          type="button"
          // The cell underneath selects on click; the button must not.
          onClick={(event) => event.stopPropagation()}
          aria-label={isFiltered ? t("filter.edit") : t("filter.column")}
          // Measured off Excel's own render (CopyPicture of a filtered range):
          // a 17 px square pinned to the bottom-right corner with a 1 px inset,
          // a 1 px #a6acb3 edge and a white-to-grey face. A filtered column does
          // not change the button — only the glyph, a caret becomes a funnel.
          // `pointer-events-auto` because the canvas grid puts this inside a
          // click-through overlay, where the default would swallow the press.
          style={{
            width: 17 * zoom,
            height: 17 * zoom,
            right: zoom,
            bottom: zoom,
          }}
          className="pointer-events-auto absolute flex items-center justify-center border border-[#a6acb3] bg-linear-to-b from-[#feffff] to-[#eff2f7] text-[#656769] hover:brightness-[0.97]"
        >
          <svg
            width={15 * zoom}
            height={15 * zoom}
            viewBox="0 0 15 15"
            aria-hidden="true"
            fill="currentColor"
          >
            {isFiltered ? (
              // Measured off Excel's own button: a 9-px funnel that tapers to a
              // 3-px stem high on the right, and a small caret under its left.
              <path d="M4.5 3H15L11 6V11H8V6Z M1.5 9H7.5L4.5 12Z" />
            ) : (
              // The plain caret: 7 x 4, centred.
              <path d="M4 6h7l-3.5 3.5z" />
            )}
          </svg>
        </button>
      </PopoverTrigger>

      <PopoverContent
        align="start"
        className="w-64 p-0"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="border-b border-border p-1.5">
          <Input
            autoFocus
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder={t("filter.searchValues")}
            className="h-6 text-[11px]"
          />
        </div>

        <ScrollArea className="h-56">
          <div className="flex flex-col gap-px p-1">
            {options.map((option) => (
              <label
                key={option.value || "\u0000empty"}
                className="flex cursor-pointer items-center gap-2 rounded-sm px-1.5 py-1 text-[11px] hover:bg-accent"
              >
                <Checkbox
                  checked={draft.has(option.value)}
                  onCheckedChange={() => toggle(option.value)}
                  className="size-3.5"
                />
                <span className="min-w-0 flex-1 truncate" title={option.value}>
                  {option.value || t("filter.blank")}
                </span>
                <span className="shrink-0 tabular-nums text-muted-foreground">
                  {option.count}
                </span>
              </label>
            ))}

            {options.length === 0 && !busy && (
              <p className="px-2 py-3 text-[11px] text-muted-foreground">
                {t("filter.noMatch")}
              </p>
            )}
          </div>
        </ScrollArea>

        {error && (
          <p className="border-t border-border px-2 py-1 text-[10px] text-destructive">
            {error}
          </p>
        )}

        <div className="flex items-center justify-between gap-2 border-t border-border p-1.5">
          <span className="flex items-center gap-1 text-[10px] text-muted-foreground">
            {busy && <Loader2 className="size-3 animate-spin" />}
            {draft.size > 0
              ? t("filter.selected", { count: draft.size })
              : t("filter.allValues")}
          </span>
          <div className="flex items-center gap-1">
            <Button
              size="sm"
              variant="ghost"
              className="h-6 px-2 text-[11px]"
              onClick={() => {
                onApply(null);
                setOpen(false);
              }}
            >
              {t("filter.clear")}
            </Button>
            <Button
              size="sm"
              className="h-6 px-2 text-[11px]"
              onClick={() => {
                onApply(draft.size > 0 ? [...draft] : null);
                setOpen(false);
              }}
            >
              {t("filter.apply")}
            </Button>
          </div>
        </div>
      </PopoverContent>
    </Popover>
  );
}
