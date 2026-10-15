import { useEffect, useRef, useState } from "react";
import { Loader2, X } from "lucide-react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { cellRef } from "@/lib/column";
import { searchWorkbook, type SearchHit } from "@/lib/excel";
import { cn } from "@/lib/utils";

interface SearchPanelProps {
  path: string;
  sheet: string;
  onJump: (hit: SearchHit) => void;
  onClose: () => void;
}

type Scope = "sheet" | "workbook";

export function SearchPanel({ path, sheet, onJump, onClose }: SearchPanelProps) {
  const { t } = useTranslation();
  const [query, setQuery] = useState("");
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [scope, setScope] = useState<Scope>("sheet");
  const [hits, setHits] = useState<SearchHit[]>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  /**
   * The hit the arrows are on, or null before the first one. It is held as the
   * cell rather than as a place in the list because walking it moves the
   * sheet, and a jump to another sheet has the panel searching again — the
   * same hits come back as a new array, in which that cell is still the cell
   * the walk is on.
   */
  const [current, setCurrent] = useState<SearchHit | null>(null);
  /** Where in the list that is, or -1 for a cell this search no longer lists. */
  const at = current
    ? hits.findIndex(
        (hit) =>
          hit.sheet === current.sheet && hit.row === current.row && hit.col === current.col,
      )
    : -1;
  /** That row, so the list follows the highlight rather than losing it. */
  const currentRef = useRef<HTMLButtonElement>(null);

  // A different search is a different list to walk.
  useEffect(() => {
    setCurrent(null);
  }, [path, query, scope, caseSensitive, wholeCell]);

  useEffect(() => {
    currentRef.current?.scrollIntoView({ block: "nearest" });
  }, [at]);

  useEffect(() => {
    if (!query.trim()) {
      setHits([]);
      setError(null);
      setBusy(false);
      return;
    }

    let cancelled = false;
    setBusy(true);

    const timer = window.setTimeout(() => {
      searchWorkbook(path, query, scope === "sheet" ? sheet : null, {
        caseSensitive,
        wholeCell,
      })
        .then((result) => {
          if (cancelled) return;
          setHits(result);
          setError(null);
        })
        .catch((reason: unknown) => {
          if (cancelled) return;
          setHits([]);
          setError(String(reason));
        })
        .finally(() => {
          if (!cancelled) setBusy(false);
        });
    }, 200);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [path, sheet, query, caseSensitive, wholeCell, scope]);

  const toggleClass = (on: boolean) =>
    cn(
      "rounded-sm px-1.5 py-0.5 text-[10px] transition-colors",
      on
        ? "bg-primary text-primary-foreground"
        : "bg-muted text-muted-foreground hover:bg-accent",
    );

  /**
   * The arrows walk the list a hit at a time, and the sheet goes with them, so
   * that reading down the results is looking at them. Excel's Find All walks
   * its own list the same way.
   *
   * The ends hold rather than wrapping: walking off the last hit back to the
   * first would read as the list having jumped somewhere.
   */
  const step = (delta: number) => {
    if (hits.length === 0) return;
    const next = Math.min(Math.max(at + delta, 0), hits.length - 1);
    if (next === at) return;
    setCurrent(hits[next]);
    onJump(hits[next]);
  };

  return (
    <div
      className="flex h-full min-h-0 flex-col border-l border-border bg-sidebar"
      // On the panel rather than on the field, so the arrows answer wherever
      // in it the focus is — and so the default is taken away from the list,
      // which would otherwise scroll the hits past a highlight that never
      // moved.
      onKeyDown={(event) => {
        if (event.key !== "ArrowDown" && event.key !== "ArrowUp") return;
        if (hits.length === 0) return;
        event.preventDefault();
        step(event.key === "ArrowDown" ? 1 : -1);
      }}
    >
      <div className="flex shrink-0 items-center justify-between border-b border-border px-2 py-1.5">
        <span className="text-xs font-medium">{t("search.title")}</span>
        <Button
          size="sm"
          variant="ghost"
          className="size-6 p-0"
          onClick={onClose}
          aria-label={t("search.close")}
        >
          <X className="size-3.5" />
        </Button>
      </div>

      <div className="shrink-0 space-y-1.5 border-b border-border p-2">
        <Input
          autoFocus
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder={t("search.placeholder")}
          className="h-7 text-xs select-text"
        />

        <div className="flex flex-wrap items-center gap-1">
          <button
            type="button"
            className={toggleClass(scope === "sheet")}
            onClick={() => setScope("sheet")}
          >
            {t("search.thisSheet")}
          </button>
          <button
            type="button"
            className={toggleClass(scope === "workbook")}
            onClick={() => setScope("workbook")}
          >
            {t("search.workbook")}
          </button>
          <span className="mx-0.5 h-3 w-px bg-border" />
          <button
            type="button"
            className={toggleClass(caseSensitive)}
            onClick={() => setCaseSensitive((v) => !v)}
          >
            Aa
          </button>
          <button
            type="button"
            className={toggleClass(wholeCell)}
            onClick={() => setWholeCell((v) => !v)}
          >
            {t("search.wholeCell")}
          </button>
        </div>

        <div className="flex items-center gap-2 text-[10px] text-muted-foreground">
          {busy && <Loader2 className="size-3 animate-spin" />}
          <span>
            {query.trim()
              ? t("search.results", {
                  count: hits.length,
                  plus: hits.length >= 5000 ? "+" : "",
                })
              : t("search.typeToSearch")}
          </span>
        </div>
      </div>

      {error && (
        <p className="shrink-0 border-b border-border bg-destructive/10 px-2 py-1.5 text-[10px] text-destructive">
          {error}
        </p>
      )}

      <ScrollArea className="min-h-0 flex-1">
        <div className="flex flex-col gap-px p-1">
          {hits.map((hit, index) => {
            const walked = index === at;
            return (
              <button
                key={`${hit.sheet}:${hit.row}:${hit.col}`}
                ref={walked ? currentRef : null}
                type="button"
                onClick={() => {
                  setCurrent(hit);
                  onJump(hit);
                }}
                className={cn(
                  "rounded-sm px-2 py-1.5 text-left text-[11px]",
                  walked ? "bg-[#0F6CBD] text-white" : "hover:bg-accent/50",
                )}
              >
                <div
                  className={cn(
                    "flex items-center gap-1.5 text-[10px]",
                    walked ? "text-white/75" : "text-muted-foreground",
                  )}
                >
                  <span
                    className={cn(
                      "rounded-sm px-1 py-px font-mono font-medium",
                      walked ? "bg-white/25 text-white" : "bg-muted text-foreground",
                    )}
                  >
                    {cellRef(hit.row, hit.col)}
                  </span>
                  <span className="truncate" title={hit.sheet}>
                    {hit.sheet}
                  </span>
                </div>
                <div
                  className={cn(
                    "mt-0.5 truncate select-text",
                    walked ? "text-white" : "text-foreground/90",
                  )}
                  title={hit.text}
                >
                  {hit.text || " "}
                </div>
              </button>
            );
          })}
        </div>
      </ScrollArea>
    </div>
  );
}
