import { useEffect, useRef, useState } from "react";
import { ChevronDown, Loader2, X } from "lucide-react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
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

/**
 * The searches this session has run, newest first — and the whole of where they
 * live. They are the app's own memory rather than the reader's: a query is
 * worth having back while the workbooks it was asked of are still open, and a
 * workbook has nowhere to keep one anyway. Closing the app forgets them, which
 * is the whole of the forgetting.
 */
const RECENT_LIMIT = 8;
const recentQueries: string[] = [];

function remember(query: string) {
  const text = query.trim();
  if (!text || recentQueries[0] === text) return;
  const at = recentQueries.indexOf(text);
  if (at >= 0) recentQueries.splice(at, 1);
  recentQueries.unshift(text);
  if (recentQueries.length > RECENT_LIMIT) recentQueries.length = RECENT_LIMIT;
}

export function SearchPanel({ path, sheet, onJump, onClose }: SearchPanelProps) {
  const { t } = useTranslation();
  const [query, setQuery] = useState("");
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [scope, setScope] = useState<Scope>("sheet");
  const [hits, setHits] = useState<SearchHit[]>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  /** What to offer back, which is the list above as it stands. */
  const [recent, setRecent] = useState<string[]>(() => [...recentQueries]);
  const inputRef = useRef<HTMLInputElement>(null);
  const queryRef = useRef(query);
  queryRef.current = query;

  /**
   * A search is the reader's once they have looked somewhere with it, and the
   * box's own contents when the panel goes are the same act's leftovers.
   */
  const keep = (text: string) => {
    remember(text);
    setRecent([...recentQueries]);
  };

  useEffect(() => () => remember(queryRef.current), []);

  /**
   * The field's own box, so the list can hang under it at its width. Measured
   * rather than a class: the panel is dragged wider and narrower, and the list
   * is a menu in a portal, which no longer sits in a layout that could size it.
   */
  const fieldRef = useRef<HTMLDivElement>(null);
  const [fieldWidth, setFieldWidth] = useState(0);

  useEffect(() => {
    const el = fieldRef.current;
    if (!el) return;
    const measure = () => setFieldWidth(el.getBoundingClientRect().width);
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

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

  /** Everything that moves the sheet to a hit goes through here. */
  const jumpTo = (hit: SearchHit) => {
    setCurrent(hit);
    keep(query);
    onJump(hit);
  };

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
    jumpTo(hits[next]);
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
        {/* Excel's Find what: a field with the list of what has been searched
            for hanging off a button inside its right edge. */}
        <div ref={fieldRef} className="relative">
          <Input
            ref={inputRef}
            autoFocus
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder={t("search.placeholder")}
            className="h-7 pr-6 text-xs select-text"
          />
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              {/* The field's own height rather than a chip floating in it, so
                  the list it opens hangs under the field rather than under a
                  button that stops short of the field's bottom. */}
              <button
                type="button"
                aria-label={t("search.history")}
                title={t("search.history")}
                className="absolute inset-y-0.5 right-0.5 flex w-5 items-center justify-center rounded-md text-muted-foreground transition-colors hover:bg-accent hover:text-foreground"
              >
                <ChevronDown className="size-3.5" />
              </button>
            </DropdownMenuTrigger>
            <DropdownMenuContent
              // Hung under the field and as wide as it, off the button's own
              // right edge — that button is what opens it, and it sits at the
              // field's end.
              align="end"
              sideOffset={4}
              style={{ width: fieldWidth || undefined }}
              // The list is picked from to go on typing, so the field takes the
              // keyboard back rather than the button, which has done its job.
              onCloseAutoFocus={(event) => {
                event.preventDefault();
                inputRef.current?.focus();
              }}
            >
              <DropdownMenuLabel className="text-[10px]">{t("search.history")}</DropdownMenuLabel>
              {recent.length === 0 ? (
                <DropdownMenuItem disabled className="text-[11px] text-muted-foreground">
                  {t("search.historyEmpty")}
                </DropdownMenuItem>
              ) : (
                recent.map((text) => (
                  <DropdownMenuItem
                    key={text}
                    className="text-xs"
                    onSelect={() => setQuery(text)}
                  >
                    <span className="truncate">{text}</span>
                  </DropdownMenuItem>
                ))
              )}
            </DropdownMenuContent>
          </DropdownMenu>
        </div>

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

      {/* Always, not on hover: the arrows walk this list, and the bar is how
          far down the walk is — a bar that appears only when the pointer
          happens to be over it says nothing while the keyboard is doing the
          moving. */}
      <ScrollArea type="always" className="min-h-0 flex-1">
        <div className="flex flex-col gap-px p-1">
          {hits.map((hit, index) => {
            const walked = index === at;
            return (
              <button
                key={`${hit.sheet}:${hit.row}:${hit.col}`}
                ref={walked ? currentRef : null}
                type="button"
                onClick={() => jumpTo(hit)}
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
