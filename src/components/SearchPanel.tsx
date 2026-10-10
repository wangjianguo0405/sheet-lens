import { useEffect, useState } from "react";
import { Loader2, X } from "lucide-react";

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
  const [query, setQuery] = useState("");
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [scope, setScope] = useState<Scope>("sheet");
  const [hits, setHits] = useState<SearchHit[]>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

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

  return (
    <div className="flex h-full min-h-0 flex-col border-l border-border bg-sidebar">
      <div className="flex shrink-0 items-center justify-between border-b border-border px-2 py-1.5">
        <span className="text-xs font-medium">Search</span>
        <Button
          size="sm"
          variant="ghost"
          className="size-6 p-0"
          onClick={onClose}
          aria-label="Close search"
        >
          <X className="size-3.5" />
        </Button>
      </div>

      <div className="shrink-0 space-y-1.5 border-b border-border p-2">
        <Input
          autoFocus
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder="Find in sheet…"
          className="h-7 text-xs select-text"
        />

        <div className="flex flex-wrap items-center gap-1">
          <button
            type="button"
            className={toggleClass(scope === "sheet")}
            onClick={() => setScope("sheet")}
          >
            This sheet
          </button>
          <button
            type="button"
            className={toggleClass(scope === "workbook")}
            onClick={() => setScope("workbook")}
          >
            Workbook
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
            Whole cell
          </button>
        </div>

        <div className="flex items-center gap-2 text-[10px] text-muted-foreground">
          {busy && <Loader2 className="size-3 animate-spin" />}
          <span>
            {query.trim()
              ? `${hits.length}${hits.length >= 5000 ? "+" : ""} result${hits.length === 1 ? "" : "s"}`
              : "Type to search"}
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
          {hits.map((hit) => (
            <button
              key={`${hit.sheet}:${hit.row}:${hit.col}`}
              type="button"
              onClick={() => onJump(hit)}
              className="rounded-sm px-2 py-1.5 text-left text-[11px] hover:bg-accent"
            >
              <div className="flex items-center gap-1.5 text-[10px] text-muted-foreground">
                <span className="rounded-sm bg-muted px-1 py-px font-mono font-medium text-foreground">
                  {cellRef(hit.row, hit.col)}
                </span>
                <span className="truncate" title={hit.sheet}>
                  {hit.sheet}
                </span>
              </div>
              <div className="mt-0.5 truncate text-foreground/90 select-text" title={hit.text}>
                {hit.text || " "}
              </div>
            </button>
          ))}
        </div>
      </ScrollArea>
    </div>
  );
}
