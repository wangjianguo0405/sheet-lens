import { useEffect, useState } from "react";
import { ChevronDown, ChevronUp } from "lucide-react";

import { cn } from "@/lib/utils";

interface FormulaBarProps {
  /** The name box: `A2`, `3:3` or `B:B`. */
  reference: string;
  value: string;
}

/**
 * Excel's formula bar, which doubles as the only place a long value can be
 * read in full — grid cells are clipped at the column edge.
 */
export function FormulaBar({ reference, value }: FormulaBarProps) {
  const [expanded, setExpanded] = useState(false);

  // Collapse again whenever the selection moves somewhere else.
  useEffect(() => {
    setExpanded(false);
  }, [reference]);

  return (
    <div
      className={cn(
        // A fixed height and a fixed font size: the bar must not resize itself
        // as the selection moves, whatever size the selected cell happens to
        // use. Only the chevron adds rows.
        "flex shrink-0 items-stretch border-b border-border bg-background text-xs leading-tight select-text",
        !expanded && "h-7",
      )}
    >
      <div
        className="flex w-24 shrink-0 items-center justify-center border-r border-border px-2 font-mono"
        title={reference}
      >
        <span className="truncate">{reference}</span>
      </div>

      <div className="flex min-w-0 flex-1 items-center gap-1 px-2">
        <span
          className={cn(
            "min-w-0 flex-1",
            expanded
              ? "max-h-24 overflow-y-auto break-all whitespace-pre-wrap"
              : "truncate whitespace-pre",
          )}
        >
          {value}
        </span>
        {value !== "" && (
          <button
            type="button"
            aria-label={expanded ? "Collapse" : "Expand"}
            title={expanded ? "Collapse" : "Expand"}
            onClick={() => setExpanded((open) => !open)}
            className="shrink-0 rounded-sm p-0.5 text-muted-foreground hover:bg-accent"
          >
            {expanded ? (
              <ChevronUp className="size-3.5" />
            ) : (
              <ChevronDown className="size-3.5" />
            )}
          </button>
        )}
      </div>
    </div>
  );
}
