import { useEffect, useState } from "react";
import { ChevronDown, ChevronUp } from "lucide-react";
import { useTranslation } from "react-i18next";

import { cn } from "@/lib/utils";

interface FormulaBarProps {
  /** The name box: `A2`, `3:3` or `B:B`. */
  reference: string;
  value: string;
}

/**
 * The value area's scrollbar, held to one width.
 *
 * Left to itself each engine draws its own — WebView2's is about 17px — and the
 * chevron above it is placed to line up with it, so a few pixels of difference
 * is what made the two look stacked slightly off rather than in a column.
 */
const VALUE_SCROLLBAR =
  "[&::-webkit-scrollbar]:w-3.5 [&::-webkit-scrollbar-thumb]:rounded-sm [&::-webkit-scrollbar-thumb]:bg-border";

/**
 * Excel's formula bar, which doubles as the only place a long value can be
 * read in full — grid cells are clipped at the column edge.
 */
export function FormulaBar({ reference, value }: FormulaBarProps) {
  const { t } = useTranslation();
  const [expanded, setExpanded] = useState(false);

  // Collapse again whenever the selection moves somewhere else.
  useEffect(() => {
    setExpanded(false);
  }, [reference]);

  /**
   * The chevron sits in a strip of its own whichever way the bar is laid out, so
   * opening the value area below it leaves the button exactly where it was
   * rather than shifting it into the middle of a taller row. The strip is as
   * wide as the scrollbar below it — see `VALUE_SCROLLBAR` — so the two read as
   * one column along the edge.
   */
  const toggle = value !== "" && (
    <button
      type="button"
      aria-label={expanded ? t("formulaBar.collapse") : t("formulaBar.expand")}
      title={expanded ? t("formulaBar.collapse") : t("formulaBar.expand")}
      onClick={() => setExpanded((open) => !open)}
      className="my-1.5 flex size-3.5 shrink-0 items-center justify-center rounded-sm text-muted-foreground hover:bg-accent"
    >
      {expanded ? <ChevronUp className="size-3.5" /> : <ChevronDown className="size-3.5" />}
    </button>
  );

  return (
    <div
      className={cn(
        // A fixed height and a fixed font size: the bar must not resize itself
        // as the selection moves, whatever size the selected cell happens to
        // use. Only the chevron adds rows — and only downwards, which is why the
        // row starts its children at the top rather than stretching them.
        //
        // `overflow-hidden` is load-bearing: `truncate` clips the span, but the
        // span's own height follows its content, so without this a value with a
        // newline in it grew the row instead of being cut off at one line.
        //
        // The bar is chrome, in the grey the rest of the chrome uses, and the
        // two fields on it are white. Opening the value area is then the bar
        // showing through underneath the name box, which is what Excel does.
        //
        // `pr-1.5` keeps the fields off the right edge; what shows in that gap
        // is the bar's own grey, which is also what frames them.
        "flex shrink-0 items-start overflow-hidden border-b border-border bg-muted pr-1.5 text-xs leading-tight select-text",
        !expanded && "h-7",
      )}
    >
      <div
        // The name box is a field of its own, and keeps its height so the
        // reference stays put when the value grows past it.
        className="flex h-7 w-24 shrink-0 items-center justify-center border-r border-border bg-background px-2 font-mono"
        title={reference}
      >
        <span className="truncate">{reference}</span>
      </div>

      {/*
        The value's field, white like the name box and no taller than its own
        content, so the bar's grey is what fills whatever the two of them leave.
      */}
      <div
        className={cn(
          "flex min-w-0 flex-1 flex-col bg-background",
          !expanded && "h-7",
        )}
      >
        {expanded && (
          // `self-end` puts the strip against the right edge: a column's cross
          // axis starts at the left, and without it the chevron jumps there.
          <div className="flex h-7 w-3.5 shrink-0 items-center justify-center self-end">
            {toggle}
          </div>
        )}
        {/*
          The two branches of the value are separate elements on purpose:
          `truncate` and `whitespace-pre-wrap` both set `white-space`, and as
          classes on one span the later one in the stylesheet won — which broke
          the single line the collapsed bar is supposed to show.
        */}
        <div className="flex min-w-0 flex-1 items-start">
          <span
            className={cn(
              "min-w-0 flex-1 pl-2",
              expanded
                ? cn(
                    "max-h-24 overflow-y-auto break-all whitespace-pre-wrap pb-1.5",
                    VALUE_SCROLLBAR,
                  )
                : "truncate py-1.5",
            )}
          >
            {value}
          </span>
          {!expanded && toggle}
        </div>
      </div>
    </div>
  );
}
