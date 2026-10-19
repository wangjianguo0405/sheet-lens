import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { ChevronDown, ChevronUp } from "lucide-react";
import { useTranslation } from "react-i18next";

import { cn } from "@/lib/utils";
import { parseTarget, type Target } from "@/lib/column";

interface FormulaBarProps {
  /** The name box: `A2`, `3:3` or `B:B`. */
  reference: string;
  value: string;
  /**
   * A reference typed into the name box, which is the way Excel moves the
   * selection: `C16`, or `Sheet1!C16` to go to another sheet with it.
   */
  onGo?: (target: Target) => void;
}

/**
 * Excel's formula bar, which doubles as the only place a long value can be
 * read in full — grid cells are clipped at the column edge.
 */
export function FormulaBar({ reference, value, onGo }: FormulaBarProps) {
  const { t } = useTranslation();
  const [expanded, setExpanded] = useState(false);
  const valueRef = useRef<HTMLSpanElement>(null);
  /** The value area's own scrollbar, measured rather than styled. */
  const [barWidth, setBarWidth] = useState(0);
  /**
   * What is being typed into the name box, or null to show the selection.
   * Kept apart from `reference` so a half-typed reference is not the selection
   * and Escape can put back what is.
   */
  const [typed, setTyped] = useState<string | null>(null);

  // Collapse again whenever the selection moves somewhere else.
  useEffect(() => {
    setExpanded(false);
  }, [reference]);

  /**
   * The chevron above the value area is placed to line up with the scrollbar
   * below it, and left to itself each engine draws a different width — so the
   * strip takes the width the engine actually used, measured off the span.
   *
   * Laid out rather than deferred: the two are side by side, and one frame of
   * the wrong width is one frame of them not lining up.
   */
  useLayoutEffect(() => {
    const el = valueRef.current;
    if (!el) return;
    setBarWidth(el.offsetWidth - el.clientWidth);
  }, [expanded, value]);

  /**
   * The chevron sits in a strip of its own whichever way the bar is laid out, so
   * opening the value area below it leaves the button exactly where it was
   * rather than shifting it into the middle of a taller row.
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
        // reference stays put when the value grows past it. It is an input
        // rather than a label because Excel's is: type a reference into it and
        // the selection goes there.
        className="flex h-7 w-24 shrink-0 items-center border-r border-border bg-background px-1 font-mono"
        title={reference}
      >
        <input
          value={typed ?? reference}
          aria-label={t("goTo.title")}
          spellCheck={false}
          onChange={(event) => setTyped(event.target.value)}
          onFocus={(event) => {
            // All of it selected, so typing replaces the reference rather than
            // appending to it — which is what Excel does when it is clicked.
            setTyped(reference);
            event.currentTarget.select();
          }}
          onBlur={() => setTyped(null)}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              const target = parseTarget(event.currentTarget.value);
              setTyped(null);
              // Unreadable or unknown, and the box goes back to saying where
              // the selection actually is; there is nothing to report.
              if (target) onGo?.(target);
              event.currentTarget.blur();
            } else if (event.key === "Escape") {
              setTyped(null);
              event.currentTarget.blur();
            }
          }}
          className="w-full min-w-0 bg-transparent text-center text-xs outline-none"
        />
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
          <div
            style={barWidth ? { width: barWidth } : undefined}
            className="flex h-7 w-3.5 shrink-0 items-center justify-center self-end"
          >
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
            ref={valueRef}
            className={cn(
              "min-w-0 flex-1 pl-2",
              expanded
                ? "max-h-24 overflow-y-auto break-all whitespace-pre-wrap pb-1.5"
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
