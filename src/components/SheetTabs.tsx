import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import { ChevronLeft, ChevronRight } from "lucide-react";
import { useTranslation } from "react-i18next";

import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuLabel,
  ContextMenuRadioGroup,
  ContextMenuRadioItem,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";
import type { SheetInfo } from "@/lib/excel";
import { cn } from "@/lib/utils";

interface SheetTabsProps {
  sheets: SheetInfo[];
  active: string;
  onSelect: (name: string) => void;
  /** Drawn at the right end of the row: the sheet's own scrollbar. */
  trailing?: ReactNode;
}

/**
 * The sheet name, in whichever of black and white survives the tab's colour —
 * which is what Excel does, and what decides whether 入出力定義's light green
 * takes black type or 変更履歴's blue takes white. The weights are the usual
 * luma ones, so a colour lands on the side its brightness says.
 */
function readableOn(color: string): string {
  const channel = (at: number) => parseInt(color.slice(at, at + 2), 16);
  const luma =
    0.299 * channel(1) + 0.587 * channel(3) + 0.114 * channel(5);
  return luma > 128 ? "#000000" : "#ffffff";
}

/**
 * The active tab is drawn the way Windows draws a folder tab rather than filled
 * like the rest: a pale wash of the sheet's colour, brightest at the top, with
 * Excel's green rule under the name. A sheet with no colour of its own gets the
 * same shape in plain white, which is what makes an uncoloured strip readable.
 */
function activeWash(color: string): string {
  return `linear-gradient(180deg, color-mix(in srgb, ${color} 2%, #fff), color-mix(in srgb, ${color} 45%, #fff))`;
}

/**
 * Excel's sheet strip along the bottom of the sheet area: two arrows to walk
 * the tabs, and the tabs themselves. What sits at the right end is the sheet's
 * scrollbar, not this row's — see `SheetScrollBar`.
 *
 * The arrows move the tabs, not the selection, which is what they do in Excel:
 * stepping the active sheet is Ctrl+PageUp/PageDown, and that scrolls the
 * strip by itself when it has to. The strip keeps its own scrollbar hidden —
 * a press on the tabs is a press on a sheet, and a bar there would be read as
 * the sheet's.
 */
export function SheetTabs({ sheets, active, onSelect, trailing }: SheetTabsProps) {
  const { t } = useTranslation();
  const strip = useRef<HTMLDivElement>(null);
  /** Where the tabs are scrolled to, which decides what the arrows can do. */
  const [bar, setBar] = useState({ left: 0, view: 0, total: 0 });

  const sync = useCallback(() => {
    const node = strip.current;
    if (!node) return;
    setBar({
      left: node.scrollLeft,
      view: node.clientWidth,
      total: node.scrollWidth,
    });
  }, []);

  useEffect(() => {
    const el = strip.current;
    if (!el) return;
    const observer = new ResizeObserver(sync);
    observer.observe(el);
    sync();
    return () => observer.disconnect();
  }, [sync]);

  // Another workbook brings tabs of other widths, and the strip's own box does
  // not change when it does — so the observer above never hears about it.
  useEffect(sync, [sheets, sync]);

  /**
   * A sheet the user never saw is a sheet they did not pick, so a switch from
   * elsewhere — Ctrl+PageDown, a search hit, Go To — scrolls its tab into view.
   */
  useEffect(() => {
    const tab = strip.current?.querySelector<HTMLElement>(
      `[data-sheet="${CSS.escape(active)}"]`,
    );
    tab?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [active, sheets]);

  /** The tabs' left edges in scroll coordinates, which the arrows step between. */
  const edges = useCallback(() => {
    const el = strip.current;
    if (!el) return [];
    const origin = el.getBoundingClientRect().left - el.scrollLeft;
    return [...el.querySelectorAll<HTMLElement>("[data-sheet]")].map((tab) =>
      Math.round(tab.getBoundingClientRect().left - origin),
    );
  }, []);

  const overflowing = bar.total > bar.view + 1;
  const atStart = bar.left <= 0;
  const atEnd = bar.left + bar.view >= bar.total - 1;

  const step = useCallback(
    (direction: 1 | -1) => {
      const node = strip.current;
      if (!node) return;
      const stops = edges();
      // The nearest tab edge strictly off to that side, so a press always
      // moves the strip by at least one tab rather than a nudge.
      const target =
        direction > 0
          ? (stops.find((edge) => edge > node.scrollLeft + 1) ?? node.scrollWidth)
          : ([...stops].reverse().find((edge) => edge < node.scrollLeft - 1) ?? 0);
      node.scrollTo({ left: target, behavior: "smooth" });
    },
    [edges],
  );

  const arrowClass = (enabled: boolean) =>
    cn(
      // A disabled arrow still belongs to the right-click list, so it must not
      // hold the pointer: the press has to reach the box around it.
      "flex items-center border-border px-1.5 disabled:pointer-events-none",
      enabled
        ? "text-muted-foreground hover:bg-accent/40 hover:text-foreground"
        : "text-muted-foreground/30",
    );

  return (
    <div className="flex shrink-0 items-stretch border-t border-border bg-muted/40">
      {/*
        Excel's sheet list hangs off a right-click on the tab arrows: with more
        sheets than fit, it is how you reach one without walking the arrows.
      */}
      <ContextMenu>
        <ContextMenuTrigger asChild>
          <div className="flex shrink-0 items-stretch">
            <button
              type="button"
              aria-label={t("sheets.scrollLeft")}
              title={t("sheets.scrollLeft")}
              disabled={!overflowing || atStart}
              onClick={() => step(-1)}
              className={cn(arrowClass(overflowing && !atStart), "border-r")}
            >
              <ChevronLeft className="size-3.5" />
            </button>
            <button
              type="button"
              aria-label={t("sheets.scrollRight")}
              title={t("sheets.scrollRight")}
              disabled={!overflowing || atEnd}
              onClick={() => step(1)}
              className={cn(arrowClass(overflowing && !atEnd), "border-r")}
            >
              <ChevronRight className="size-3.5" />
            </button>
          </div>
        </ContextMenuTrigger>
        <ContextMenuContent className="max-h-80 min-w-40">
          <ContextMenuLabel className="text-[10px]">
            {t("sheets.list")}
          </ContextMenuLabel>
          <ContextMenuRadioGroup value={active} onValueChange={onSelect}>
            {sheets.map((sheet) => (
              <ContextMenuRadioItem
                key={sheet.name}
                value={sheet.name}
                className="text-xs"
              >
                <span className="truncate">{sheet.name}</span>
              </ContextMenuRadioItem>
            ))}
          </ContextMenuRadioGroup>
        </ContextMenuContent>
      </ContextMenu>

      <div
        ref={strip}
        onScroll={sync}
        className="flex min-w-24 flex-1 items-stretch overflow-x-auto [scrollbar-width:none] [&::-webkit-scrollbar]:hidden"
      >
        {sheets.map((sheet) => {
          const isActive = sheet.name === active;
          return (
            <button
              key={sheet.name}
              type="button"
              data-sheet={sheet.name}
              onClick={() => onSelect(sheet.name)}
              title={t("sheets.tabTitle", {
                name: sheet.name,
                rows: sheet.rows,
                cols: sheet.cols,
              })}
              className={cn(
                "shrink-0 border-r border-b-2 border-border px-3 py-1 text-xs whitespace-nowrap transition",
                isActive
                  ? "border-b-[#107C41] bg-background font-medium text-foreground"
                  : "border-b-transparent text-muted-foreground hover:bg-accent/40 hover:text-foreground",
                // A tab filled with its own colour has the grey the others take
                // on hover sitting behind an inline fill that wins, so it
                // lightens instead — the same "this one is under the pointer".
                !isActive && sheet.tabColor && "hover:brightness-110",
              )}
              // A sheet the workbook gives a tab colour is filled with it, name
              // and all — Excel paints the whole tab, not a strip along its
              // foot. The current sheet keeps the raised tab instead, washed
              // with that same colour, so it still says which sheet it is
              // without going solid like the others.
              style={
                sheet.tabColor
                  ? isActive
                    ? { backgroundImage: activeWash(sheet.tabColor) }
                    : {
                        backgroundColor: sheet.tabColor,
                        color: readableOn(sheet.tabColor),
                      }
                  : undefined
              }
            >
              {sheet.name}
            </button>
          );
        })}
      </div>

      {trailing}
    </div>
  );
}
