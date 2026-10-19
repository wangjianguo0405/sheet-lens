import { Grid3x3, Minus, Plus, TableCellsSplit } from "lucide-react";
import { useTranslation } from "react-i18next";

import { Separator } from "@/components/ui/separator";
import type { SelectionSize } from "@/lib/selection";
import { cn } from "@/lib/utils";
import type { ViewMode } from "@/lib/viewMode";
import {
  DEFAULT_ZOOM,
  MAX_ZOOM,
  MIN_ZOOM,
  ZOOM_STOPS,
  stepZoom,
  zoomIndex,
} from "@/lib/zoom";

interface StatusBarProps {
  fileName: string | null;
  /** How many sheets the workbook holds, said beside its name. */
  sheetCount: number | null;
  sheet: string | null;
  size: string | null;
  /** What the sheet's filter is hiding, or null when none is. */
  filtered: { shown: number; total: number } | null;
  /** How many rows and columns the selection spans. */
  span: SelectionSize | null;
  /** Which of Excel's two views the sheet is drawn in. */
  view: ViewMode;
  onView: (view: ViewMode) => void;
  /** View scale as a percentage, the way Excel names it. */
  zoom: number;
  onZoom: (percent: number) => void;
}

/** The two views, in the order Excel's own buttons sit in. */
const VIEWS = [
  { mode: "normal", label: "status.viewNormal", Icon: Grid3x3 },
  { mode: "pageBreak", label: "status.viewPageBreak", Icon: TableCellsSplit },
] as const;

/** Each selection shape, mapped to the phrase that counts it. */
const SPAN_KEY = {
  rows: "status.spanRows",
  cols: "status.spanCols",
  block: "status.spanBlock",
  areas: "status.spanAreas",
} as const;

export function StatusBar({
  fileName,
  sheetCount,
  sheet,
  size,
  filtered,
  span,
  view,
  onView,
  zoom,
  onZoom,
}: StatusBarProps) {
  const { t } = useTranslation();

  return (
    <div className="flex shrink-0 items-center gap-2 border-t border-border bg-muted/40 px-3 py-1 text-[11px] text-muted-foreground">
      {/*
        The bar names the workbook and the sheet and says how far the grid
        reaches, and nothing about the cell: the name box and the formula bar
        are directly above the sheet and already say which cell is current and
        what is in it.
      */}
      <span className="truncate" title={fileName ?? undefined}>
        {fileName ?? t("status.noFile")}
      </span>
      {/*
        How many sheets are inside is the one thing about the workbook the
        window does not say anywhere else: the list counts them only when it is
        open, and the tabs only ever show the ones that fit.
      */}
      {sheetCount !== null && (
        <>
          <Separator orientation="vertical" className="h-3" />
          <span className="shrink-0 whitespace-nowrap tabular-nums">
            {t("sheets.count", { count: sheetCount })}
          </span>
        </>
      )}
      {sheet && (
        <>
          <Separator orientation="vertical" className="h-3" />
          <span className="truncate" title={sheet}>
            {sheet}
          </span>
        </>
      )}
      {size && (
        <>
          <Separator orientation="vertical" className="h-3" />
          <span className="tabular-nums">{size}</span>
        </>
      )}
      {/*
        What the filter is hiding, which Excel says here too rather than in a
        band over the sheet: the rows are already gone from the grid, and a
        count of them costs the sheet a row of height to repeat that.
      */}
      {filtered && (
        <>
          <Separator orientation="vertical" className="h-3" />
          <span className="shrink-0 whitespace-nowrap tabular-nums">
            {t("filter.showing", {
              shown: filtered.shown,
              total: filtered.total,
            })}
          </span>
        </>
      )}
      {/*
        A whole run or block is the one thing those two cannot say — the name
        box addresses `A1:B2` or `3:7` and never counts it — so it sits at this
        end, beside the zoom, where Excel keeps its own figures for a selection.
      */}
      {sheet && (
        <div className="ml-auto flex shrink-0 items-center gap-1.5">
          {span && (
            <span className="mr-1.5 whitespace-nowrap tabular-nums">
              {t(SPAN_KEY[span.kind], span)}
            </span>
          )}
          {/*
            Excel's two views, at the right end of the bar as its own buttons
            are, with the zoom beyond them.
          */}
          {VIEWS.map(({ mode, label, Icon }) => (
            <button
              key={mode}
              type="button"
              aria-label={t(label)}
              aria-pressed={view === mode}
              title={t(label)}
              onClick={() => onView(mode)}
              className={cn(
                "rounded-sm p-0.5 hover:bg-accent hover:text-foreground",
                view === mode && "bg-accent text-foreground",
              )}
            >
              <Icon className="size-3.5" />
            </button>
          ))}
          <Separator orientation="vertical" className="h-3" />
          {/* Excel's zoom lives in the bottom-right corner of the window. */}
          <button
            type="button"
            aria-label={t("status.zoomOut")}
            title={t("status.zoomOut")}
            disabled={zoom <= MIN_ZOOM}
            onClick={() => onZoom(stepZoom(zoom, -1))}
            className="rounded-sm p-0.5 hover:bg-accent hover:text-foreground disabled:pointer-events-none disabled:opacity-40"
          >
            <Minus className="size-3" />
          </button>
          <input
            type="range"
            aria-label={t("status.zoom")}
            min={0}
            max={ZOOM_STOPS.length - 1}
            step={1}
            value={zoomIndex(zoom)}
            onChange={(event) => onZoom(ZOOM_STOPS[Number(event.target.value)])}
            className="h-1 w-24 cursor-pointer accent-[#107C41]"
          />
          <button
            type="button"
            aria-label={t("status.zoomIn")}
            title={t("status.zoomIn")}
            disabled={zoom >= MAX_ZOOM}
            onClick={() => onZoom(stepZoom(zoom, 1))}
            className="rounded-sm p-0.5 hover:bg-accent hover:text-foreground disabled:pointer-events-none disabled:opacity-40"
          >
            <Plus className="size-3" />
          </button>
          <button
            type="button"
            title={t("status.resetZoom")}
            onClick={() => onZoom(DEFAULT_ZOOM)}
            className="w-9 text-right tabular-nums text-foreground hover:underline"
          >
            {zoom}%
          </button>
        </div>
      )}
    </div>
  );
}
