import { Minus, Plus } from "lucide-react";
import { useTranslation } from "react-i18next";

import { Separator } from "@/components/ui/separator";
import type { SelectionSize } from "@/lib/selection";
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
  sheet: string | null;
  cell: string | null;
  kind: string | null;
  value: string | null;
  size: string | null;
  /** How many rows and columns the selection spans. */
  span: SelectionSize | null;
  /** View scale as a percentage, the way Excel names it. */
  zoom: number;
  onZoom: (percent: number) => void;
}

/** Each selection shape, mapped to the phrase that counts it. */
const SPAN_KEY = {
  rows: "status.spanRows",
  cols: "status.spanCols",
  block: "status.spanBlock",
} as const;

/** The cell kinds the backend reports, mapped to their translated labels. */
const KIND_KEY = {
  text: "status.kind.text",
  number: "status.kind.number",
  bool: "status.kind.bool",
  date: "status.kind.date",
  error: "status.kind.error",
} as const;

export function StatusBar({
  fileName,
  sheet,
  cell,
  kind,
  value,
  size,
  span,
  zoom,
  onZoom,
}: StatusBarProps) {
  const { t } = useTranslation();

  return (
    <div className="flex shrink-0 items-center gap-2 border-t border-border bg-muted/40 px-3 py-1 text-[11px] text-muted-foreground">
      <span className="truncate" title={fileName ?? undefined}>
        {fileName ?? t("status.noFile")}
      </span>
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
      {cell && (
        <>
          <Separator orientation="vertical" className="h-3" />
          <span className="font-medium text-foreground select-text">{cell}</span>
          {span && (
            <span className="whitespace-nowrap tabular-nums">
              {t(SPAN_KEY[span.kind], span)}
            </span>
          )}
          {kind && kind !== "empty" && (
            <span className="rounded-sm bg-muted px-1 py-px text-[10px]">
              {t(KIND_KEY[kind as keyof typeof KIND_KEY], { defaultValue: kind })}
            </span>
          )}
          <span className="min-w-0 flex-1 truncate text-foreground/80 select-text" title={value ?? ""}>
            {value}
          </span>
        </>
      )}
      {cell && (
        // Excel's zoom lives in the bottom-right corner of the window.
        <div className="ml-auto flex shrink-0 items-center gap-1.5">
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
            className="h-1 w-24 cursor-pointer accent-[#217346]"
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
