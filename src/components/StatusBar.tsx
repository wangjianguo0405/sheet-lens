import { Minus, Plus } from "lucide-react";

import { Separator } from "@/components/ui/separator";
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
  /** View scale as a percentage, the way Excel names it. */
  zoom: number;
  onZoom: (percent: number) => void;
}

export function StatusBar({
  fileName,
  sheet,
  cell,
  kind,
  value,
  size,
  zoom,
  onZoom,
}: StatusBarProps) {
  return (
    <div className="flex shrink-0 items-center gap-2 border-t border-border bg-muted/40 px-3 py-1 text-[11px] text-muted-foreground">
      <span className="truncate" title={fileName ?? undefined}>
        {fileName ?? "No file"}
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
          {kind && kind !== "empty" && (
            <span className="rounded-sm bg-muted px-1 py-px text-[10px]">
              {kind}
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
            aria-label="Zoom out"
            title="Zoom out"
            disabled={zoom <= MIN_ZOOM}
            onClick={() => onZoom(stepZoom(zoom, -1))}
            className="rounded-sm p-0.5 hover:bg-accent hover:text-foreground disabled:pointer-events-none disabled:opacity-40"
          >
            <Minus className="size-3" />
          </button>
          <input
            type="range"
            aria-label="Zoom"
            min={0}
            max={ZOOM_STOPS.length - 1}
            step={1}
            value={zoomIndex(zoom)}
            onChange={(event) => onZoom(ZOOM_STOPS[Number(event.target.value)])}
            className="h-1 w-24 cursor-pointer accent-[#217346]"
          />
          <button
            type="button"
            aria-label="Zoom in"
            title="Zoom in"
            disabled={zoom >= MAX_ZOOM}
            onClick={() => onZoom(stepZoom(zoom, 1))}
            className="rounded-sm p-0.5 hover:bg-accent hover:text-foreground disabled:pointer-events-none disabled:opacity-40"
          >
            <Plus className="size-3" />
          </button>
          <button
            type="button"
            title="Reset the zoom to 100%"
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
