import { Ban } from "lucide-react";
import { useTranslation } from "react-i18next";

import {
  ContextMenuItem,
  ContextMenuLabel,
  ContextMenuSeparator,
} from "@/components/ui/context-menu";
import {
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
} from "@/components/ui/dropdown-menu";
import { MARK_SOFT, MARK_STANDARD } from "@/lib/marks";
import { cn } from "@/lib/utils";

interface MarkPaletteProps {
  /**
   * Which menu this is going into. The two Radix menus are separate primitives
   * with the same shape, so the palette is written once and handed whichever
   * item, label and rule belong to the menu it is in.
   */
  variant: "dropdown" | "context";
  /** The colour the paint bucket is holding, ringed so it can be found. */
  current: string | null;
  onPick: (fill: string | null) => void;
}

/**
 * The colours a mark can take, laid out the way Excel's fill palette is: No
 * Fill, then a row of standard colours over a row of lighter ones.
 */
export function MarkPalette({ variant, current, onPick }: MarkPaletteProps) {
  const { t } = useTranslation();
  const Item = variant === "dropdown" ? DropdownMenuItem : ContextMenuItem;
  const Label = variant === "dropdown" ? DropdownMenuLabel : ContextMenuLabel;
  const Rule =
    variant === "dropdown" ? DropdownMenuSeparator : ContextMenuSeparator;

  const swatches = (colours: string[]) => (
    // Fixed tracks, not `grid-cols-10`: an `fr` track may be narrower than the
    // swatch, and the swatch then spills over the gap between them until the
    // row reads as one unbroken band. A fixed track cannot be squeezed, so the
    // palette widens to hold ten of them and the gaps stay gaps.
    <div className="grid grid-cols-[repeat(10,1rem)] justify-between gap-y-1 px-1.5 py-1.5">
      {colours.map((fill) => (
        <Item
          key={fill}
          title={fill}
          aria-label={fill}
          onSelect={() => onPick(fill)}
          style={{ background: fill }}
          className={cn(
            "size-4 rounded-[3px] p-0 ring-1 ring-foreground/15 ring-inset",
            "focus:ring-2 focus:ring-foreground/60",
            current === fill && "ring-2 ring-foreground/70",
          )}
        />
      ))}
    </div>
  );

  return (
    <>
      <Item className="text-xs" onSelect={() => onPick(null)}>
        <Ban className="size-3.5" />
        {t("mark.clear")}
      </Item>
      <Rule />
      <Label className="text-[10px]">{t("mark.standard")}</Label>
      {swatches(MARK_STANDARD)}
      <Label className="text-[10px]">{t("mark.soft")}</Label>
      {swatches(MARK_SOFT)}
    </>
  );
}
