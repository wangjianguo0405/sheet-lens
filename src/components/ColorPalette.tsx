import { useState, type ReactNode } from "react";
import { Ban } from "lucide-react";
import { useTranslation } from "react-i18next";

import {
  ContextMenuItem,
  ContextMenuLabel,
  ContextMenuSeparator,
} from "@/components/ui/context-menu";
import { Hint } from "@/components/Hint";
import {
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
} from "@/components/ui/dropdown-menu";
import {
  readRecentColors,
  RECENT_SHOWN,
  rememberColor,
  STANDARD_COLORS,
  THEME_COLORS,
} from "@/lib/palette";
import { cn } from "@/lib/utils";

interface ColorPaletteProps {
  /**
   * Which menu this is going into. The two Radix menus are separate primitives
   * with the same shape, so the palette is written once and handed whichever
   * item, label and rule belong to the menu it is in.
   */
  variant: "dropdown" | "context";
  /**
   * What the colour being picked applies to — a cell's fill, a sheet tab, a
   * review mark — and which swatch is therefore ringed.
   */
  current: string | null;
  onPick: (fill: string | null) => void;
  /**
   * What the last item says, which is the one thing the uses of this palette do
   * not share: a cell is filled and a sheet tab is coloured, and Excel names the
   * two differently.
   */
  clearLabel?: string;
}

/**
 * Excel's colour palette, laid out as Excel lays it out: the theme's colours as
 * a block of tints, the standard row under it, then the colours just used, with
 * No Fill at the foot.
 */
export function ColorPalette({ variant, current, onPick, clearLabel }: ColorPaletteProps) {
  const { t } = useTranslation();
  /**
   * Read once, as the menu opens. The palette is unmounted while its menu is
   * shut, so this is the list as it stood when the menu was last closed — which
   * is what a "recently used" row is, and why the app carries no state for it.
   */
  const [recent] = useState(readRecentColors);
  const Item = variant === "dropdown" ? DropdownMenuItem : ContextMenuItem;
  const Label = variant === "dropdown" ? DropdownMenuLabel : ContextMenuLabel;
  const Rule =
    variant === "dropdown" ? DropdownMenuSeparator : ContextMenuSeparator;

  // Keys are positions, not colours: a theme tint can land on the same hex as
  // one of the standard colours, and two children may not share a key.
  const swatch = (fill: string, key: string) => (
    <Hint key={key} label={fill}>
      <Item
        aria-label={fill}
        // Every pick is remembered, whichever menu it was made in: the row is one
        // row for the whole app, so a colour used on a cell is there when the next
        // sheet tab wants one.
        onSelect={() => {
          rememberColor(fill);
          onPick(fill);
        }}
        style={{ background: fill }}
        className={cn(
          "size-[22px] rounded-[2px] p-0 ring-1 ring-foreground/15 ring-inset",
          "focus:ring-2 focus:ring-foreground/60",
          current === fill && "ring-2 ring-foreground/70",
        )}
      />
    </Hint>
  );

  /**
   * What an unused slot in the recent row is drawn as: the same square, with no
   * colour in it and nothing to press. Excel keeps the row at ten slots whether
   * they have been used or not, and a reader looking for a colour they used a
   * moment ago is looking at a row that is where the colours will be.
   */
  const emptySlot = (at: number) => (
    <span
      key={`slot ${at}`}
      aria-hidden
      className="size-[22px] rounded-[2px] ring-1 ring-foreground/15 ring-inset"
    />
  );

  /**
   * The theme's own row: the ten colours a workbook's theme is made of, each one
   * a square of its own with a gutter beside it, which is how Excel draws the
   * row the tints hang from. The margin under it is the room Excel leaves before
   * the tints begin.
   */
  const themeSwatch = (fill: string, at: number) => (
    <Hint key={`theme ${at}`} label={fill}>
      <Item
        aria-label={fill}
        onSelect={() => {
          rememberColor(fill);
          onPick(fill);
        }}
        style={{ background: fill }}
        className={cn(
          "mb-[8px] size-[22px] rounded-[2px] p-0 ring-1 ring-foreground/15 ring-inset",
          "focus:ring-2 focus:ring-foreground/60",
          current === fill && "ring-2 ring-foreground/70",
        )}
      />
    </Hint>
  );

  /**
   * And the tints and shades under it: each one a square like the row above,
   * with the same gutter beside it, but with no border and no room *between* the
   * rows — the colours are the separations there, and a border on every one of
   * them would cut the forty into squares Excel does not draw.
   *
   * Nothing is left to ring them from the outside, either: a hovered tint has
   * its neighbours below and above painted over any ring drawn outside it, so
   * the two rings this one can take are drawn inside its own square.
   */
  const tintSwatch = (fill: string, at: number) => (
    <Hint key={`tint ${at}`} label={fill}>
      <Item
        aria-label={fill}
        onSelect={() => {
          rememberColor(fill);
          onPick(fill);
        }}
        style={{ background: fill }}
        className={cn(
          "size-[22px] rounded-none p-0",
          "focus:ring-2 focus:ring-foreground/60 focus:ring-inset",
          current === fill && "ring-2 ring-foreground/70 ring-inset",
        )}
      />
    </Hint>
  );

  /**
   * A row of ten, with Excel's own gutters between them, for the standard and
   * recent rows — the two that are a band of separate swatches rather than a
   * block of colour.
   *
   * Fixed tracks, not `grid-cols-10`: an `fr` track may be narrower than the
   * swatch, and the swatch then spills over the gutter between them until the
   * row reads as one unbroken band. A fixed track cannot be squeezed, so the
   * palette is exactly as wide as ten of them and their gutters — which is what
   * the menu around it widens to, rather than the block sitting in a menu with
   * a slack edge beside it.
   */
  const swatchRow = (slots: ReactNode[], key: number) => (
    <div
      key={key}
      className="grid w-max grid-cols-[repeat(10,22px)] gap-[3px] px-1.5"
    >
      {slots}
    </div>
  );

  return (
    <>
      <Label className="text-[10px]">{t("palette.theme")}</Label>
      {/*
        The gutters run one way only: between the columns, as they do in the
        standard row, and not between the rows of tints, which sit against one
        another.
      */}
      <div className="grid w-max grid-cols-[repeat(10,22px)] gap-x-[3px] px-1.5">
        {THEME_COLORS[0].map(themeSwatch)}
        {THEME_COLORS.slice(1).flat().map(tintSwatch)}
      </div>
      <Label className="text-[10px]">{t("palette.standard")}</Label>
      {swatchRow(
        STANDARD_COLORS.map((fill, at) => swatch(fill, `standard ${at}`)),
        -1,
      )}
      <Label className="text-[10px]">{t("palette.recent")}</Label>
      {swatchRow(
        [
          ...recent.map((fill, at) => swatch(fill, `recent ${at}`)),
          ...Array.from({ length: RECENT_SHOWN - recent.length }, (_, at) =>
            emptySlot(at),
          ),
        ],
        -2,
      )}
      <Rule />
      <Item className="text-xs" onSelect={() => onPick(null)}>
        <Ban className="size-3.5" />
        {clearLabel ?? t("palette.clear")}
      </Item>
    </>
  );
}
