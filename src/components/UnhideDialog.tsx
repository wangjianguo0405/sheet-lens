import { useEffect, useMemo, useState } from "react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { cn } from "@/lib/utils";

/**
 * The list box is six names tall whatever it holds, which is the shape Excel's
 * own Unhide list has: a box that changed height with every workbook would be a
 * different dialog each time.
 */
const ROW_H = 20;
const ROWS = 6;

/**
 * How wide the dialog is. Fixed, because a dialog placed by hand is a dialog
 * that can be placed off the screen: the width is what says how far left of the
 * window's edge its anchor may sit, and the tab a menu was opened on can be the
 * last one along.
 */
const WIDTH = 240;

/** Where the menu that opened this was, in viewport coordinates. */
export interface UnhideAnchor {
  left: number;
  /** Distance up from the viewport's foot, so the dialog's own floor lands there. */
  bottom: number;
}

interface UnhideDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  /** The sheets off the strip, in the workbook's own order. */
  sheets: string[];
  /**
   * Stands where the menu that opened it stood: the dialog's bottom-left corner
   * on the menu's own, so it takes the place of the thing that just closed
   * rather than appearing somewhere the pointer is not.
   */
  at: UnhideAnchor | null;
  onUnhide: (sheets: string[]) => void;
}

/**
 * Excel's sheet dialog, reached from a tab's own menu because the hidden sheet
 * has no tab of its own to right-click.
 *
 * It lists what the file hides and what the reader hid alike: to whoever is
 * looking at the tab row they are the same thing, which is where this is opened
 * from. Nothing it does is written back — the flag it clears belongs to this
 * window, and reopening the workbook puts the file's own answer back.
 *
 * Excel's own dialog takes one name. This one takes several — a file can hide
 * a dozen sheets at a time — and a filter box, because the six-name box it
 * shares with Excel is not where a long list can be found by eye. The filter
 * only narrows what is shown: a name picked before the query was typed stays
 * picked, and is part of what the button confirms.
 */
export function UnhideDialog({
  open,
  onOpenChange,
  sheets,
  at,
  onUnhide,
}: UnhideDialogProps) {
  const { t } = useTranslation();
  /** In the order they were picked, so the last one becomes the current sheet. */
  const [picked, setPicked] = useState<string[]>([]);
  const [query, setQuery] = useState("");
  /** The row a Shift-click measures its range from, as an index into `sheets`. */
  const [anchor, setAnchor] = useState<number | null>(null);

  /** Opens on the last name in the list, chosen and ready to confirm. */
  useEffect(() => {
    if (!open) return;
    setPicked(sheets.length ? [sheets[sheets.length - 1]] : []);
    setQuery("");
    setAnchor(null);
  }, [open, sheets]);

  const visible = useMemo(() => {
    const needle = query.trim().toLocaleLowerCase();
    return needle
      ? sheets.filter((name) => name.toLocaleLowerCase().includes(needle))
      : sheets;
  }, [sheets, query]);

  const toggle = (name: string, shiftKey: boolean) => {
    const index = sheets.indexOf(name);
    if (shiftKey && anchor !== null) {
      const [from, to] = anchor <= index ? [anchor, index] : [index, anchor];
      const range = sheets.slice(from, to + 1);
      setPicked((prev) => [...prev, ...range.filter((one) => !prev.includes(one))]);
      return;
    }
    setAnchor(index);
    setPicked((prev) =>
      prev.includes(name) ? prev.filter((one) => one !== name) : [...prev, name],
    );
  };

  /**
   * Confirms everything picked. A double-click on a row that is not picked
   * confirms that one name alone, which is what double-clicking a list of
   * names has always done here — the row's own two clicks leave it unpicked,
   * so the intent is read from that rather than from the click count.
   */
  const confirm = (only?: string) => {
    const names = only ? [only] : picked;
    if (names.length) onUnhide(names);
    onOpenChange(false);
  };

  /**
   * The anchor itself, pulled back inside the window when the tab it came from
   * was near the right edge — the exact corner where it can be had, the visible
   * dialog where it cannot.
   */
  const placed = at
    ? { left: Math.max(8, Math.min(at.left, window.innerWidth - WIDTH - 8)), bottom: at.bottom }
    : null;

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent
        className={cn(
          "sm:max-w-60",
          // The anchor replaces the centred placement the shared dialog opens
          // with: both the offset and the transform that centres it go.
          placed && "top-auto left-0 translate-x-0 translate-y-0",
        )}
        style={placed ? { left: placed.left, bottom: placed.bottom } : undefined}
      >
        <DialogHeader>
          <DialogTitle className="text-sm">{t("sheets.unhideTitle")}</DialogTitle>
        </DialogHeader>

        <Input
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter") confirm();
          }}
          placeholder={t("sheets.unhideFilter")}
          aria-label={t("sheets.unhideFilter")}
          className="h-7 rounded border-border px-2 text-xs"
        />

        <div
          role="listbox"
          aria-label={t("sheets.unhideTitle")}
          aria-multiselectable="true"
          className="overflow-y-auto rounded border border-border"
          style={{ height: ROW_H * ROWS }}
        >
          {visible.map((name) => (
            <button
              key={name}
              type="button"
              role="option"
              aria-selected={picked.includes(name)}
              onClick={(event) => toggle(name, event.shiftKey)}
              onDoubleClick={() =>
                confirm(picked.includes(name) ? undefined : name)
              }
              className={cn(
                "block w-full truncate px-2 text-left text-xs",
                // The same blue a search hit is walked to in, so a list of
                // names selects the same way in both panels.
                picked.includes(name)
                  ? "bg-[#0F6CBD] text-white"
                  : "hover:bg-accent/40",
              )}
              style={{ height: ROW_H, lineHeight: `${ROW_H}px` }}
            >
              {name}
            </button>
          ))}
          {visible.length === 0 && (
            <p className="px-2 py-2 text-xs text-muted-foreground">
              {t("sheets.none")}
            </p>
          )}
        </div>

        <DialogFooter>
          <Button
            size="sm"
            className="h-7 text-xs"
            disabled={picked.length === 0}
            onClick={() => confirm()}
          >
            {picked.length > 1
              ? `${t("sheets.unhideGo")} (${picked.length})`
              : t("sheets.unhideGo")}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
