import { useEffect, useState } from "react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
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
  onUnhide: (sheet: string) => void;
}

/**
 * Excel's sheet dialog, reached from a tab's own menu because the hidden sheet
 * has no tab of its own to right-click.
 *
 * It lists what the file hides and what the reader hid alike: to whoever is
 * looking at the tab row they are the same thing, which is where this is opened
 * from. Nothing it does is written back — the flag it clears belongs to this
 * window, and reopening the workbook puts the file's own answer back.
 */
export function UnhideDialog({
  open,
  onOpenChange,
  sheets,
  at,
  onUnhide,
}: UnhideDialogProps) {
  const { t } = useTranslation();
  const [picked, setPicked] = useState<string | null>(null);

  /** Opens on the last name in the list, chosen and ready to confirm. */
  useEffect(() => {
    if (open) setPicked(sheets[sheets.length - 1] ?? null);
  }, [open, sheets]);

  const confirm = () => {
    if (picked) onUnhide(picked);
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

        <div
          role="listbox"
          aria-label={t("sheets.unhideTitle")}
          className="overflow-y-auto rounded border border-border"
          style={{ height: ROW_H * ROWS }}
        >
          {sheets.map((name) => (
            <button
              key={name}
              type="button"
              role="option"
              aria-selected={name === picked}
              onClick={() => setPicked(name)}
              onDoubleClick={confirm}
              className={cn(
                "block w-full truncate px-2 text-left text-xs",
                // The same blue a search hit is walked to in, so a list of
                // names selects the same way in both panels.
                name === picked
                  ? "bg-[#0F6CBD] text-white"
                  : "hover:bg-accent/40",
              )}
              style={{ height: ROW_H, lineHeight: `${ROW_H}px` }}
            >
              {name}
            </button>
          ))}
        </div>

        <DialogFooter>
          <Button
            size="sm"
            className="h-7 text-xs"
            disabled={!picked}
            onClick={confirm}
          >
            {t("sheets.unhideGo")}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
