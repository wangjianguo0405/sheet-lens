import { GoToDialog } from "@/components/GoToDialog";
import { UnhideDialog, type UnhideAnchor } from "@/components/UnhideDialog";
import { UpdateDialog } from "@/components/UpdateDialog";
import type { UpdateStatus } from "@/lib/update";

export interface AppDialogsProps {
  gotoOpen: boolean;
  setGotoOpen: (open: boolean) => void;
  /** The sheets on the strip, for Go To's own list. */
  sheetNames: string[];
  currentSheet: string;
  onGo: (sheet: string, row: number, col: number) => void;

  unhideOpen: boolean;
  setUnhideOpen: (open: boolean) => void;
  hiddenSheets: string[];
  unhideAt: UnhideAnchor | null;
  onUnhide: (names: string[]) => void;

  updateOpen: boolean;
  setUpdateOpen: (open: boolean) => void;
  update: UpdateStatus | null;
  setUpdateAvailable: (available: boolean) => void;
}

/**
 * The window's own dialogs, mounted once and opened from wherever asks for
 * them: Go To, the sheet Unhide list, and the update check.
 */
export function AppDialogs({
  gotoOpen,
  setGotoOpen,
  sheetNames,
  currentSheet,
  onGo,
  unhideOpen,
  setUnhideOpen,
  hiddenSheets,
  unhideAt,
  onUnhide,
  updateOpen,
  setUpdateOpen,
  update,
  setUpdateAvailable,
}: AppDialogsProps) {
  return (
    <>
      <GoToDialog
        open={gotoOpen}
        onOpenChange={setGotoOpen}
        sheets={sheetNames}
        currentSheet={currentSheet}
        onGo={onGo}
      />

      <UnhideDialog
        open={unhideOpen}
        onOpenChange={setUnhideOpen}
        sheets={hiddenSheets}
        at={unhideAt}
        onUnhide={onUnhide}
      />

      <UpdateDialog
        open={updateOpen}
        onOpenChange={setUpdateOpen}
        status={update}
        onAvailableChange={setUpdateAvailable}
      />
    </>
  );
}
