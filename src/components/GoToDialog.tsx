import { useEffect, useState } from "react";

import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { parseRef } from "@/lib/column";

interface GoToDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  sheets: string[];
  currentSheet: string;
  onGo: (sheet: string, row: number, col: number) => void;
}

export function GoToDialog({
  open,
  onOpenChange,
  sheets,
  currentSheet,
  onGo,
}: GoToDialogProps) {
  const [value, setValue] = useState("");
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (open) {
      setValue("");
      setError(null);
    }
  }, [open]);

  const submit = () => {
    const trimmed = value.trim();
    if (!trimmed) return;

    const bang = trimmed.lastIndexOf("!");
    const sheetPart = bang >= 0 ? trimmed.slice(0, bang).replace(/^'|'$/g, "") : null;
    const refPart = bang >= 0 ? trimmed.slice(bang + 1) : trimmed;

    const target = parseRef(refPart);
    if (!target) {
      setError("Enter a reference such as G408 or Sheet1!G408.");
      return;
    }

    let sheet = currentSheet;
    if (sheetPart) {
      const match = sheets.find((name) => name === sheetPart);
      if (!match) {
        setError(`No sheet named "${sheetPart}".`);
        return;
      }
      sheet = match;
    }

    onGo(sheet, target.row, target.col);
    onOpenChange(false);
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-sm">
        <DialogHeader>
          <DialogTitle className="text-sm">Go to cell</DialogTitle>
          <DialogDescription className="text-xs">
            Jump to a cell by reference. Prefix with a sheet name to switch
            sheets, for example <code>DB抽出仕様!G408</code>.
          </DialogDescription>
        </DialogHeader>

        <Input
          autoFocus
          value={value}
          onChange={(event) => {
            setValue(event.target.value);
            setError(null);
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter") submit();
          }}
          placeholder="G408"
          className="h-8 font-mono text-xs"
        />

        {error && <p className="text-[11px] text-destructive">{error}</p>}

        <DialogFooter>
          <Button size="sm" className="h-7 text-xs" onClick={submit}>
            Go
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
