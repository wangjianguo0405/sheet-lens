import { useEffect, useState } from "react";
import { Trans, useTranslation } from "react-i18next";

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
import { parseTarget } from "@/lib/column";

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
  const { t } = useTranslation();
  const [value, setValue] = useState("");
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (open) {
      setValue("");
      setError(null);
    }
  }, [open]);

  const submit = () => {
    if (!value.trim()) return;
    const target = parseTarget(value);
    if (!target) {
      setError(t("goTo.invalid"));
      return;
    }

    let sheet = currentSheet;
    if (target.sheet) {
      const match = sheets.find((name) => name === target.sheet);
      if (!match) {
        setError(t("goTo.noSheet", { name: target.sheet }));
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
          <DialogTitle className="text-sm">{t("goTo.title")}</DialogTitle>
          <DialogDescription className="text-xs">
            <Trans
              i18nKey="goTo.description"
              values={{ ref: "Summary!A1" }}
              components={{ code: <code /> }}
            />
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
            {t("goTo.go")}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
