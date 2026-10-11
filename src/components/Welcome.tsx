import { FileSpreadsheet } from "lucide-react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";

interface WelcomeProps {
  onOpen: () => void;
  error: string | null;
}

export function Welcome({ onOpen, error }: WelcomeProps) {
  const { t } = useTranslation();

  return (
    <div className="flex h-full w-full flex-col items-center justify-center gap-4 text-center">
      <FileSpreadsheet className="size-10 text-muted-foreground/50" />
      <div className="space-y-1">
        <h1 className="text-sm font-semibold">SheetLens</h1>
        <p className="max-w-sm text-xs text-muted-foreground">
          {t("welcome.tagline")}
        </p>
      </div>

      <Button size="sm" className="h-8 text-xs" onClick={onOpen}>
        {t("welcome.open")}
      </Button>

      <div className="space-y-0.5 text-[11px] text-muted-foreground">
        <p>
          <kbd className="rounded-sm bg-muted px-1">Ctrl</kbd> +{" "}
          <kbd className="rounded-sm bg-muted px-1">O</kbd>{" "}
          {t("welcome.kbdOpen")} ·{" "}
          <kbd className="rounded-sm bg-muted px-1">Ctrl</kbd> +{" "}
          <kbd className="rounded-sm bg-muted px-1">F</kbd>{" "}
          {t("welcome.kbdSearch")} ·{" "}
          <kbd className="rounded-sm bg-muted px-1">Ctrl</kbd> +{" "}
          <kbd className="rounded-sm bg-muted px-1">G</kbd>{" "}
          {t("welcome.kbdGoTo")}
        </p>
        <p>.xlsx · .xlsm · .xls · .xlsb · .ods</p>
      </div>

      {error && (
        <p className="max-w-md rounded-sm bg-destructive/10 px-2 py-1 text-[11px] text-destructive">
          {error}
        </p>
      )}
    </div>
  );
}
