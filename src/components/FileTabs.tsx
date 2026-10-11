import { X } from "lucide-react";
import { useTranslation } from "react-i18next";

import { cn } from "@/lib/utils";

export interface FileTab {
  path: string;
  fileName: string;
}

interface FileTabsProps {
  tabs: FileTab[];
  active: string | null;
  onSelect: (path: string) => void;
  onClose: (path: string) => void;
}

export function FileTabs({ tabs, active, onSelect, onClose }: FileTabsProps) {
  const { t } = useTranslation();

  if (tabs.length === 0) return null;

  return (
    <div className="flex shrink-0 items-stretch overflow-x-auto border-b border-border bg-muted/40">
      {tabs.map((tab) => {
        const isActive = tab.path === active;
        return (
          <div
            key={tab.path}
            className={cn(
              "group flex max-w-64 shrink-0 items-center gap-1 border-r border-border px-2 py-1 text-xs",
              isActive
                ? "bg-background text-foreground"
                : "text-muted-foreground hover:bg-accent/40",
            )}
          >
            <button
              type="button"
              className="truncate"
              title={tab.path}
              onClick={() => onSelect(tab.path)}
            >
              {tab.fileName}
            </button>
            <button
              type="button"
              aria-label={t("tabs.close", { name: tab.fileName })}
              className={cn(
                "rounded-sm p-0.5 transition-opacity hover:bg-accent",
                isActive ? "opacity-60" : "opacity-0 group-hover:opacity-60",
              )}
              onClick={() => onClose(tab.path)}
            >
              <X className="size-3" />
            </button>
          </div>
        );
      })}
    </div>
  );
}
