import { Plus, X } from "lucide-react";
import type { MouseEvent as ReactMouseEvent } from "react";
import { useTranslation } from "react-i18next";

import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";
import { copyText } from "@/lib/clipboard";
import { cn } from "@/lib/utils";

/**
 * Middle-click closes a tab, the gesture every browser puts on one. It listens
 * for `auxclick` rather than `mousedown` so that a left press still selects the
 * tab on release, and it takes the whole tab rather than just its × — the name
 * is the part being aimed at.
 */
function middleClose(close: () => void) {
  return (event: ReactMouseEvent) => {
    if (event.button !== 1) return;
    close();
  };
}

export interface FileTab {
  path: string;
  fileName: string;
}

interface FileTabsProps {
  tabs: FileTab[];
  active: string | null;
  onSelect: (path: string) => void;
  onClose: (path: string) => void;
  /** Whether the start page has a tab of its own at the moment. */
  homeOpen: boolean;
  /** Whether that tab is the one showing. */
  atHome: boolean;
  /** Show the start page. */
  onHome: () => void;
  /** Close its tab, which goes back to the workbook that was showing. */
  onCloseHome: () => void;
  /**
   * The "+" at the end of the strip. It opens the picker when the start page is
   * already showing, and opens the start page otherwise.
   */
  onNew: () => void;
}

export function FileTabs({
  tabs,
  active,
  onSelect,
  onClose,
  homeOpen,
  atHome,
  onHome,
  onCloseHome,
  onNew,
}: FileTabsProps) {
  const { t } = useTranslation();

  if (tabs.length === 0 && !homeOpen) return null;

  return (
    <div className="flex shrink-0 items-stretch overflow-x-auto border-b border-border bg-muted/40">
      {tabs.map((tab) => {
        const isActive = tab.path === active;
        return (
          <ContextMenu key={tab.path}>
            <ContextMenuTrigger asChild>
              <div
                className={cn(
                  "group flex max-w-64 shrink-0 items-center gap-1 border-r border-border px-2 py-1 text-xs",
                  isActive
                    ? "bg-background text-foreground"
                    : "text-muted-foreground hover:bg-accent/40",
                )}
                onAuxClick={middleClose(() => onClose(tab.path))}
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
            </ContextMenuTrigger>
            {/* The name and the path are the two things worth having somewhere
                else — a bug report, a chat message — and neither is easy to get
                at otherwise, since the tab shows only the name and truncates
                even that. */}
            <ContextMenuContent className="min-w-44">
              <ContextMenuItem
                className="text-xs"
                onSelect={() => void copyText(tab.fileName)}
              >
                {t("tabs.copyName")}
              </ContextMenuItem>
              <ContextMenuItem
                className="text-xs"
                onSelect={() => void copyText(tab.path)}
              >
                {t("tabs.copyPath")}
              </ContextMenuItem>
            </ContextMenuContent>
          </ContextMenu>
        );
      })}

      {/*
        The start page, as a tab of its own. It exists only once the "+" has
        been pressed, and its own × puts the workbook that was showing back.
      */}
      {homeOpen && (
        <div
          className={cn(
            "group flex shrink-0 items-center gap-1 border-r border-border px-2 py-1 text-xs",
            atHome ? "bg-background text-foreground" : "text-muted-foreground hover:bg-accent/40",
          )}
          onAuxClick={middleClose(onCloseHome)}
        >
          <button type="button" className="truncate" onClick={onHome}>
            {t("tabs.new")}
          </button>
          <button
            type="button"
            aria-label={t("tabs.close", { name: t("tabs.new") })}
            className={cn(
              "rounded-sm p-0.5 transition-opacity hover:bg-accent",
              atHome ? "opacity-60" : "opacity-0 group-hover:opacity-60",
            )}
            onClick={onCloseHome}
          >
            <X className="size-3" />
          </button>
        </div>
      )}

      <button
        type="button"
        aria-label={t("tabs.openAnother")}
        title={t("tabs.openAnother")}
        onClick={onNew}
        className="flex shrink-0 items-center px-2 text-muted-foreground hover:bg-accent/40 hover:text-foreground"
      >
        <Plus className="size-3.5" />
      </button>
    </div>
  );
}
