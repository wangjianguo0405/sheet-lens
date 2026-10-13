import {
  Download,
  FolderOpen,
  Languages,
  PanelBottom,
  PanelLeft,
} from "lucide-react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuShortcut,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { LOCALES, setLocale, type Locale } from "@/lib/i18n";
import { fileName, RECENT_SHOWN } from "@/lib/recent";
import { cn } from "@/lib/utils";

interface ToolbarProps {
  onOpen: () => void;
  onToggleSidebar: () => void;
  onToggleSheetTabs: () => void;
  sidebarOpen: boolean;
  sheetTabsOpen: boolean;
  hasWorkbook: boolean;
  /** Newest first. The File menu shows the first few and puts the rest under More. */
  recentFiles: string[];
  /** Those of them the last attempt to open failed on, which are shown greyed. */
  missingFiles: string[];
  onOpenRecent: (path: string) => void;
  onCloseWorkbook: () => void;
  /** False for a build that cannot replace itself, which hides the entry. */
  canUpdate: boolean;
  updateAvailable: boolean;
  onShowUpdate: () => void;
}

export function Toolbar({
  onOpen,
  onToggleSidebar,
  onToggleSheetTabs,
  sidebarOpen,
  sheetTabsOpen,
  hasWorkbook,
  recentFiles,
  missingFiles,
  onOpenRecent,
  onCloseWorkbook,
  canUpdate,
  updateAvailable,
  onShowUpdate,
}: ToolbarProps) {
  const { t, i18n } = useTranslation();

  const missing = new Set(missingFiles);
  const recent = (paths: string[]) =>
    paths.map((path) => (
      <DropdownMenuItem
        key={path}
        className="text-xs"
        onSelect={() => onOpenRecent(path)}
      >
        <span
          // Greyed, but not disabled: a file that has gone is often a file that
          // is coming back, and a click is how it gets another try.
          className={cn("truncate", missing.has(path) && "text-muted-foreground")}
          title={missing.has(path) ? t("toolbar.missing", { path }) : path}
        >
          {fileName(path)}
        </span>
      </DropdownMenuItem>
    ));

  return (
    <div className="flex shrink-0 items-center gap-1 border-b border-border px-2 py-1.5">
      {/* Decorative: the wordmark beside it already names the app, so it repeats
          nothing for a screen reader. `draggable` off, or the webview starts a
          native drag of the bitmap when the bar is dragged. */}
      <img src="/icon.png" alt="" draggable={false} className="mr-1.5 size-5" />
      <span className="mr-2 text-xs font-semibold tracking-wide">SheetLens</span>

      {/* One button per sheet picker, and each stands on its own: either one
          may be showing whether or not the other is, so neither appears or
          disappears as the other is pressed. */}
      <Button
        size="sm"
        variant={sidebarOpen ? "secondary" : "ghost"}
        className="h-7 px-2"
        disabled={!hasWorkbook}
        onClick={onToggleSidebar}
        aria-pressed={sidebarOpen}
        title={
          sidebarOpen ? t("toolbar.hideSheetList") : t("toolbar.showSheetList")
        }
      >
        <PanelLeft className="size-3.5" />
      </Button>

      <Button
        size="sm"
        variant={sheetTabsOpen ? "secondary" : "ghost"}
        className="h-7 px-2"
        disabled={!hasWorkbook}
        onClick={onToggleSheetTabs}
        aria-pressed={sheetTabsOpen}
        title={
          sheetTabsOpen ? t("toolbar.hideSheetTabs") : t("toolbar.showSheetTabs")
        }
      >
        <PanelBottom className="size-3.5" />
      </Button>

      {/* The File menu. Everything in it is about the workbook as a whole
          rather than about where the cursor is, which is why none of it lives
          in the grid's own menus. No icon on the trigger: a folder means "open
          something", which is only the first of the three things in here. */}
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <Button size="sm" variant="ghost" className="h-7 px-2 text-xs">
            {t("toolbar.file")}
          </Button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="min-w-52">
          <DropdownMenuItem className="text-xs" onSelect={onOpen}>
            <FolderOpen className="size-3.5" />
            {t("toolbar.open")}
            <DropdownMenuShortcut>Ctrl+O</DropdownMenuShortcut>
          </DropdownMenuItem>

          {/* Only when there is something to list: an empty stretch between two
              rules reads as a menu that failed to load. "More" in turn only
              earns its chevron once its submenu would hold anything. */}
          {recentFiles.length > 0 && (
            <>
              <DropdownMenuSeparator />
              {recent(recentFiles.slice(0, RECENT_SHOWN))}
              {recentFiles.length > RECENT_SHOWN && (
                <DropdownMenuSub>
                  <DropdownMenuSubTrigger className="text-xs">
                    {t("toolbar.more")}
                  </DropdownMenuSubTrigger>
                  <DropdownMenuSubContent className="min-w-52">
                    {recent(recentFiles.slice(RECENT_SHOWN))}
                  </DropdownMenuSubContent>
                </DropdownMenuSub>
              )}
            </>
          )}

          <DropdownMenuSeparator />
          <DropdownMenuItem
            className="text-xs"
            disabled={!hasWorkbook}
            onSelect={onCloseWorkbook}
          >
            {t("toolbar.close")}
            <DropdownMenuShortcut>Ctrl+W</DropdownMenuShortcut>
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>

      <div className="ml-auto flex items-center gap-1">
        {canUpdate && (
          <Button
            size="sm"
            variant={updateAvailable ? "secondary" : "ghost"}
            className="relative h-7 px-2"
            onClick={onShowUpdate}
            title={
              updateAvailable
                ? t("toolbar.updateAvailable")
                : t("toolbar.checkUpdates")
            }
          >
            <Download className="size-3.5" />
            {updateAvailable && (
              <span className="absolute -top-0.5 -right-0.5 size-1.5 rounded-full bg-[#107C41]" />
            )}
          </Button>
        )}

        {/* Each entry is named in its own language, whatever the UI is in. */}
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button
              size="sm"
              variant="ghost"
              className="h-7 px-2"
              title={t("toolbar.language")}
            >
              <Languages className="size-3.5" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" className="min-w-36">
            <DropdownMenuLabel className="text-[10px]">
              {t("toolbar.language")}
            </DropdownMenuLabel>
            <DropdownMenuRadioGroup
              value={i18n.language}
              onValueChange={(value) => setLocale(value as Locale)}
            >
              {LOCALES.map((locale) => (
                <DropdownMenuRadioItem
                  key={locale}
                  value={locale}
                  className="text-xs"
                >
                  {t(`language.${locale}`)}
                </DropdownMenuRadioItem>
              ))}
            </DropdownMenuRadioGroup>
          </DropdownMenuContent>
        </DropdownMenu>
      </div>
    </div>
  );
}
