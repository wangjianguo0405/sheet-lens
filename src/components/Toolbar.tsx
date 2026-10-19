import {
  ChevronDown,
  Download,
  FolderOpen,
  Languages,
  PaintBucket,
  PanelBottom,
  PanelLeft,
  X,
} from "lucide-react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import { ColorPalette } from "@/components/ColorPalette";
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

/**
 * The File menu's width, and its "More" submenu's.
 *
 * A menu is as wide as its trigger by default, and this trigger is the word
 * "ファイル" — far too narrow for a workbook's name, so every one of them came
 * out with an ellipsis. It takes the width of the longest name instead, up to a
 * quarter of the window: a name long enough to need more than that is one whose
 * tail nobody reads anyway, and a menu half the screen wide is worse than a
 * shortened name.
 *
 * The cap is the window rather than Radix's own `content-available-width`,
 * which is measured against the box the menu hangs off rather than the screen:
 * on a submenu it resolves to the parent menu's width — far more than the room
 * actually left to the right of it — so it capped nothing at all and let a long
 * name run off the edge.
 */
const MENU_WIDTH = "w-max min-w-52 max-w-[25vw]";

interface ToolbarProps {
  onOpen: () => void;
  onToggleSidebar: () => void;
  onToggleSheetTabs: () => void;
  sidebarOpen: boolean;
  sheetTabsOpen: boolean;
  hasWorkbook: boolean;
  /** Shades the selection, or lifts the shading off it when `fill` is null. */
  onMark: (fill: string | null) => void;
  /** What the bucket fills with when it is pressed rather than its caret. */
  markColor: string;
  /** Newest first. The File menu shows the first few and puts the rest under More. */
  recentFiles: string[];
  /** Those of them the last attempt to open failed on, which are shown greyed. */
  missingFiles: string[];
  onOpenRecent: (path: string) => void;
  onForgetRecent: (path: string) => void;
  onClearRecents: () => void;
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
  onMark,
  markColor,
  recentFiles,
  missingFiles,
  onOpenRecent,
  onForgetRecent,
  onClearRecents,
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
        {/* Swallows the pointer so taking an entry out is not also opening it,
            and so the menu stays up to take out another. */}
        <button
          type="button"
          aria-label={t("toolbar.forget", { name: fileName(path) })}
          title={t("toolbar.forget", { name: fileName(path) })}
          onPointerDown={(event) => {
            event.preventDefault();
            event.stopPropagation();
          }}
          onPointerUp={(event) => event.stopPropagation()}
          onClick={(event) => {
            event.stopPropagation();
            onForgetRecent(path);
          }}
          className="ml-auto shrink-0 rounded-sm p-0.5 opacity-0 transition-opacity group-hover/dropdown-menu-item:opacity-60 group-data-[highlighted]/dropdown-menu-item:opacity-60 group-focus/dropdown-menu-item:opacity-60 hover:bg-accent hover:text-foreground"
        >
          <X className="size-3" />
        </button>
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
        <DropdownMenuContent align="start" className={MENU_WIDTH}>
          <DropdownMenuItem
            className="text-xs"
            icon={<FolderOpen />}
            onSelect={onOpen}
          >
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
                  <DropdownMenuSubContent className={MENU_WIDTH}>
                    {recent(recentFiles.slice(RECENT_SHOWN))}
                  </DropdownMenuSubContent>
                </DropdownMenuSub>
              )}
              <DropdownMenuSeparator />
              <DropdownMenuItem className="text-xs" onSelect={onClearRecents}>
                {t("toolbar.clearRecent")}
              </DropdownMenuItem>
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

      {/* The review fill, as Excel's paint bucket: the bucket itself fills the
          selection with the colour it is holding, and the caret beside it opens
          the palette. Its colours are the app's own — the file is never
          written to. */}
      <div className="flex items-stretch">
        <Button
          size="sm"
          variant="ghost"
          className="h-7 rounded-r-none px-1.5"
          disabled={!hasWorkbook}
          onClick={() => onMark(markColor)}
          title={t("toolbar.fill")}
        >
          <span className="flex flex-col items-center justify-center">
            <PaintBucket className="size-3.5" />
            <span
              aria-hidden
              className="mt-px h-[3px] w-4 rounded-[1px] ring-1 ring-foreground/20 ring-inset"
              style={{ background: markColor }}
            />
          </span>
        </Button>
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button
              size="sm"
              variant="ghost"
              className="h-7 rounded-l-none px-0.5"
              disabled={!hasWorkbook}
              title={t("toolbar.fill")}
            >
              <ChevronDown className="size-3" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start" className="min-w-[13.5rem]">
            <ColorPalette variant="dropdown" current={markColor} onPick={onMark} />
          </DropdownMenuContent>
        </DropdownMenu>
      </div>

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
