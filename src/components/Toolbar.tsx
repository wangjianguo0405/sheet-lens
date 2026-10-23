import {
  Download,
  FolderOpen,
  Languages,
  X,
} from "lucide-react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import { Hint } from "@/components/Hint";
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
  hasWorkbook: boolean;
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
  hasWorkbook,
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
        <Hint label={t("toolbar.forget", { name: fileName(path) })}>
          <button
            type="button"
            aria-label={t("toolbar.forget", { name: fileName(path) })}
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
        </Hint>
      </DropdownMenuItem>
    ));

  return (
    <div className="flex shrink-0 items-center gap-1 border-b border-border px-2 py-1.5">
      {/* Decorative: the wordmark beside it already names the app, so it repeats
          nothing for a screen reader. `draggable` off, or the webview starts a
          native drag of the bitmap when the bar is dragged. */}
      <img src="/icon.png" alt="" draggable={false} className="mr-1.5 size-5" />
      <span className="mr-2 text-xs font-semibold tracking-wide">SheetLens</span>

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

      <div className="ml-auto flex items-center gap-1">
        {canUpdate && (
          <Hint
            label={
              updateAvailable
                ? t("toolbar.updateAvailable")
                : t("toolbar.checkUpdates")
            }
          >
            <Button
              size="sm"
              variant={updateAvailable ? "secondary" : "ghost"}
              className="relative h-7 px-2"
              onClick={onShowUpdate}
            >
              <Download className="size-3.5" />
              {updateAvailable && (
                <span className="absolute -top-0.5 -right-0.5 size-1.5 rounded-full bg-[#107C41]" />
              )}
            </Button>
          </Hint>
        )}

        {/* Each entry is named in its own language, whatever the UI is in. */}
        <DropdownMenu>
          <Hint label={t("toolbar.language")}>
            <DropdownMenuTrigger asChild>
              <Button size="sm" variant="ghost" className="h-7 px-2">
                <Languages className="size-3.5" />
              </Button>
            </DropdownMenuTrigger>
          </Hint>
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
