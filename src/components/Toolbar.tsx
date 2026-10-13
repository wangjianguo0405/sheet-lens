import {
  Crosshair,
  Download,
  FolderOpen,
  Languages,
  PanelBottom,
  PanelLeft,
  Search,
} from "lucide-react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { LOCALES, setLocale, type Locale } from "@/lib/i18n";

interface ToolbarProps {
  onOpen: () => void;
  onToggleSearch: () => void;
  onToggleSidebar: () => void;
  onToggleSheetTabs: () => void;
  onGoTo: () => void;
  searchOpen: boolean;
  sidebarOpen: boolean;
  sheetTabsOpen: boolean;
  hasWorkbook: boolean;
  /** False for a build that cannot replace itself, which hides the entry. */
  canUpdate: boolean;
  updateAvailable: boolean;
  onShowUpdate: () => void;
}

export function Toolbar({
  onOpen,
  onToggleSearch,
  onToggleSidebar,
  onToggleSheetTabs,
  onGoTo,
  searchOpen,
  sidebarOpen,
  sheetTabsOpen,
  hasWorkbook,
  canUpdate,
  updateAvailable,
  onShowUpdate,
}: ToolbarProps) {
  const { t, i18n } = useTranslation();

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

      <Button
        size="sm"
        variant="ghost"
        className="h-7 gap-1.5 px-2 text-xs"
        onClick={onOpen}
      >
        <FolderOpen className="size-3.5" />
        {t("toolbar.open")}
      </Button>

      <Button
        size="sm"
        variant={searchOpen ? "secondary" : "ghost"}
        className="h-7 gap-1.5 px-2 text-xs"
        disabled={!hasWorkbook}
        onClick={onToggleSearch}
      >
        <Search className="size-3.5" />
        {t("toolbar.search")}
      </Button>

      <Button
        size="sm"
        variant="ghost"
        className="h-7 gap-1.5 px-2 text-xs"
        disabled={!hasWorkbook}
        onClick={onGoTo}
      >
        <Crosshair className="size-3.5" />
        {t("toolbar.goTo")}
      </Button>

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
              <span className="absolute -top-0.5 -right-0.5 size-1.5 rounded-full bg-[#217346]" />
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
