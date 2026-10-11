import { Crosshair, Download, FolderOpen, PanelLeft, Search } from "lucide-react";

import { Button } from "@/components/ui/button";

interface ToolbarProps {
  onOpen: () => void;
  onToggleSearch: () => void;
  onToggleSidebar: () => void;
  onGoTo: () => void;
  searchOpen: boolean;
  sidebarOpen: boolean;
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
  onGoTo,
  searchOpen,
  sidebarOpen,
  hasWorkbook,
  canUpdate,
  updateAvailable,
  onShowUpdate,
}: ToolbarProps) {
  return (
    <div className="flex shrink-0 items-center gap-1 border-b border-border px-2 py-1.5">
      <span className="mr-2 text-xs font-semibold tracking-wide">SheetLens</span>

      <Button
        size="sm"
        variant={sidebarOpen ? "secondary" : "ghost"}
        className="h-7 px-2"
        disabled={!hasWorkbook}
        onClick={onToggleSidebar}
        aria-pressed={sidebarOpen}
        title={sidebarOpen ? "Hide the sheet list" : "Show the sheet list"}
      >
        <PanelLeft className="size-3.5" />
      </Button>

      <Button
        size="sm"
        variant="ghost"
        className="h-7 gap-1.5 px-2 text-xs"
        onClick={onOpen}
      >
        <FolderOpen className="size-3.5" />
        Open
      </Button>

      <Button
        size="sm"
        variant={searchOpen ? "secondary" : "ghost"}
        className="h-7 gap-1.5 px-2 text-xs"
        disabled={!hasWorkbook}
        onClick={onToggleSearch}
      >
        <Search className="size-3.5" />
        Search
      </Button>

      <Button
        size="sm"
        variant="ghost"
        className="h-7 gap-1.5 px-2 text-xs"
        disabled={!hasWorkbook}
        onClick={onGoTo}
      >
        <Crosshair className="size-3.5" />
        Go To
      </Button>

      {canUpdate && (
        <Button
          size="sm"
          variant={updateAvailable ? "secondary" : "ghost"}
          className="relative ml-auto h-7 px-2"
          onClick={onShowUpdate}
          title={
            updateAvailable
              ? "A newer version is available"
              : "Check for updates"
          }
        >
          <Download className="size-3.5" />
          {updateAvailable && (
            <span className="absolute -top-0.5 -right-0.5 size-1.5 rounded-full bg-[#217346]" />
          )}
        </Button>
      )}
    </div>
  );
}
